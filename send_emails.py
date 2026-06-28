"""
Bulk Email Automation
=====================
Reads recipients from an Excel file and sends a personalized HTML email to each
one using a chosen campaign template, via Gmail SMTP.

Usage:
    python send_emails.py                      # interactive: pick a campaign
    python send_emails.py --campaign welcome   # use campaign by name
    python send_emails.py --excel recipients.xlsx --campaign welcome
    python send_emails.py --dry-run            # preview, send nothing

Excel format (first sheet, with a header row):
    name | email | <any other columns you want as placeholders>

Templates live in the templates/ folder. Each campaign is a folder:
    templates/<campaign>/subject.txt   -> the subject line (supports {name} etc.)
    templates/<campaign>/body.html     -> the HTML body  (supports {name} etc.)
"""

import argparse
import csv
import imaplib
import json
import os
import smtplib
import ssl
import sys
import time
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Formatter

import openpyxl

# Make console output handle emoji/unicode on Windows terminals (cp1252).
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:  # noqa: BLE001
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(ROOT, "templates")
CONFIG_PATH = os.path.join(ROOT, "config.json")
LOG_PATH = os.path.join(ROOT, "sent_log.csv")


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
def load_config():
    if not os.path.exists(CONFIG_PATH):
        sys.exit(
            "ERROR: config.json not found.\n"
            "Copy config.example.json to config.json and fill in your Gmail + App Password.\n"
            "See README.md for how to create a Gmail App Password."
        )
    with open(CONFIG_PATH, encoding="utf-8-sig") as f:
        cfg = json.load(f)
    for key in ("sender_email", "password"):
        if not cfg.get(key) or cfg[key].startswith("YOUR_"):
            sys.exit(f"ERROR: '{key}' is not set in config.json. Please fill it in.")
    cfg.setdefault("sender_name", cfg["sender_email"])
    cfg.setdefault("smtp_host", "smtp.hostinger.com")
    cfg.setdefault("smtp_port", 465)
    cfg.setdefault("delay_seconds", 5)
    cfg.setdefault("max_per_run", 20)
    cfg.setdefault("imap_host", "imap.hostinger.com")
    cfg.setdefault("imap_port", 993)
    cfg.setdefault("sent_folder", "INBOX.Sent")
    cfg.setdefault("save_to_sent", True)
    cfg.setdefault("attachments", [])
    # Validate any configured attachments exist up front.
    for path in cfg["attachments"]:
        full = path if os.path.isabs(path) else os.path.join(ROOT, path)
        if not os.path.exists(full):
            sys.exit(f"ERROR: attachment not found: {full}")
    return cfg


# ---------------------------------------------------------------------------
# Templates / campaigns
# ---------------------------------------------------------------------------
def list_campaigns():
    if not os.path.isdir(TEMPLATES_DIR):
        return []
    out = []
    for name in sorted(os.listdir(TEMPLATES_DIR)):
        d = os.path.join(TEMPLATES_DIR, name)
        if os.path.isdir(d) and os.path.exists(os.path.join(d, "body.html")):
            out.append(name)
    return out


def load_campaign(name):
    d = os.path.join(TEMPLATES_DIR, name)
    body_path = os.path.join(d, "body.html")
    subject_path = os.path.join(d, "subject.txt")
    if not os.path.exists(body_path):
        sys.exit(f"ERROR: campaign '{name}' has no body.html ({body_path})")
    with open(body_path, encoding="utf-8") as f:
        body = f.read()
    subject = "Hello {name}"
    if os.path.exists(subject_path):
        with open(subject_path, encoding="utf-8") as f:
            subject = f.read().strip()
    return subject, body


def placeholders(template):
    """Return the set of {field} names used in a template string."""
    return {fn for _, fn, _, _ in Formatter().parse(template) if fn}


def render(template, row):
    """Fill {field} placeholders; missing fields become empty strings."""
    safe = {k: ("" if v is None else str(v)) for k, v in row.items()}
    for field in placeholders(template):
        safe.setdefault(field, "")
    return template.format(**safe)


# ---------------------------------------------------------------------------
# Recipients
# ---------------------------------------------------------------------------
def read_recipients(excel_path):
    if not os.path.exists(excel_path):
        sys.exit(f"ERROR: Excel file not found: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit("ERROR: Excel file is empty.")
    headers = [str(h).strip().lower() if h is not None else "" for h in rows[0]]
    if "email" not in headers:
        sys.exit("ERROR: Excel must have an 'email' column in the first row.")
    def clean(val):
        """Strip normal + non-breaking whitespace from a cell value."""
        if val is None:
            return ""
        return str(val).replace("\xa0", " ").strip()

    recipients = []
    seen = set()
    for raw in rows[1:]:
        row = {headers[i]: clean(raw[i]) for i in range(len(headers)) if headers[i]}
        # Email: remove all internal whitespace and trailing punctuation.
        email = "".join(row.get("email", "").split()).rstrip(",;")
        if not email or "@" not in email or " " in email:
            continue
        if email.lower() in seen:  # de-duplicate within the file
            continue
        seen.add(email.lower())
        row["email"] = email
        name = " ".join(row.get("name", "").split())
        row["name"] = name if name else email.split("@")[0]
        recipients.append(row)
    return recipients


# ---------------------------------------------------------------------------
# Sent log (dedupe per campaign)
# ---------------------------------------------------------------------------
def load_already_sent(campaign):
    sent = set()
    if not os.path.exists(LOG_PATH):
        return sent
    with open(LOG_PATH, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r.get("campaign") == campaign and r.get("status") == "sent":
                sent.add(r.get("email", "").lower())
    return sent


def log_result(campaign, email, status, detail=""):
    new = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["timestamp", "campaign", "email", "status", "detail"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), campaign, email, status, detail])


# ---------------------------------------------------------------------------
# Sending
# ---------------------------------------------------------------------------
def build_message(cfg, to_email, subject, html_body):
    attachments = cfg.get("attachments", [])
    # Use a 'mixed' container when there are files, so body + files coexist.
    outer = MIMEMultipart("mixed") if attachments else MIMEMultipart("alternative")
    outer["From"] = f"{cfg['sender_name']} <{cfg['sender_email']}>"
    outer["To"] = to_email
    outer["Subject"] = subject

    if attachments:
        body_part = MIMEMultipart("alternative")
        body_part.attach(MIMEText(html_body, "html", "utf-8"))
        outer.attach(body_part)
        for path in attachments:
            full = path if os.path.isabs(path) else os.path.join(ROOT, path)
            with open(full, "rb") as f:
                part = MIMEApplication(f.read(), _subtype="pdf")
            part.add_header("Content-Disposition", "attachment",
                            filename=os.path.basename(full))
            outer.attach(part)
    else:
        outer.attach(MIMEText(html_body, "html", "utf-8"))
    return outer


def open_sent_imap(cfg):
    """Open an IMAP connection for saving copies to the Sent folder, or None."""
    if not cfg.get("save_to_sent"):
        return None
    try:
        imap = imaplib.IMAP4_SSL(cfg["imap_host"], int(cfg["imap_port"]))
        imap.login(cfg["sender_email"], cfg["password"])
        return imap
    except Exception as e:  # noqa: BLE001
        print(f"WARNING: could not open IMAP to save Sent copies: {e}")
        return None


def save_to_sent(imap, cfg, msg):
    """Append a sent message into the Sent folder (so it shows in webmail)."""
    if imap is None:
        return
    try:
        imap.append(cfg["sent_folder"], "(\\Seen)", imaplib.Time2Internaldate(time.time()),
                    msg.as_bytes())
    except Exception as e:  # noqa: BLE001
        print(f"   (warning: could not save copy to Sent: {e})")


def main():
    parser = argparse.ArgumentParser(description="Send personalized bulk emails from an Excel list.")
    parser.add_argument("--excel", default=os.path.join(ROOT, "recipients.xlsx"),
                        help="Path to the Excel file (default: recipients.xlsx)")
    parser.add_argument("--campaign", help="Campaign/template name (folder under templates/)")
    parser.add_argument("--dry-run", action="store_true", help="Preview without sending")
    parser.add_argument("--no-dedupe", action="store_true", help="Send even if already sent before")
    parser.add_argument("--limit", type=int, help="Only send to the first N recipients in the queue")
    parser.add_argument("--only", help="Only send to this exact email address (for testing)")
    parser.add_argument("--to", help="Ad-hoc test send to this address (need not be in the Excel)")
    parser.add_argument("--yes", action="store_true", help="Skip the confirmation prompt")
    args = parser.parse_args()

    cfg = load_config()
    campaigns = list_campaigns()
    if not campaigns:
        sys.exit("ERROR: no campaigns found in templates/. Each campaign is a folder with body.html.")

    campaign = args.campaign
    if not campaign:
        print("Available campaigns:")
        for i, c in enumerate(campaigns, 1):
            print(f"  {i}. {c}")
        choice = input("Pick a campaign (number or name): ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(campaigns):
            campaign = campaigns[int(choice) - 1]
        else:
            campaign = choice
    if campaign not in campaigns:
        sys.exit(f"ERROR: campaign '{campaign}' not found. Available: {', '.join(campaigns)}")

    subject_tpl, body_tpl = load_campaign(campaign)
    recipients = read_recipients(args.excel)

    if args.to:
        recipients = [{"email": args.to.strip(), "name": "there"}]

    if args.only:
        recipients = [r for r in recipients if r["email"].lower() == args.only.lower()]
        if not recipients:
            sys.exit(f"ERROR: --only address '{args.only}' not found in the Excel list.")

    already = set() if args.no_dedupe else load_already_sent(campaign)
    queue = [r for r in recipients if r["email"].lower() not in already]
    skipped = len(recipients) - len(queue)

    if args.limit is not None:
        queue = queue[:args.limit]

    cap = cfg["max_per_run"]
    if len(queue) > cap:
        print(f"NOTE: {len(queue)} recipients in queue; capping this run at max_per_run={cap}.")
        queue = queue[:cap]

    print(f"\nCampaign : {campaign}")
    print(f"Excel    : {args.excel}")
    print(f"Subject  : {render(subject_tpl, recipients[0]) if recipients else subject_tpl}")
    print(f"To send  : {len(queue)}   (skipped already-sent: {skipped})")
    print(f"Mode     : {'DRY RUN (no emails sent)' if args.dry_run else 'LIVE'}\n")

    if not queue:
        print("Nothing to send. Done.")
        return

    if args.dry_run:
        for r in queue:
            print(f"  -> {r['email']:35} | {render(subject_tpl, r)}")
        print("\nDry run complete. No emails were sent.")
        return

    if not args.yes:
        confirm = input(f"Send {len(queue)} emails as {cfg['sender_email']}? [y/N]: ").strip().lower()
        if confirm != "y":
            print("Aborted.")
            return

    context = ssl.create_default_context()
    host, port = cfg["smtp_host"], int(cfg["smtp_port"])
    sent = failed = 0
    if port == 465:
        server = smtplib.SMTP_SSL(host, port, context=context)
    else:
        server = smtplib.SMTP(host, port)
        server.starttls(context=context)
    imap = open_sent_imap(cfg)
    with server:
        server.login(cfg["sender_email"], cfg["password"])
        for i, r in enumerate(queue, 1):
            subject = render(subject_tpl, r)
            body = render(body_tpl, r)
            msg = build_message(cfg, r["email"], subject, body)
            try:
                server.sendmail(cfg["sender_email"], r["email"], msg.as_string())
                sent += 1
                save_to_sent(imap, cfg, msg)
                log_result(campaign, r["email"], "sent")
                print(f"[{i}/{len(queue)}] sent -> {r['email']}")
            except Exception as e:  # noqa: BLE001
                failed += 1
                log_result(campaign, r["email"], "failed", str(e))
                print(f"[{i}/{len(queue)}] FAILED -> {r['email']}: {e}")
            if i < len(queue):
                time.sleep(cfg["delay_seconds"])

    if imap is not None:
        try:
            imap.logout()
        except Exception:  # noqa: BLE001
            pass

    print(f"\nDone. Sent: {sent}, Failed: {failed}. Log: sent_log.csv")


if __name__ == "__main__":
    main()
