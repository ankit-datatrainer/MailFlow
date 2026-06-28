"""
Backfill the Sent folder.

The bulk run sent emails over SMTP but did not save copies to the IMAP Sent
folder. This script rebuilds each already-sent message (from sent_log.csv +
recipients) and appends a copy to the Sent folder so they appear in webmail.

Usage:
    python backfill_sent.py --campaign ai-template-one
"""
import argparse
import csv
import imaplib
import os
import time

import send_emails as se  # reuse config/template/render/build_message

ROOT = os.path.dirname(os.path.abspath(__file__))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--campaign", required=True)
    ap.add_argument("--excel", default=os.path.join(ROOT, "recipients.xlsx"))
    args = ap.parse_args()

    cfg = se.load_config()
    subject_tpl, body_tpl = se.load_campaign(args.campaign)

    # name lookup by email (fallback to 'there' for ad-hoc/test addresses)
    names = {}
    for r in se.read_recipients(args.excel):
        names[r["email"].lower()] = r["name"]

    # collect successfully-sent emails for this campaign from the log
    log_path = os.path.join(ROOT, "sent_log.csv")
    targets = []
    seen = set()
    with open(log_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row.get("campaign") == args.campaign and row.get("status") == "sent":
                email = row["email"]
                if email.lower() in seen:
                    continue
                seen.add(email.lower())
                targets.append(email)

    print(f"Backfilling {len(targets)} sent copies into {cfg['sent_folder']} ...")
    imap = imaplib.IMAP4_SSL(cfg["imap_host"], int(cfg["imap_port"]))
    imap.login(cfg["sender_email"], cfg["password"])
    done = fail = 0
    for i, email in enumerate(targets, 1):
        row = {"email": email, "name": names.get(email.lower(), "there")}
        subject = se.render(subject_tpl, row)
        body = se.render(body_tpl, row)
        msg = se.build_message(cfg, email, subject, body)
        try:
            imap.append(cfg["sent_folder"], "(\\Seen)",
                        imaplib.Time2Internaldate(time.time()), msg.as_bytes())
            done += 1
            print(f"[{i}/{len(targets)}] saved -> {email}")
        except Exception as e:  # noqa: BLE001
            fail += 1
            print(f"[{i}/{len(targets)}] FAILED -> {email}: {e}")
    imap.logout()
    print(f"\nDone. Saved to Sent: {done}, Failed: {fail}")


if __name__ == "__main__":
    main()
