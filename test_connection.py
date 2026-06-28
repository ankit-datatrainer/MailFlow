"""Quick check that the SMTP login works (logs in and logs out, sends nothing)."""
import json
import os
import smtplib
import ssl
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(ROOT, "config.json"), encoding="utf-8-sig") as f:
    cfg = json.load(f)

host = cfg.get("smtp_host", "smtp.hostinger.com")
port = int(cfg.get("smtp_port", 465))
print(f"Connecting to {host}:{port} as {cfg['sender_email']} ...")
ctx = ssl.create_default_context()
try:
    if port == 465:
        s = smtplib.SMTP_SSL(host, port, context=ctx, timeout=20)
    else:
        s = smtplib.SMTP(host, port, timeout=20)
        s.starttls(context=ctx)
    with s:
        s.login(cfg["sender_email"], cfg["password"])
    print("SUCCESS: login worked. You are ready to send.")
except Exception as e:  # noqa: BLE001
    print(f"FAILED: {e}")
    sys.exit(1)
