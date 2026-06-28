import imaplib, json, os
ROOT = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(ROOT, "config.json"), encoding="utf-8-sig") as f:
    cfg = json.load(f)
host = cfg.get("imap_host", "imap.hostinger.com")
M = imaplib.IMAP4_SSL(host, 993)
M.login(cfg["sender_email"], cfg["password"])
typ, data = M.list()
for line in data:
    print(line.decode(errors="replace"))
M.logout()
