# Bulk Email Automation

Send personalized HTML emails to a list of customers from an Excel file, using
Gmail. Supports multiple campaigns/templates, {placeholders}, throttling,
a send log with dedupe, and a dry-run preview.

## 1. One-time setup

### a) Create a Gmail App Password (required)
Gmail will not let a script log in with your normal password. You need a 16-char
"App Password":

1. Turn on 2-Step Verification: https://myaccount.google.com/security
2. Go to https://myaccount.google.com/apppasswords
3. Create a new app password (name it e.g. "Email Automation").
4. Copy the 16-character code (looks like `abcd efgh ijkl mnop`).

### b) Fill in config.json
Open `config.json` and set:

```json
{
  "sender_email": "gilminahn@gmail.com",
  "sender_name": "Your Name or Brand",
  "app_password": "abcdefghijklmnop",   // paste the 16 chars, no spaces
  "delay_seconds": 5,                    // pause between emails (anti-spam)
  "max_per_run": 20                      // safety cap per run
}
```

> `config.json` holds a secret — do not share it or commit it anywhere public.

## 2. Prepare your recipient list

Make an Excel file (`.xlsx`). The first row is the header. You **must** have an
`email` column; `name` is recommended. Any extra columns (e.g. `company`) become
usable placeholders.

| name       | email             | company  |
|------------|-------------------|----------|
| John Doe   | john@example.com  | Acme Inc |
| Jane Smith | jane@example.com  | Globex   |

Save it as `recipients.xlsx` in this folder. A ready-made sample is included
(`recipients_sample.xlsx`); run `python make_sample_excel.py` to regenerate it.

## 3. Templates / campaigns

Each campaign is a folder under `templates/`:

```
templates/
  welcome/
    subject.txt    <- subject line, supports {name}, {company}, etc.
    body.html      <- HTML email body, supports the same placeholders
  promo/
    subject.txt
    body.html
```

Two ready-made campaigns are included: **welcome** and **promo**.
To add your own, copy a folder, rename it, and edit `subject.txt` + `body.html`.
Use `{name}`, `{email}`, or any Excel column name in curly braces.

## 4. Send

Always preview first with `--dry-run` (sends nothing):

```powershell
python send_emails.py --campaign promo --dry-run
```

Then send for real (it will ask for confirmation):

```powershell
python send_emails.py --campaign promo
```

Or run interactively (it lists campaigns and asks which one):

```powershell
python send_emails.py
```

Useful options:

| Option            | Meaning                                              |
|-------------------|------------------------------------------------------|
| `--campaign NAME` | Which template folder to use                          |
| `--excel PATH`    | Use a different Excel file (default `recipients.xlsx`)|
| `--dry-run`       | Preview only, send nothing                            |
| `--no-dedupe`     | Re-send even to people already emailed for this campaign |

## 5. How "20 at a time" works

`max_per_run` in `config.json` caps each run (default 20). If your list has more,
only the first 20 not-yet-sent recipients go out per run. Run it again to send
the next batch — already-sent people are skipped automatically (tracked in
`sent_log.csv`). Use `--no-dedupe` to override.

## Files

- `send_emails.py` – the main program
- `config.json` – your Gmail credentials & settings (keep private)
- `templates/` – campaign templates
- `recipients.xlsx` – your customer list
- `sent_log.csv` – auto-created record of every send (for dedupe & auditing)
- `make_sample_excel.py` – regenerates the sample list

## Notes on deliverability
- Gmail allows ~500 emails/day on free accounts. Keep `delay_seconds` >= 3.
- For large/marketing volumes, a service like SendGrid/Mailgun is better.
- Always include a way to opt out (the promo template has an unsubscribe line).
