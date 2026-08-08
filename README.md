# GVM Infra Developers Website

Construction & renovation estimate website for the Telangana market.
English + Telugu · Customer login/signup · Renovation & Build-New cost calculators
(residential + commercial) · Appointment booking with WhatsApp · Admin dashboard · All data saved to Excel.

## Run it on your computer (to preview)

1. Install Python from python.org (if not installed)
2. Copy `.env.example` to `.env` and fill in `ADMIN_PASSWORD` and `SECRET_KEY` with your own values
3. Open a terminal in this folder and run:
   ```
   pip install -r requirements.txt
   export $(cat .env | xargs)   # loads ADMIN_PASSWORD and SECRET_KEY
   python app.py
   ```
4. Open http://localhost:5000 in your browser

## Logins

- Customers: sign up on the site
- **Admin:** go to `/admin-login` → username `admin`, password: whatever you set `ADMIN_PASSWORD` to in `.env`
  (never commit `.env` — it's gitignored; only `.env.example` with placeholders is committed)

## Where the data lives

Everything (customers, estimates, appointments) is saved to `data.xlsx` in this
folder. Open it in Excel anytime, or download it from the Admin dashboard.

## Change prices / phone number / company name

All of it is at the top of `app.py` in plain text — edit and restart.

## Put it online + run the Meta ad

Full step-by-step instructions in **META-AD-AND-LAUNCH-GUIDE.md**.

## Customer welcome email

When someone signs up, they get a bilingual (English + Telugu) welcome email in
your brand colours: what they can do next, a button to the estimator, your
phone, WhatsApp and company email. Replies go to `COMPANY_EMAIL`.

It needs the same SMTP settings as the security alerts (below). Until those are
set, signup works exactly as before and no mail is attempted. Sending happens on
a background thread and failures are swallowed, so a mail problem can never
break or slow down a signup.

For the best deliverability set `SMTP_USER` to a mailbox on your own domain
(Google Workspace), so the email comes **from** `@gvminfradevelopers.com` rather
than a `gmail.com` address. Gmail only lets you send as the account you log in
as, which is why `From` follows `SMTP_USER` and your company address goes in
`Reply-To`.

## Security & email alerts

The site hardens itself and watches for attacks. Everything is detected and
logged out of the box; to get **emails** when something suspicious happens, set
these in Render → Settings → Environment:

| Variable | Value |
|---|---|
| `ALERT_EMAIL` | the inbox that receives alerts |
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | your Gmail address |
| `SMTP_PASS` | a **Gmail App Password** (16 chars) |

Gmail rejects your normal password. Make an App Password at
Google Account → Security → 2-Step Verification → App passwords.

Then open **Admin → Security → "Send me a test alert"** to confirm delivery.

**You get an email when:** someone tries to guess the admin password, your admin
dashboard is logged into, repeated customer logins fail, the site is being
scanned, one address floods it with requests, a form arrives without a valid
CSRF token, or an injection payload (SQL / XSS / traversal / command) is spotted.

Alerts are capped at 12/hour with a 30-minute cooldown per issue, so a scanner
can't flood your inbox. The Admin → Security table always shows the full recent
history, emailed or not.

**Also hardened:** CSRF tokens on every form, `Secure`/`HttpOnly`/`SameSite`
cookies, a Content-Security-Policy, HSTS, clickjacking and MIME-sniffing
protection, a 256 KB request cap, 15-minute lockout after 5 bad admin logins,
8-character minimum passwords, and formula-injection sanitising so a malicious
name can't run code when you open `database.xlsx` in Excel.

Counters are in-memory: they reset on restart, and they are not a substitute for
a WAF. They are sized for a site this busy.

## Tests

`python test_site.py` — 107 automated checks (signup, login, estimates, admin,
Excel, headers, CSRF, lockouts, injection detection, alerting).
