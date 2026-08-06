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

## Tests

`python test_site.py` — 64 automated checks (signup, login, estimates, admin, Excel, security).
