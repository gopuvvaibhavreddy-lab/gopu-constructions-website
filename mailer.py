"""
Outgoing email for the GVM Infra site: one SMTP path shared by the security
alerts and the customer welcome message.

Configured entirely through environment variables (see .env.example):

    SMTP_HOST=smtp.gmail.com
    SMTP_PORT=587
    SMTP_USER=<the mailbox that actually sends>
    SMTP_PASS=<Gmail App Password, 16 chars>
    COMPANY_EMAIL=vaibhavreddy@gvminfradevelopers.com   # shown on the site, Reply-To
    SITE_URL=https://www.gvminfradevelopers.com

With SMTP unset nothing is sent and nothing breaks — send() logs and returns
False, so a missing password can never stop a customer signing up.

Note on the From address: Gmail will only send as the account it authenticated
as, so From is always SMTP_USER (with the company name as the display name) and
COMPANY_EMAIL goes in Reply-To. Replies still land in the company inbox.
"""

import os
import html
import logging
import smtplib
import threading
from email.message import EmailMessage
from email.utils import formatdate, formataddr

log = logging.getLogger("gvm.mailer")

SMTP_HOST     = os.environ.get("SMTP_HOST", "smtp.gmail.com").strip()
SMTP_PORT     = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER     = os.environ.get("SMTP_USER", "").strip()
SMTP_PASS     = os.environ.get("SMTP_PASS", "")
COMPANY_EMAIL = os.environ.get("COMPANY_EMAIL", "vaibhavreddy@gvminfradevelopers.com").strip()
COMPANY_NAME  = os.environ.get("SITE_NAME", "GVM Infra Developers")
SITE_URL      = os.environ.get("SITE_URL", "https://www.gvminfradevelopers.com").rstrip("/")
PHONE         = os.environ.get("COMPANY_PHONE", "+91 8332899003")
WHATSAPP      = os.environ.get("COMPANY_WHATSAPP", "918332899003")

# Swapped out by the test suite so tests never touch a real SMTP server.
TRANSPORT = None


def configured():
    return bool(SMTP_USER and SMTP_PASS)


def _deliver(msg):
    if SMTP_PORT == 465:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            s.ehlo()
            s.starttls()
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)


def send(to, subject, text, html_body=None, reply_to=None):
    """Send one email. Returns True if handed to the transport.

    Raises nothing: callers are request handlers, and a mail problem must never
    turn into a 500 for the customer.
    """
    transport = TRANSPORT or _deliver
    if TRANSPORT is None and not configured():
        log.info("email not sent (SMTP not configured): %r to %s", subject, to)
        return False
    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = formataddr((COMPANY_NAME, SMTP_USER or COMPANY_EMAIL))
        msg["To"] = to
        msg["Reply-To"] = reply_to or COMPANY_EMAIL
        msg["Date"] = formatdate(localtime=True)
        msg.set_content(text)
        if html_body:
            msg.add_alternative(html_body, subtype="html")
        transport(msg)
        log.info("email sent: %r to %s", subject, to)
        return True
    except Exception as exc:
        log.error("email failed (%r to %s): %s", subject, to, exc)
        return False


def send_async(*a, **kw):
    """Fire-and-forget, so a slow SMTP server never delays a page load."""
    threading.Thread(target=send, args=a, kwargs=kw, daemon=True).start()


# --------------------------- WELCOME EMAIL -----------------------------
WELCOME_SUBJECT = f"Welcome to {COMPANY_NAME} — your account is ready"


def welcome_text(name):
    return f"""Hello {name},

Welcome to {COMPANY_NAME}. Your account is ready.

మీకు స్వాగతం! మీ ఖాతా సిద్ధంగా ఉంది.

Here is what you can do now:

  * Get an instant cost estimate — room-wise, per-sqft rates for renovation
    or new construction, residential or commercial. Every rate is published
    up front, so there are no hidden line items later.

  * Book a free site visit — we measure on site and turn your estimate into
    an exact quotation at no charge.

Start here: {SITE_URL}

Questions? Just reply to this email, call us on {PHONE}, or message us on
WhatsApp: https://wa.me/{WHATSAPP}

We look forward to building with you.

{COMPANY_NAME}
Building Trust · All of Telangana
{COMPANY_EMAIL} · {PHONE}
"""


def welcome_html(name):
    safe = html.escape(name)
    return f"""\
<!DOCTYPE html>
<html><body style="margin:0;padding:0;background:#EEF3F8;
    font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;color:#16212C;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
         style="background:#EEF3F8;padding:24px 12px;">
    <tr><td align="center">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
             style="max-width:560px;background:#ffffff;border:1px solid #CBD8E4;">
        <tr><td style="background:#0B3B63;padding:24px;text-align:center;">
          <div style="color:#F6F3EA;font-size:20px;font-weight:700;
                      letter-spacing:.06em;text-transform:uppercase;">
            {html.escape(COMPANY_NAME)}</div>
          <div style="color:#9FB8CE;font-size:11px;letter-spacing:.22em;
                      text-transform:uppercase;margin-top:6px;">Building Trust</div>
        </td></tr>
        <tr><td style="padding:28px 26px 8px;">
          <h1 style="margin:0 0 6px;font-size:22px;color:#0B3B63;">Welcome, {safe}!</h1>
          <p style="margin:0 0 4px;font-size:15px;line-height:1.55;">
            Your account is ready.</p>
          <p style="margin:0 0 18px;font-size:14px;color:#5C7086;font-style:italic;">
            మీకు స్వాగతం! మీ ఖాతా సిద్ధంగా ఉంది.</p>
        </td></tr>
        <tr><td style="padding:0 26px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
                 style="border-left:3px solid #A8431F;padding-left:14px;">
            <tr><td style="padding-bottom:14px;">
              <div style="font-weight:700;font-size:15px;">Get an instant estimate</div>
              <div style="font-size:14px;line-height:1.55;color:#5C7086;">
                Room-wise, per-sqft rates for renovation or new construction.
                Every rate is published up front — no hidden line items later.</div>
            </td></tr>
            <tr><td>
              <div style="font-weight:700;font-size:15px;">Book a free site visit</div>
              <div style="font-size:14px;line-height:1.55;color:#5C7086;">
                We measure on site and turn your estimate into an exact
                quotation, at no charge.</div>
            </td></tr>
          </table>
        </td></tr>
        <tr><td style="padding:24px 26px;">
          <a href="{SITE_URL}" style="display:inline-block;background:#A8431F;
             color:#F6F3EA;text-decoration:none;font-weight:700;font-size:15px;
             padding:13px 26px;border-radius:3px;">Get my estimate &rarr;</a>
        </td></tr>
        <tr><td style="padding:0 26px 26px;font-size:14px;line-height:1.6;">
          Questions? Reply to this email, call
          <a href="tel:{PHONE.replace(' ', '')}" style="color:#A8431F;">{PHONE}</a>,
          or <a href="https://wa.me/{WHATSAPP}" style="color:#1e7a4c;">message us on
          WhatsApp</a>.
        </td></tr>
        <tr><td style="background:#0B3B63;padding:18px 26px;text-align:center;
                       color:#9FB8CE;font-size:12px;line-height:1.6;">
          {html.escape(COMPANY_NAME)} · All of Telangana<br>
          <a href="mailto:{COMPANY_EMAIL}" style="color:#C1502B;">{COMPANY_EMAIL}</a>
          · {PHONE}
        </td></tr>
      </table>
    </td></tr>
  </table>
</body></html>"""


def send_welcome(to, name):
    """Greet a customer who just signed up. Never blocks, never raises."""
    display = (name or "there").strip() or "there"
    send_async(to, WELCOME_SUBJECT, welcome_text(display), welcome_html(display))
