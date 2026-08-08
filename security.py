"""
Security hardening + malicious-activity email alerts for the GVM Infra site.
==========================================================================

Two jobs:

1. HARDENING — security headers, CSRF tokens on every form, hardened session
   cookies, request size caps, brute-force lockouts, and a sanitiser that stops
   spreadsheet formula injection reaching database.xlsx.

2. WATCHING — spots suspicious traffic and emails you about it. Detects admin
   brute force, credential stuffing, vulnerability scanning, request floods,
   injection payloads, CSRF failures, and every successful admin login.

Turn alert emails on with these environment variables (all optional — with none
set, everything still gets detected and logged, just not emailed):

    ALERT_EMAIL=you@gmail.com          where alerts are sent
    SMTP_HOST=smtp.gmail.com
    SMTP_PORT=587
    SMTP_USER=you@gmail.com
    SMTP_PASS=<16-char Gmail App Password, NOT your normal password>

Gmail needs an "App Password" (Google Account -> Security -> 2-Step
Verification -> App passwords). A normal password will be rejected.

Note: counters live in memory, so they reset when the service restarts and are
per-worker if you ever run more than one gunicorn worker. That is fine for
alerting on this size of site; it is not a substitute for a WAF.
"""

import os
import re
import threading
import logging
import secrets
from collections import defaultdict, deque
from datetime import datetime, timedelta, timezone

from flask import request, session, abort, g
from werkzeug.middleware.proxy_fix import ProxyFix

import mailer

log = logging.getLogger("gvm.security")

# ----------------------------- CONFIG ---------------------------------
ALERT_EMAIL = os.environ.get("ALERT_EMAIL", "").strip()
SITE_NAME   = os.environ.get("SITE_NAME", "GVM Infra Developers")
# SMTP host/port/user/pass live in mailer.py — one place, shared with customer mail.

# Hardened cookies need HTTPS. Render serves HTTPS; plain `python app.py` does
# not, so the dev server in app.py turns this off for localhost.
SECURE_COOKIES = os.environ.get("SECURE_COOKIES", "1") != "0"

MAX_BODY_BYTES = 256 * 1024      # no legitimate form here is anywhere near this
IST = timezone(timedelta(hours=5, minutes=30))

# thresholds: (how many, within how many seconds)
ADMIN_FAIL_LIMIT   = (5, 900)     # 5 bad admin logins in 15 min -> lock + alert
LOGIN_FAIL_LIMIT   = (10, 900)    # 10 bad customer logins in 15 min -> lock
SIGNUP_LIMIT       = (5, 3600)    # 5 new accounts per hour per IP
NOT_FOUND_LIMIT    = (15, 300)    # 15 x 404 in 5 min -> someone is scanning
FLOOD_LIMIT        = (150, 60)    # 150 requests in 60s from one IP
LOCKOUT_SECONDS    = 900

ALERT_COOLDOWN   = 1800   # don't re-send the same kind for the same IP within 30 min
ALERT_HOURLY_CAP = 12     # hard ceiling on emails per hour, so you can't be mail-bombed

# Paths that only a scanner ever asks for. Seen constantly in this site's logs.
PROBE_PATTERNS = re.compile(
    r"(?i)(^/\.git|^/\.env|/wp-admin|/wp-login|/wp-content|/xmlrpc\.php|/phpmyadmin"
    r"|/\.aws|/\.ssh|/config\.(php|json|yml)$|/backup|/\.well-known/security"
    r"|/vendor/|/cgi-bin/|/actuator|/solr/|/manager/html|\.(sql|bak|old|zip|tar\.gz)$)"
)

# Payload patterns. Deliberately specific: this site takes Telugu text, names
# with apostrophes and free-text addresses, so a bare quote must not trip an alert.
PAYLOAD_PATTERNS = [
    ("sql_injection", re.compile(
        r"(?i)(\bunion\s+select\b|\bselect\b.{0,40}\bfrom\b|\binsert\s+into\b"
        r"|\bdrop\s+table\b|\bupdate\b.{0,30}\bset\b|\bor\s+1\s*=\s*1\b"
        r"|';\s*--|\bsleep\s*\(|\bbenchmark\s*\(|\bwaitfor\s+delay\b)")),
    ("xss_attempt", re.compile(
        r"(?i)(<\s*script|javascript\s*:|\bon(error|load|mouseover)\s*=|<\s*iframe"
        r"|<\s*svg[^>]*\bon|document\.cookie|<\s*img[^>]*\bonerror)")),
    ("path_traversal", re.compile(
        r"(?i)(\.\./\.\./|\.\.\\\.\.\\|/etc/passwd\b|c:\\windows\\)")),
    ("template_injection", re.compile(r"(\{\{.{0,40}\}\}|\$\{.{0,40}\})")),
    ("command_injection", re.compile(
        r"(?i)(;\s*(cat|ls|wget|curl|nc|bash|sh|whoami|id)\s|\|\s*(bash|sh)\b"
        r"|\$\((cat|ls|id|whoami)|\bnc\s+-e\b)")),
]

SEVERITY_ORDER = {"low": 0, "medium": 1, "high": 2, "critical": 3}

# --------------------------- SHARED STATE ------------------------------
_lock = threading.Lock()
_hits = defaultdict(deque)        # (bucket, ip) -> deque[timestamp]
_locks_until = {}                 # (bucket, ip) -> unix ts
_alert_sent_at = {}               # (kind, ip) -> unix ts
_alert_times = deque()            # timestamps of sent alerts, for the hourly cap
_events = deque(maxlen=200)       # recent events, newest last, for the admin page

# Swapped out by the test suite so tests never send real mail.
ALERT_SENDER = None


def _now():
    return datetime.now(tz=timezone.utc).timestamp()


def _ist(ts=None):
    dt = datetime.fromtimestamp(ts or _now(), tz=IST)
    return dt.strftime("%d %b %Y, %I:%M %p IST")


def client_ip():
    """Real client IP. ProxyFix has already resolved X-Forwarded-For."""
    return (request.remote_addr or "unknown") if request else "unknown"


# --------------------------- RATE LIMITING -----------------------------
def _bump(bucket, ip, limit):
    """Record a hit. Returns (count_in_window, tripped)."""
    count, window = limit
    now = _now()
    with _lock:
        dq = _hits[(bucket, ip)]
        dq.append(now)
        while dq and now - dq[0] > window:
            dq.popleft()
        return len(dq), len(dq) >= count


def is_locked(bucket, ip=None):
    """True if this IP is currently locked out of `bucket`."""
    ip = ip or client_ip()
    with _lock:
        until = _locks_until.get((bucket, ip), 0)
        if until > _now():
            return True
        _locks_until.pop((bucket, ip), None)
        return False


def _lock_out(bucket, ip, seconds=LOCKOUT_SECONDS):
    with _lock:
        _locks_until[(bucket, ip)] = _now() + seconds
        _hits.pop((bucket, ip), None)


def _prune():
    """Keep the dicts from growing without bound on a long-running process."""
    now = _now()
    with _lock:
        for key in [k for k, dq in _hits.items() if not dq or now - dq[-1] > 3600]:
            _hits.pop(key, None)
        for key in [k for k, t in _locks_until.items() if t < now]:
            _locks_until.pop(key, None)
        for key in [k for k, t in _alert_sent_at.items() if now - t > 86400]:
            _alert_sent_at.pop(key, None)


# ------------------------------ EVENTS ---------------------------------
def note(kind, detail, severity="medium", ip=None, email=True):
    """Record a security event, and email it if it clears the alert filters."""
    ip = ip or client_ip()
    ev = {
        "at": _now(),
        "kind": kind,
        "detail": str(detail)[:500],
        "severity": severity,
        "ip": ip,
        "path": (request.path if request else "-"),
        "ua": ((request.headers.get("User-Agent", "") or "-")[:200] if request else "-"),
        # Raw header kept alongside the resolved IP: if the proxy chain ever
        # changes and ProxyFix picks the wrong hop, the alert still shows the truth.
        "xff": ((request.headers.get("X-Forwarded-For", "") or "-")[:120] if request else "-"),
    }
    with _lock:
        _events.append(ev)
    log.warning("SECURITY %s [%s] ip=%s path=%s %s",
                kind, severity, ip, ev["path"], ev["detail"])
    if email:
        _maybe_alert(ev)
    return ev


def recent_events(limit=40, min_severity="low"):
    floor = SEVERITY_ORDER.get(min_severity, 0)
    with _lock:
        evs = [e for e in _events if SEVERITY_ORDER.get(e["severity"], 0) >= floor]
    return list(reversed(evs))[:limit]


def _maybe_alert(ev):
    """Cooldown per (kind, ip) plus a global hourly cap, then send."""
    if SEVERITY_ORDER.get(ev["severity"], 0) < SEVERITY_ORDER["medium"]:
        return  # low severity is logged and shown in admin, not emailed
    key = (ev["kind"], ev["ip"])
    now = _now()
    with _lock:
        if now - _alert_sent_at.get(key, 0) < ALERT_COOLDOWN:
            return
        while _alert_times and now - _alert_times[0] > 3600:
            _alert_times.popleft()
        if len(_alert_times) >= ALERT_HOURLY_CAP:
            return
        _alert_sent_at[key] = now
        _alert_times.append(now)
    _send_async(ev)


def _send_async(ev):
    sender = ALERT_SENDER or _smtp_send
    threading.Thread(target=_safe_send, args=(sender, ev), daemon=True).start()


def _safe_send(sender, ev):
    try:
        sender(ev)
    except Exception as exc:                      # never let alerting break the site
        log.error("security alert send failed: %s", exc)


HEADLINES = {
    "admin_bruteforce":     "Someone is trying to guess your admin password",
    "admin_login_failed":   "Failed admin login attempt",
    "admin_login_ok":       "Your admin dashboard was accessed",
    "login_bruteforce":     "Repeated failed customer logins (possible stuffing)",
    "vuln_scan":            "Your site is being scanned for vulnerabilities",
    "probe_path":           "Request for a known attack path",
    "request_flood":        "Unusually high request rate from one address",
    "csrf_failure":         "A form was submitted without a valid CSRF token",
    "signup_abuse":         "Many accounts created from one address",
    "sql_injection":        "SQL injection attempt in a form field",
    "xss_attempt":          "Cross-site scripting attempt in a form field",
    "path_traversal":       "Path traversal attempt",
    "template_injection":   "Template injection attempt",
    "command_injection":    "Command injection attempt",
    "oversized_request":    "Oversized request body rejected",
}

ADVICE = {
    "admin_bruteforce": "That address is now locked out for 15 minutes. If this "
                        "keeps happening, change ADMIN_PASSWORD in Render -> "
                        "Settings -> Environment to something long and random.",
    "admin_login_ok":   "If this was not you, change ADMIN_PASSWORD immediately in "
                        "Render -> Settings -> Environment.",
    "vuln_scan":        "Normal background noise on any public site. Nothing to do "
                        "unless it is sustained — all of these paths return 404.",
}


def _body(ev):
    lines = [
        f"{HEADLINES.get(ev['kind'], ev['kind'])}",
        "",
        f"Site:      {SITE_NAME}",
        f"Severity:  {ev['severity'].upper()}",
        f"When:      {_ist(ev['at'])}",
        f"From IP:   {ev['ip']}" + (f"   (X-Forwarded-For: {ev['xff']})"
                                    if ev.get("xff", "-") not in ("-", ev["ip"]) else ""),
        f"Path:      {ev['path']}",
        f"Browser:   {ev['ua']}",
        f"Detail:    {ev['detail']}",
    ]
    if ev["kind"] in ADVICE:
        lines += ["", "What to do:", ADVICE[ev["kind"]]]
    others = [e for e in recent_events(8) if e is not ev]
    if others:
        lines += ["", "Other recent events:"]
        lines += [f"  {_ist(e['at'])}  {e['severity']:<8} {e['kind']:<20} {e['ip']}"
                  for e in others]
    lines += ["", "-- ", "Automated alert from your website's security monitor.",
              "Full history is on your Admin dashboard under 'Security'."]
    return "\n".join(lines)


def _smtp_send(ev):
    """Hand an alert to the shared mailer (same SMTP settings as customer mail)."""
    if not ALERT_EMAIL:
        log.info("security alert not emailed (ALERT_EMAIL unset): %s", ev["kind"])
        return
    subject = (f"[{SITE_NAME} security] {ev['severity'].upper()}: "
               f"{HEADLINES.get(ev['kind'], ev['kind'])}")
    mailer.send(ALERT_EMAIL, subject, _body(ev))


def alerts_configured():
    return bool(ALERT_EMAIL and mailer.configured())


def send_test_alert():
    """Fire a harmless alert so you can confirm email delivery works."""
    ev = note("admin_login_ok", "Test alert triggered from the Admin dashboard.",
              severity="low", email=False)
    _send_async(ev)
    return alerts_configured()


# ------------------------------- CSRF ----------------------------------
def csrf_token():
    tok = session.get("_csrf")
    if not tok:
        tok = secrets.token_urlsafe(32)
        session["_csrf"] = tok
    return tok


def _check_csrf():
    sent = request.form.get("_csrf", "")
    held = session.get("_csrf", "")
    if not held or not sent or not secrets.compare_digest(sent, held):
        note("csrf_failure", f"POST {request.path} without a valid token",
             severity="medium")
        abort(400)


# ------------------ SPREADSHEET FORMULA INJECTION ----------------------
def spreadsheet_safe(value):
    """Neutralise formula injection before a user string reaches database.xlsx.

    A name like `=HYPERLINK("http://evil","click")` is inert on the website
    (Jinja escapes it) but executes when the owner opens the workbook in Excel.
    Prefixing with an apostrophe makes Excel treat it as text.
    """
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


# --------------------------- REQUEST HOOKS -----------------------------
def _scan_payloads():
    """Look for injection payloads in the query string and form fields."""
    blobs = []
    if request.query_string:
        blobs.append(request.query_string.decode("utf-8", "replace"))
    if request.method == "POST":
        for k, v in request.form.items():
            if k != "_csrf":
                blobs.append(f"{k}={v}")
    joined = "\n".join(blobs)[:8000]
    if not joined:
        return
    for kind, rx in PAYLOAD_PATTERNS:
        m = rx.search(joined)
        if m:
            # Alert only — never block. Output is escaped everywhere, and a
            # false positive must not stop a real customer booking a visit.
            note(kind, f"matched {m.group(0)[:80]!r} in {request.method} {request.path}",
                 severity="high" if kind != "template_injection" else "medium")
            return


def init_app(app):
    """Wire hardening + monitoring into the Flask app."""
    # Render terminates TLS and proxies; without this every client IP is 127.0.0.1
    # and the alerts would be useless.
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

    app.config.update(
        MAX_CONTENT_LENGTH=MAX_BODY_BYTES,
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SECURE=SECURE_COOKIES,
        SESSION_COOKIE_SAMESITE="Lax",
        PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
    )
    app.jinja_env.globals["csrf_token"] = csrf_token

    @app.before_request
    def _guard():
        ip = client_ip()
        g.sec_ip = ip

        if request.method == "POST":
            _check_csrf()

        _, flooding = _bump("req", ip, FLOOD_LIMIT)
        if flooding and not is_locked("flood", ip):
            _lock_out("flood", ip, 300)
            note("request_flood",
                 f"{FLOOD_LIMIT[0]}+ requests in {FLOOD_LIMIT[1]}s", severity="high")
        if is_locked("flood", ip):
            abort(429)

        if PROBE_PATTERNS.search(request.path):
            count, scanning = _bump("probe", ip, NOT_FOUND_LIMIT)
            note("vuln_scan" if scanning else "probe_path",
                 f"requested {request.path} (probe #{count} from this IP)",
                 severity="medium" if scanning else "low")

        _scan_payloads()
        if secrets.randbelow(200) == 0:
            _prune()

    @app.after_request
    def _headers(resp):
        if resp.status_code == 404 and not PROBE_PATTERNS.search(request.path):
            count, scanning = _bump("nf", client_ip(), NOT_FOUND_LIMIT)
            if scanning:
                note("vuln_scan", f"{count} missing-page requests in "
                                  f"{NOT_FOUND_LIMIT[1]}s", severity="medium")
        resp.headers.setdefault("X-Content-Type-Options", "nosniff")
        resp.headers.setdefault("X-Frame-Options", "DENY")
        resp.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
        resp.headers.setdefault("Permissions-Policy",
                                "geolocation=(), microphone=(), camera=(), payment=()")
        resp.headers.setdefault("Cross-Origin-Opener-Policy", "same-origin")
        resp.headers.setdefault(
            "Content-Security-Policy",
            "default-src 'self'; "
            # 'unsafe-inline' is required by the inline <style>/<script> blocks and
            # the two oninput= handlers in the estimate forms. Even so, this blocks
            # any externally hosted script, framing, and form hijacking.
            "script-src 'self' 'unsafe-inline'; "
            "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
            "font-src 'self' https://fonts.gstatic.com; "
            "img-src 'self' data:; "
            "connect-src 'self'; "
            "object-src 'none'; base-uri 'none'; frame-ancestors 'none'; "
            "form-action 'self'")
        if SECURE_COOKIES:
            resp.headers.setdefault("Strict-Transport-Security",
                                    "max-age=31536000; includeSubDomains")
        return resp

    @app.errorhandler(413)
    def _too_big(e):
        note("oversized_request", "request body over the size limit", severity="medium")
        return "Request too large", 413

    return app


# ------------------- HOOKS CALLED FROM app.py ROUTES -------------------
def admin_login_failed(username):
    ip = client_ip()
    count, tripped = _bump("adminfail", ip, ADMIN_FAIL_LIMIT)
    if tripped:
        _lock_out("adminfail", ip)
        note("admin_bruteforce",
             f"{count} failed admin logins; locked out for "
             f"{LOCKOUT_SECONDS // 60} min. Last username tried: {username!r}",
             severity="critical")
    else:
        note("admin_login_failed",
             f"failed admin login {count}/{ADMIN_FAIL_LIMIT[0]} with username {username!r}",
             severity="medium")


def admin_login_ok():
    note("admin_login_ok", "successful admin login", severity="high")


def login_failed(email):
    ip = client_ip()
    count, tripped = _bump("loginfail", ip, LOGIN_FAIL_LIMIT)
    if tripped:
        _lock_out("loginfail", ip, 600)
        note("login_bruteforce",
             f"{count} failed customer logins from this IP; locked out 10 min. "
             f"Last email tried: {email!r}", severity="high")
    else:
        note("login_failed", f"failed login for {email!r} ({count}/{LOGIN_FAIL_LIMIT[0]})",
             severity="low")


def signup_allowed():
    """False if this IP has already created too many accounts this hour.

    SIGNUP_LIMIT is a permitted count, so the Nth signup still goes through and
    the N+1th is refused — unlike the lockout buckets, where hitting the limit
    is itself the failure.
    """
    count, _ = _bump("signup", client_ip(), SIGNUP_LIMIT)
    if count > SIGNUP_LIMIT[0]:
        note("signup_abuse", f"{count} signup attempts from this IP within the hour",
             severity="medium")
        return False
    return True
