
"""
Gopu Constructions - Telangana Construction & Renovation Website
================================================================
Features: Sign up / Login, Renovation & Build-New cost estimators
(residential + commercial, room-wise Telangana 2026 rates),
appointment booking (WhatsApp), Admin dashboard, Excel data storage.

Run:  pip install flask openpyxl werkzeug
      export ADMIN_PASSWORD=... SECRET_KEY=...   (see .env.example)
      python app.py
Admin login:  admin / (see ADMIN_PASSWORD env var)
"""

import os
from datetime import datetime
from threading import Lock
from functools import wraps

from flask import (Flask, request, redirect, url_for, session,
                   render_template_string, send_file, flash)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook, load_workbook

# ----------------------------- CONFIG ---------------------------------
COMPANY_NAME    = "Gopu Constructions"
COMPANY_TAGLINE = "Building Telangana, One Home at a Time"
COMPANY_TAGLINE_TE = "మీ కల ఇల్లు – మా బాధ్యత"
PHONE           = "+91 8332899003"
WHATSAPP        = "918332899003"          # digits only, for wa.me links
SERVICE_AREA    = "All of Telangana"
ADMIN_USER      = "admin"
ADMIN_PASSWORD  = os.environ.get("ADMIN_PASSWORD", "dev-admin-change-me")
SECRET_KEY      = os.environ.get("SECRET_KEY", "dev-secret-change-me")
if ADMIN_PASSWORD == "dev-admin-change-me" or SECRET_KEY == "dev-secret-change-me":
    print("WARNING: ADMIN_PASSWORD/SECRET_KEY not set in environment — using insecure dev defaults.")
EXCEL_FILE      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

# ------------------- TELANGANA 2026 PRICES (Rs / sqft) ----------------
# Sources: NoBroker, Infralens, GharKaBudget, AECORD (July 2026 averages)
RENOVATION_PRICES = {
    "residential": {
        "Bedroom":        {"en": "Bedroom",        "te": "పడక గది",      "rate": 1050},
        "Bathroom":       {"en": "Bathroom",       "te": "స్నానాల గది",   "rate": 2200},
        "Kitchen":        {"en": "Kitchen (Modular)", "te": "వంటగది",    "rate": 2800},
        "Living Room":    {"en": "Living / Hall",  "te": "హాలు",          "rate": 1000},
        "Full House":     {"en": "Full House",     "te": "పూర్తి ఇల్లు",   "rate": 1250},
    },
    "commercial": {
        "Office Space":   {"en": "Office Space",   "te": "ఆఫీస్",         "rate": 1400},
        "Washroom":       {"en": "Washroom",       "te": "వాష్‌రూమ్",      "rate": 2500},
        "Pantry/Kitchen": {"en": "Pantry / Kitchen", "te": "పాంట్రీ",     "rate": 3000},
        "Retail/Shop":    {"en": "Retail / Shop Floor", "te": "షాపు",     "rate": 1600},
        "Full Property":  {"en": "Full Property",  "te": "పూర్తి భవనం",   "rate": 1550},
    },
}
BUILD_PRICES = {
    "residential": {
        "Basic":    {"rate": 1650, "desc": "Solid structure, standard fittings"},
        "Standard": {"rate": 1950, "desc": "Vitrified tiles, branded fittings"},
        "Premium":  {"rate": 2600, "desc": "Designer finishes, premium brands"},
    },
    "commercial": {
        "Basic":    {"rate": 1850, "desc": "Grey structure + basic finish"},
        "Standard": {"rate": 2200, "desc": "Ready-to-occupy commercial finish"},
        "Premium":  {"rate": 2900, "desc": "High-end showroom / office grade"},
    },
}

app = Flask(__name__)
app.secret_key = SECRET_KEY
xl_lock = Lock()

# --------------------------- EXCEL BACKEND ----------------------------
SHEETS = {
    "Users":        ["ID", "Name", "Email", "Phone", "PasswordHash", "SignupDate"],
    "Estimates":    ["ID", "UserEmail", "UserName", "Service", "PropertyType",
                     "Details", "TotalSqft", "EstimatedCost", "Date"],
    "Appointments": ["ID", "Name", "Phone", "Email", "Service", "PreferredDate",
                     "PreferredTime", "Location", "Notes", "Status", "BookedOn"],
}

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        wb.remove(wb.active)
        for name, headers in SHEETS.items():
            ws = wb.create_sheet(name)
            ws.append(headers)
        wb.save(EXCEL_FILE)

def xl_append(sheet, row):
    with xl_lock:
        init_excel()  # self-heal if file was deleted
        wb = load_workbook(EXCEL_FILE)
        ws = wb[sheet]
        row = [ws.max_row] + row  # auto ID
        ws.append(row)
        wb.save(EXCEL_FILE)
        return row[0]

def xl_rows(sheet):
    with xl_lock:
        init_excel()  # self-heal if file was deleted
        wb = load_workbook(EXCEL_FILE, read_only=True)
        ws = wb[sheet]
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        wb.close()
        return rows

def safe_float(v):
    """Parse user-supplied number safely; never negative, never crash."""
    try:
        return max(0.0, float(v))
    except (TypeError, ValueError):
        return 0.0

def form_key(key):
    """HTML-safe form field name for a room key."""
    return "sqft_" + key.replace(" ", "_").replace("/", "_")

def find_user(email):
    for r in xl_rows("Users"):
        if r[2] and r[2].lower() == email.lower():
            return {"id": r[0], "name": r[1], "email": r[2], "phone": r[3], "pw": r[4]}
    return None

init_excel()

# ----------------------------- HELPERS --------------------------------
def login_required(f):
    @wraps(f)
    def wrap(*a, **k):
        if "user" not in session:
            return redirect(url_for("login"))
        return f(*a, **k)
    return wrap

def admin_required(f):
    @wraps(f)
    def wrap(*a, **k):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login"))
        return f(*a, **k)
    return wrap

def inr(n):
    """Indian number format: 12,34,567"""
    try:
        n = int(round(float(n)))
    except (TypeError, ValueError):
        return "0"
    s = str(n)
    if len(s) <= 3:
        return s
    last3, rest = s[-3:], s[:-3]
    parts = []
    while len(rest) > 2:
        parts.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        parts.insert(0, rest)
    return ",".join(parts + [last3])

app.jinja_env.filters["inr"] = inr
app.jinja_env.globals["fkey"] = form_key

# ----------------------------- TEMPLATES ------------------------------
BASE = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{ company }} | Construction & Renovation in Telangana</title>
<style>
:root{--navy:#0f2a43;--orange:#f5a623;--orange-d:#e08c00;--bg:#f6f8fa;--card:#fff;--txt:#22303e;--mut:#67788a;}
*{margin:0;padding:0;box-sizing:border-box;font-family:'Segoe UI',system-ui,sans-serif;}
body{background:var(--bg);color:var(--txt);min-height:100vh;display:flex;flex-direction:column;}
header{background:var(--navy);color:#fff;padding:14px 5%;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;}
.logo{font-size:1.3rem;font-weight:800;}.logo span{color:var(--orange);}
nav a{color:#fff;text-decoration:none;margin-left:18px;font-size:.95rem;}
nav a:hover{color:var(--orange);}
.btn{display:inline-block;background:var(--orange);color:var(--navy);font-weight:700;padding:12px 26px;border-radius:8px;text-decoration:none;border:none;cursor:pointer;font-size:1rem;}
.btn:hover{background:var(--orange-d);}
.btn-outline{background:transparent;border:2px solid var(--orange);color:var(--orange);}
.btn-wa{background:#25D366;color:#fff;}
main{flex:1;}
.wrap{max-width:1050px;margin:0 auto;padding:34px 5%;}
h1{font-size:2rem;margin-bottom:8px;} h2{margin:20px 0 12px;} .te{color:var(--mut);font-size:.95rem;}
.card{background:var(--card);border-radius:14px;padding:26px;box-shadow:0 2px 10px rgba(15,42,67,.08);}
.grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:22px;margin-top:22px;}
.opt{border:2px solid transparent;text-align:center;cursor:pointer;transition:.2s;text-decoration:none;color:var(--txt);}
.opt:hover{border-color:var(--orange);transform:translateY(-3px);}
.opt .ic{font-size:3rem;}
table{width:100%;border-collapse:collapse;margin-top:12px;background:#fff;}
th,td{padding:10px 12px;border-bottom:1px solid #e4e9ee;text-align:left;font-size:.93rem;}
th{background:var(--navy);color:#fff;}
tr:hover td{background:#fff7e8;}
input,select,textarea{width:100%;padding:11px;margin:6px 0 14px;border:1px solid #cfd8e0;border-radius:8px;font-size:1rem;}
label{font-weight:600;font-size:.92rem;}
.flash{background:#fff3cd;border:1px solid #ffe08a;padding:10px 14px;border-radius:8px;margin-bottom:14px;}
.total{background:var(--navy);color:#fff;padding:18px;border-radius:10px;font-size:1.25rem;font-weight:700;margin-top:14px;}
.total span{color:var(--orange);}
footer{background:var(--navy);color:#c9d6e2;text-align:center;padding:18px;font-size:.88rem;}
.hero{background:linear-gradient(rgba(15,42,67,.88),rgba(15,42,67,.88)),url('https://images.unsplash.com/photo-1541888946425-d81bb19240f5?w=1200') center/cover;color:#fff;text-align:center;padding:80px 5%;}
.hero h1{font-size:2.6rem;} .hero p{margin:14px 0 26px;font-size:1.15rem;color:#dbe6f0;}
.badge{display:inline-block;background:#e8f0e9;color:#1c7a35;padding:3px 10px;border-radius:20px;font-size:.8rem;font-weight:700;}
.badge.p{background:#fdeaea;color:#b02a2a;}
@media(max-width:600px){.hero h1{font-size:1.8rem;}nav a{margin-left:10px;}}
</style>
</head>
<body>
<header>
  <div class="logo">🏗️ <a href="{{ url_for('home') }}" style="color:inherit;text-decoration:none">{{ company.split()[0] }}<span>{{ company.split()[1:]|join(' ') }}</span></a></div>
  <nav>
    {% if session.get('user') %}
      <a href="{{ url_for('dashboard') }}">Home</a>
      <a href="{{ url_for('appointment') }}">📅 Book Appointment</a>
      <a href="{{ url_for('logout') }}">Logout ({{ session['user']['name'].split()[0] if session['user']['name'].split() else 'User' }})</a>
    {% elif session.get('is_admin') %}
      <a href="{{ url_for('admin') }}">Admin</a><a href="{{ url_for('logout') }}">Logout</a>
    {% else %}
      <a href="{{ url_for('login') }}">Login</a><a href="{{ url_for('signup') }}">Sign Up</a>
    {% endif %}
  </nav>
</header>
<main>
{% with msgs = get_flashed_messages() %}{% if msgs %}<div class="wrap" style="padding-bottom:0">
{% for m in msgs %}<div class="flash">{{ m }}</div>{% endfor %}</div>{% endif %}{% endwith %}
{{ body }}
</main>
<footer>© 2026 {{ company }} · {{ area }} · 📞 {{ phone }} · <a href="https://wa.me/{{ wa }}" style="color:var(--orange)">WhatsApp</a></footer>
</body></html>
"""

def page(body_tpl, **ctx):
    from markupsafe import Markup
    body = Markup(render_template_string(body_tpl, **ctx))  # already escaped by Jinja
    return render_template_string(BASE, body=body, company=COMPANY_NAME,
                                  phone=PHONE, wa=WHATSAPP, area=SERVICE_AREA)

# ------------------------------ ROUTES --------------------------------
@app.route("/")
def home():
    return page("""
<div class="hero">
  <h1>{{ cname }}</h1>
  <p>{{ tagline }} · <b>{{ tagline_te }}</b><br>
  Transparent per-sqft pricing for renovation & new construction across Telangana.</p>
  <a class="btn" href="{{ url_for('signup') }}">Get Instant Estimate →</a>
  <a class="btn btn-wa" href="https://wa.me/{{ wa }}?text=Hi, I want a construction estimate">💬 WhatsApp Us</a>
</div>
<div class="wrap">
  <h2 style="text-align:center">Why choose us? / మమ్మల్ని ఎందుకు ఎంచుకోవాలి?</h2>
  <div class="grid2">
    <div class="card"><div class="ic" style="font-size:2rem">💰</div><h3>Transparent Pricing</h3><p class="te">Room-wise per-sqft rates. No hidden costs. Today's Telangana market prices.</p></div>
    <div class="card"><div class="ic" style="font-size:2rem">👷</div><h3>20+ Years Experience</h3><p class="te">Trusted builder serving all of Telangana — homes, apartments & commercial.</p></div>
    <div class="card"><div class="ic" style="font-size:2rem">📅</div><h3>Free Site Visit</h3><p class="te">Book an appointment online, we visit your site and give an exact quote free.</p></div>
  </div>
</div>""", cname=COMPANY_NAME, tagline=COMPANY_TAGLINE, tagline_te=COMPANY_TAGLINE_TE, wa=WHATSAPP)

@app.route("/signup", methods=["GET", "POST"])
def signup():
    if request.method == "POST":
        name  = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        phone = request.form.get("phone", "").strip()
        pw    = request.form.get("password", "")
        if not name or "@" not in email or len(phone) < 10 or len(pw) < 6:
            flash("Please fill all fields correctly (phone: 10+ digits, password: 6+ characters).")
            return redirect(url_for("signup"))
        if find_user(email):
            flash("An account with this email already exists. Please login.")
            return redirect(url_for("login"))
        xl_append("Users", [name, email, phone, generate_password_hash(pw),
                            datetime.now().strftime("%Y-%m-%d %H:%M")])
        session["user"] = {"name": name, "email": email, "phone": phone}
        return redirect(url_for("dashboard"))
    return page("""
<div class="wrap" style="max-width:460px"><div class="card">
<h1>Sign Up</h1><p class="te">ఖాతా సృష్టించండి — free instant estimates</p>
<form method="post">
<label>Full Name / పేరు</label><input name="name" required>
<label>Email</label><input type="email" name="email" required>
<label>Phone / ఫోన్</label><input name="phone" pattern="[0-9+ -]{10,15}" required>
<label>Password</label><input type="password" name="password" minlength="6" required>
<button class="btn" style="width:100%">Create Account</button>
</form>
<p style="margin-top:12px" class="te">Already have an account? <a href="{{ url_for('login') }}">Login</a></p>
</div></div>""")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        u = find_user(email)
        if u and check_password_hash(u["pw"], request.form.get("password", "")):
            session["user"] = {"name": u["name"], "email": u["email"], "phone": u["phone"]}
            return redirect(url_for("dashboard"))
        flash("Invalid email or password.")
    return page("""
<div class="wrap" style="max-width:460px"><div class="card">
<h1>Login</h1><p class="te">లాగిన్ చేయండి</p>
<form method="post">
<label>Email</label><input type="email" name="email" required>
<label>Password</label><input type="password" name="password" required>
<button class="btn" style="width:100%">Login</button>
</form>
<p style="margin-top:12px" class="te">New here? <a href="{{ url_for('signup') }}">Sign Up</a> ·
<a href="{{ url_for('admin_login') }}" style="color:#aaa">Admin</a></p>
</div></div>""")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))

@app.route("/dashboard")
@login_required
def dashboard():
    return page("""
<div class="wrap">
<h1>Welcome, {{ session['user']['name'] }} 👋</h1>
<p class="te">మీకు ఏ సేవ కావాలి? — What would you like to do?</p>
<div class="grid2">
  <a class="card opt" href="{{ url_for('renovation') }}">
    <div class="ic">🔨</div><h2>Option 1: Renovation</h2>
    <p class="te">పునరుద్ధరణ — Room-wise renovation cost for house or commercial property. Today's Telangana rates per sqft.</p>
  </a>
  <a class="card opt" href="{{ url_for('build_new') }}">
    <div class="ic">🏠</div><h2>Option 2: Build New</h2>
    <p class="te">కొత్త నిర్మాణం — New residential or commercial construction cost estimate per sqft.</p>
  </a>
</div>
<div style="text-align:center;margin-top:30px">
  <a class="btn btn-wa" href="{{ url_for('appointment') }}">📅 Schedule an Appointment / అపాయింట్‌మెంట్ బుక్ చేయండి</a>
</div>
</div>""")

# ------------------------- RENOVATION FLOW -----------------------------
@app.route("/renovation")
@login_required
def renovation():
    ptype = request.args.get("type")
    if ptype not in ("residential", "commercial"):
        return page("""
<div class="wrap">
<h1>🔨 Renovation</h1><p class="te">Is the renovation for a House or a Commercial property?</p>
<div class="grid2">
  <a class="card opt" href="{{ url_for('renovation', type='residential') }}">
    <div class="ic">🏡</div><h2>House / ఇల్లు</h2><p class="te">Bedroom, bathroom, kitchen, hall — room-wise rates.</p></a>
  <a class="card opt" href="{{ url_for('renovation', type='commercial') }}">
    <div class="ic">🏢</div><h2>Commercial / వాణిజ్య</h2><p class="te">Office, shop, washroom, pantry — commercial rates.</p></a>
</div></div>""")
    rooms = RENOVATION_PRICES[ptype]
    return page("""
<div class="wrap">
<h1>🔨 {{ 'House' if ptype=='residential' else 'Commercial' }} Renovation — Telangana Rates (July 2026)</h1>
<p class="te">Enter area in sqft for each room you want to renovate. మీరు పునరుద్ధరించాలనుకునే గదుల విస్తీర్ణం నమోదు చేయండి.</p>
<form method="post" action="{{ url_for('save_estimate') }}" oninput="calc()">
<input type="hidden" name="service" value="Renovation">
<input type="hidden" name="ptype" value="{{ ptype }}">
<table>
<tr><th>Room / గది</th><th>Rate (₹/sqft)</th><th>Area (sqft)</th><th>Cost (₹)</th></tr>
{% for key, r in rooms.items() %}
<tr>
  <td><b>{{ r.en }}</b><br><span class="te">{{ r.te }}</span></td>
  <td>₹{{ r.rate }}</td>
  <td><input type="number" min="0" step="1" name="{{ fkey(key) }}" data-rate="{{ r.rate }}" class="sq" style="margin:0;max-width:130px" placeholder="0"></td>
  <td class="rowcost">₹0</td>
</tr>
{% endfor %}
</table>
<div class="total">Estimated Total / మొత్తం అంచనా: <span id="tot">₹0</span></div>
<p class="te" style="margin-top:8px">* Indicative estimate at current Telangana market rates. Final quote after free site visit.</p>
<button class="btn" style="margin-top:14px">💾 Save Estimate</button>
<a class="btn btn-wa" style="margin-top:14px" href="{{ url_for('appointment') }}">📅 Book Free Site Visit</a>
</form>
</div>
<script>
function calc(){let t=0;document.querySelectorAll('.sq').forEach(i=>{
 const c=(parseFloat(i.value)||0)*parseFloat(i.dataset.rate);t+=c;
 i.closest('tr').querySelector('.rowcost').textContent='₹'+c.toLocaleString('en-IN');});
 document.getElementById('tot').textContent='₹'+t.toLocaleString('en-IN');}
</script>""", ptype=ptype, rooms=rooms)

# -------------------------- BUILD NEW FLOW -----------------------------
@app.route("/build-new")
@login_required
def build_new():
    ptype = request.args.get("type")
    if ptype not in ("residential", "commercial"):
        return page("""
<div class="wrap">
<h1>🏠 Build New</h1><p class="te">What do you want to build? మీరు ఏమి నిర్మించాలనుకుంటున్నారు?</p>
<div class="grid2">
  <a class="card opt" href="{{ url_for('build_new', type='residential') }}">
    <div class="ic">🏡</div><h2>New Residential</h2><p class="te">Independent house, villa, duplex — full construction.</p></a>
  <a class="card opt" href="{{ url_for('build_new', type='commercial') }}">
    <div class="ic">🏢</div><h2>New Commercial</h2><p class="te">Shops, offices, function halls, warehouses.</p></a>
</div></div>""")
    tiers = BUILD_PRICES[ptype]
    rooms = RENOVATION_PRICES[ptype]
    return page("""
<div class="wrap">
<h1>🏠 New {{ 'Residential' if ptype=='residential' else 'Commercial' }} Construction — Telangana Rates (July 2026)</h1>
<form method="post" action="{{ url_for('save_estimate') }}" oninput="calc()">
<input type="hidden" name="service" value="Build New">
<input type="hidden" name="ptype" value="{{ ptype }}">
<div class="card">
<label>Total Built-up Area (sqft) / మొత్తం విస్తీర్ణం</label>
<input type="number" id="area" name="sqft_total" min="0" required placeholder="e.g. 1800">
<label>Quality Package</label>
<select id="tier" name="tier">
{% for name, t in tiers.items() %}
<option value="{{ name }}" data-rate="{{ t.rate }}">{{ name }} — ₹{{ t.rate }}/sqft ({{ t.desc }})</option>
{% endfor %}
</select>
<div class="total">Estimated Total / మొత్తం అంచనా: <span id="tot">₹0</span></div>
</div>
<h2>Room-wise finishing rates (same as renovation rates)</h2>
<table>
<tr><th>Room / గది</th><th>Rate (₹/sqft)</th></tr>
{% for key, r in rooms.items() %}<tr><td>{{ r.en }} <span class="te">({{ r.te }})</span></td><td>₹{{ r.rate }}</td></tr>{% endfor %}
</table>
<p class="te" style="margin-top:8px">* Package covers full structure + finishing. Final quote after free site visit & plan review.</p>
<button class="btn" style="margin-top:14px">💾 Save Estimate</button>
<a class="btn btn-wa" style="margin-top:14px" href="{{ url_for('appointment') }}">📅 Book Free Site Visit</a>
</form>
</div>
<script>
function calc(){const a=parseFloat(document.getElementById('area').value)||0;
 const r=parseFloat(document.getElementById('tier').selectedOptions[0].dataset.rate);
 document.getElementById('tot').textContent='₹'+(a*r).toLocaleString('en-IN');}
</script>""", ptype=ptype, tiers=tiers, rooms=rooms)

# --------------------------- SAVE ESTIMATE -----------------------------
@app.route("/save-estimate", methods=["POST"])
@login_required
def save_estimate():
    service = request.form.get("service", "")
    ptype   = request.form.get("ptype", "")
    if service not in ("Renovation", "Build New") or ptype not in ("residential", "commercial"):
        flash("Something went wrong — please try again.")
        return redirect(url_for("dashboard"))
    details, total_sqft, total_cost = [], 0, 0
    if service == "Renovation":
        for key, r in RENOVATION_PRICES[ptype].items():
            sq = safe_float(request.form.get(form_key(key)))
            if sq > 0:
                cost = sq * r["rate"]
                details.append(f"{r['en']}: {int(sq)} sqft @ ₹{r['rate']} = ₹{inr(cost)}")
                total_sqft += sq; total_cost += cost
    else:
        sq   = safe_float(request.form.get("sqft_total"))
        tier = request.form.get("tier", "Standard")
        if tier not in BUILD_PRICES[ptype]:
            tier = "Standard"
        rate = BUILD_PRICES[ptype][tier]["rate"]
        total_sqft, total_cost = sq, sq * rate
        details.append(f"{tier} package: {int(sq)} sqft @ ₹{rate}/sqft")
    if total_sqft == 0:
        flash("Please enter at least one area in sqft.")
        return redirect(request.referrer or url_for("dashboard"))
    u = session["user"]
    xl_append("Estimates", [u["email"], u["name"], service, ptype.title(),
                            "; ".join(details), int(total_sqft), int(total_cost),
                            datetime.now().strftime("%Y-%m-%d %H:%M")])
    return page("""
<div class="wrap" style="max-width:640px"><div class="card" style="text-align:center">
<div style="font-size:3rem">✅</div>
<h1>Estimate Saved!</h1>
<p class="te">{{ service }} ({{ ptype }}) — {{ sq }} sqft</p>
<div class="total">Estimated Cost: <span>₹{{ cost|inr }}</span></div>
<p style="margin:16px 0" class="te">Our team will review your estimate. Book a free site visit for an exact quote!</p>
<a class="btn btn-wa" href="{{ url_for('appointment') }}">📅 Book Free Site Visit</a>
<a class="btn btn-outline" href="{{ url_for('dashboard') }}">← Back to Home</a>
</div></div>""", service=service, ptype=ptype.title(), sq=int(total_sqft), cost=total_cost)

# --------------------------- APPOINTMENTS ------------------------------
@app.route("/appointment", methods=["GET", "POST"])
@login_required
def appointment():
    u = session["user"]
    if request.method == "POST":
        f = request.form
        name, phone = f.get("name", "").strip(), f.get("phone", "").strip()
        loc, date_ = f.get("location", "").strip(), f.get("date", "").strip()
        if not (name and phone and loc and date_):
            flash("Please fill in name, phone, date and location.")
            return redirect(url_for("appointment"))
        xl_append("Appointments", [name, phone, u["email"], f.get("service", ""),
                                   date_, f.get("time", ""), loc,
                                   f.get("notes", ""), "Pending",
                                   datetime.now().strftime("%Y-%m-%d %H:%M")])
        wa_msg = (f"Hi {COMPANY_NAME}, I booked an appointment. Name: {name}, "
                  f"Service: {f.get('service','')}, Date: {date_} {f.get('time','')}, Location: {loc}")
        return page("""
<div class="wrap" style="max-width:640px"><div class="card" style="text-align:center">
<div style="font-size:3rem">📅</div><h1>Appointment Booked!</h1>
<p class="te">మీ అపాయింట్‌మెంట్ నమోదైంది. We will call you at {{ phone }} to confirm.</p>
<a class="btn btn-wa" href="https://wa.me/{{ wa }}?text={{ msg|urlencode }}">💬 Confirm on WhatsApp</a>
<a class="btn btn-outline" href="{{ url_for('dashboard') }}">← Back to Home</a>
</div></div>""", phone=request.form["phone"], wa=WHATSAPP, msg=wa_msg)
    return page("""
<div class="wrap" style="max-width:560px"><div class="card">
<h1>📅 Schedule an Appointment</h1>
<p class="te">ఉచిత సైట్ విజిట్ — Free site visit & exact quotation</p>
<form method="post">
<label>Name</label><input name="name" value="{{ session['user']['name'] }}" required>
<label>Phone</label><input name="phone" value="{{ session['user']['phone'] }}" required>
<label>Service</label>
<select name="service"><option>Renovation - House</option><option>Renovation - Commercial</option>
<option>New Residential Construction</option><option>New Commercial Construction</option><option>Other / Consultation</option></select>
<label>Preferred Date</label><input type="date" name="date" min="{{ today }}" required>
<label>Preferred Time</label>
<select name="time"><option>Morning (9 AM - 12 PM)</option><option>Afternoon (12 - 4 PM)</option><option>Evening (4 - 7 PM)</option></select>
<label>Site Location (Village/City, District)</label><input name="location" placeholder="e.g. Karimnagar" required>
<label>Notes (optional)</label><textarea name="notes" rows="2"></textarea>
<button class="btn" style="width:100%">Book Appointment / బుక్ చేయండి</button>
</form></div></div>""", today=datetime.now().strftime("%Y-%m-%d"))

# --------------------------- ERROR PAGES -------------------------------
@app.errorhandler(404)
def not_found(e):
    return page("""<div class="wrap" style="text-align:center"><h1>Page not found</h1>
<a class="btn" href="{{ url_for('home') }}">← Go Home</a></div>"""), 404

@app.errorhandler(500)
def server_error(e):
    return page("""<div class="wrap" style="text-align:center"><h1>Something went wrong</h1>
<p class="te">Please try again or call us at """ + PHONE + """</p>
<a class="btn" href="{{ url_for('home') }}">← Go Home</a></div>"""), 500

# ------------------------------ ADMIN ----------------------------------
@app.route("/admin-login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if (request.form["username"] == ADMIN_USER
                and request.form["password"] == ADMIN_PASSWORD):
            session.clear()
            session["is_admin"] = True
            return redirect(url_for("admin"))
        flash("Wrong admin credentials.")
    return page("""
<div class="wrap" style="max-width:420px"><div class="card">
<h1>🔐 Admin Login</h1>
<form method="post">
<label>Username</label><input name="username" required>
<label>Password</label><input type="password" name="password" required>
<button class="btn" style="width:100%">Login</button>
</form></div></div>""")

@app.route("/admin")
@admin_required
def admin():
    users, ests, appts = xl_rows("Users"), xl_rows("Estimates"), xl_rows("Appointments")
    total_value = sum((r[7] or 0) for r in ests)
    return page("""
<div class="wrap">
<h1>🛠️ Admin Dashboard</h1>
<div class="grid2" style="grid-template-columns:repeat(auto-fit,minmax(180px,1fr))">
  <div class="card" style="text-align:center"><h2>{{ users|length }}</h2><p class="te">Customers</p></div>
  <div class="card" style="text-align:center"><h2>{{ ests|length }}</h2><p class="te">Estimates</p></div>
  <div class="card" style="text-align:center"><h2>{{ appts|length }}</h2><p class="te">Appointments</p></div>
  <div class="card" style="text-align:center"><h2>₹{{ tv|inr }}</h2><p class="te">Pipeline Value</p></div>
</div>
<div style="margin:18px 0"><a class="btn" href="{{ url_for('download_excel') }}">⬇️ Download Excel (all data)</a></div>

<h2>📅 Appointments</h2>
<table><tr><th>ID</th><th>Name</th><th>Phone</th><th>Service</th><th>Date</th><th>Time</th><th>Location</th><th>Status</th></tr>
{% for a in appts|reverse %}<tr><td>{{ a[0] }}</td><td>{{ a[1] }}</td><td><a href="tel:{{ a[2] }}">{{ a[2] }}</a></td>
<td>{{ a[4] }}</td><td>{{ a[5] }}</td><td>{{ a[6] }}</td><td>{{ a[7] }}</td>
<td><span class="badge {{ 'p' if a[9]=='Pending' else '' }}">{{ a[9] }}</span></td></tr>
{% else %}<tr><td colspan="8" class="te">No appointments yet.</td></tr>{% endfor %}</table>

<h2>💰 Estimates</h2>
<table><tr><th>ID</th><th>Customer</th><th>Service</th><th>Type</th><th>Details</th><th>Sqft</th><th>Estimate</th><th>Date</th></tr>
{% for e in ests|reverse %}<tr><td>{{ e[0] }}</td><td>{{ e[2] }}<br><span class="te">{{ e[1] }}</span></td>
<td>{{ e[3] }}</td><td>{{ e[4] }}</td><td style="max-width:280px">{{ e[5] }}</td>
<td>{{ e[6] }}</td><td><b>₹{{ e[7]|inr }}</b></td><td>{{ e[8] }}</td></tr>
{% else %}<tr><td colspan="8" class="te">No estimates yet.</td></tr>{% endfor %}</table>

<h2>👥 Customers</h2>
<table><tr><th>ID</th><th>Name</th><th>Email</th><th>Phone</th><th>Signed Up</th></tr>
{% for u in users|reverse %}<tr><td>{{ u[0] }}</td><td>{{ u[1] }}</td><td>{{ u[2] }}</td>
<td><a href="tel:{{ u[3] }}">{{ u[3] }}</a></td><td>{{ u[5] }}</td></tr>
{% else %}<tr><td colspan="5" class="te">No customers yet.</td></tr>{% endfor %}</table>
</div>""", users=users, ests=ests, appts=appts, tv=total_value)

@app.route("/admin/download")
@admin_required
def download_excel():
    return send_file(EXCEL_FILE, as_attachment=True,
                     download_name=f"gopu_constructions_data_{datetime.now():%Y%m%d}.xlsx")

# ------------------------------- MAIN ----------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
