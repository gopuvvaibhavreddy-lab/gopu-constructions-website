"""Automated test suite for the GVM Infra Developers website. Run: python3 test_site.py"""
import os, sys, tempfile

os.chdir(tempfile.mkdtemp())  # fresh dir so a fresh data.xlsx is created
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import importlib, app as appmod
appmod.EXCEL_FILE = os.path.join(os.getcwd(), "data.xlsx")
appmod.init_excel()
app = appmod.app
app.config["TESTING"] = True

P = F = 0
def t(name, got, want):
    global P, F
    if got == want: P += 1
    else: F += 1; print(f"  FAIL {name}: got {got!r}, want {want!r}")

c = app.test_client()

# ---- public pages & guards
home_html = c.get("/").data.decode()
t("home", c.get("/").status_code, 200)
t("home-has-testimonials-section", 'id="testimonials"' in home_html, True)
t("home-testimonials-count", home_html.count('class="card testimonial"'), 5)
t("404", c.get("/nope").status_code, 404)
t("signup-page", c.get("/signup").status_code, 200)
for p in ["/dashboard", "/renovation", "/build-new", "/appointment"]:
    t(f"guard{p}", c.get(p).status_code, 302)
t("guard-admin", c.get("/admin").status_code, 302)
t("guard-excel", c.get("/admin/download").status_code, 302)

# ---- signup validation
t("signup-short-pw", c.post("/signup", data=dict(name="A", email="a@b.com", phone="1234567890", password="123")).status_code, 302)
t("signup-bad-email", c.post("/signup", data=dict(name="A", email="nope", phone="1234567890", password="123456")).status_code, 302)
t("signup-empty-name", c.post("/signup", data=dict(name=" ", email="a@b.com", phone="1234567890", password="123456")).status_code, 302)
t("signup-missing-fields", c.post("/signup", data={}).status_code, 302)

# ---- good signup / duplicate / login
r = c.post("/signup", data=dict(name="Ravi Kumar", email="ravi@test.com", phone="9876543210", password="secret1"))
t("signup-ok", r.status_code, 302)
t("dashboard-after-signup", c.get("/dashboard").status_code, 200)
c.get("/logout")
t("dup-signup-case-insensitive", c.post("/signup", data=dict(name="R", email="RAVI@test.com", phone="9876543210", password="secret1"), follow_redirects=True).request.path, "/login")
t("login-wrong-pw", b"Invalid" in c.post("/login", data=dict(email="ravi@test.com", password="WRONG")).data, True)
t("login-unknown-user", b"Invalid" in c.post("/login", data=dict(email="ghost@x.com", password="x")).data, True)
t("login-ok", c.post("/login", data=dict(email="Ravi@Test.com", password="secret1")).status_code, 302)

# ---- renovation pages
t("reno-choose", c.get("/renovation").status_code, 200)
t("reno-res", c.get("/renovation?type=residential").status_code, 200)
t("reno-com", c.get("/renovation?type=commercial").status_code, 200)
t("reno-badtype-shows-chooser", c.get("/renovation?type=hack").status_code, 200)
html = c.get("/renovation?type=residential").data.decode()
t("form-keys-no-spaces", 'name="sqft_Full_House"' in html, True)
t("telugu-present", "పడక గది" in html, True)
t("html-not-escaped", "&lt;tr&gt;" not in html and '<input type="number"' in html, True)
t("home-html-not-escaped", "&lt;div" not in c.get("/").data.decode(), True)
t("live-calc-script-present", "function calc()" in html, True)
t("reno-table-wrapped", 'class="table-wrap"' in html, True)

build_html = c.get("/build-new?type=residential").data.decode()
t("build-new-page", c.get("/build-new?type=commercial").status_code, 200)
t("build-new-table-wrapped", 'class="table-wrap"' in build_html, True)

# ---- estimates
r = c.post("/save-estimate", data=dict(service="Renovation", ptype="residential",
    sqft_Bedroom="100", sqft_Full_House="500"), follow_redirects=False)
t("reno-estimate-status", r.status_code, 200)
t("reno-estimate-total-730000", "₹7,30,000" in r.data.decode(), True)  # 100*1050 + 500*1250
t("est-garbage-input", c.post("/save-estimate", data=dict(service="Renovation", ptype="residential", sqft_Bedroom="abc")).status_code, 302)
t("est-negative", c.post("/save-estimate", data=dict(service="Renovation", ptype="residential", sqft_Bedroom="-50")).status_code, 302)
t("est-all-empty", c.post("/save-estimate", data=dict(service="Renovation", ptype="residential")).status_code, 302)
t("est-bad-ptype", c.post("/save-estimate", data=dict(service="Renovation", ptype="HACK", sqft_Bedroom="10")).status_code, 302)
t("est-bad-service", c.post("/save-estimate", data=dict(service="HACK", ptype="residential")).status_code, 302)
t("est-no-fields", c.post("/save-estimate", data={}).status_code, 302)
r = c.post("/save-estimate", data=dict(service="Build New", ptype="residential", sqft_total="1500", tier="Premium"))
t("build-estimate-3900000", "₹39,00,000" in r.data.decode(), True)  # 1500*2600
r = c.post("/save-estimate", data=dict(service="Build New", ptype="commercial", sqft_total="1000", tier="HACKTIER"))
t("build-bad-tier-falls-back-standard", "₹22,00,000" in r.data.decode(), True)  # 1000*2200
t("build-huge-number", c.post("/save-estimate", data=dict(service="Build New", ptype="residential", sqft_total="999999999", tier="Basic")).status_code, 200)

# ---- appointments
t("appt-page", c.get("/appointment").status_code, 200)
r = c.post("/appointment", data=dict(name="Ravi", phone="9876543210", service="Renovation - House",
    date="2026-07-15", time="Morning (9 AM - 12 PM)", location="Warangal", notes="near temple"))
t("appt-ok", r.status_code, 200)
t("appt-whatsapp-link", "wa.me" in r.data.decode(), True)
t("appt-missing-fields", c.post("/appointment", data=dict(name="", phone="", date="", location="")).status_code, 302)
# XSS attempt stored
c.post("/appointment", data=dict(name="<script>alert(1)</script>", phone="1234567890",
    service="Other / Consultation", date="2026-07-20", time="Morning (9 AM - 12 PM)", location="Hyd"))

# ---- admin
c.get("/logout")
t("admin-wrong", b"Wrong" in c.post("/admin-login", data=dict(username="admin", password="nope")).data, True)
t("admin-ok", c.post("/admin-login", data=dict(username="admin", password=appmod.ADMIN_PASSWORD)).status_code, 302)
admin_html = c.get("/admin").data.decode()
t("admin-page", c.get("/admin").status_code, 200)
t("admin-shows-customer", "Ravi Kumar" in admin_html, True)
t("admin-shows-estimate", "₹7,30,000" in admin_html, True)
t("admin-shows-appt", "Warangal" in admin_html, True)
t("xss-escaped", "<script>alert" not in admin_html, True)
t("admin-tables-wrapped", admin_html.count('class="table-wrap"'), 3)
t("excel-download", c.get("/admin/download").status_code, 200)

# ---- excel integrity
from openpyxl import load_workbook
wb = load_workbook(appmod.EXCEL_FILE)
t("excel-sheets", wb.sheetnames, ["Users", "Estimates", "Appointments"])
users = list(wb["Users"].iter_rows(min_row=2, values_only=True))
t("excel-1-user", len(users), 1)
t("excel-pw-hashed", users[0][4].startswith(("scrypt:", "pbkdf2:")), True)
ests = list(wb["Estimates"].iter_rows(min_row=2, values_only=True))
t("excel-estimates-count", len(ests), 4)
appts = list(wb["Appointments"].iter_rows(min_row=2, values_only=True))
t("excel-appts-count", len(appts), 2)
t("excel-appt-status-pending", appts[0][9], "Pending")

# ---- excel self-heal after deletion
os.remove(appmod.EXCEL_FILE)
t("self-heal-admin", c.get("/admin").status_code, 200)
c.get("/logout")
t("self-heal-signup", c.post("/signup", data=dict(name="New", email="new@x.com", phone="1112223334", password="pass123")).status_code, 302)

# ---- inr formatting
t("inr-0", appmod.inr(0), "0")
t("inr-999", appmod.inr(999), "999")
t("inr-lakh", appmod.inr(123456), "1,23,456")
t("inr-crore", appmod.inr(12345678), "1,23,45,678")
t("inr-none", appmod.inr(None), "0")

print(f"\n{'='*40}\nRESULT: {P} passed, {F} failed\n{'='*40}")
sys.exit(1 if F else 0)
