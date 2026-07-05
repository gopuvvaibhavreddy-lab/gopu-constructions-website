# Gopu Constructions — Meta Ad Kit & Launch Guide

## 1. Website name & domain recommendations

**Recommended brand: Gopu Constructions** (family name = trust in the local market; easy to say in Telugu and English).

| Rank | Domain | Approx. cost/yr | Why |
|------|--------|-----------------|-----|
| 1 | **gopuconstructions.in** | ₹400–700 | Clean, .in ranks well for Indian searches |
| 2 | gopuconstructions.com | ₹900–1,200 | If you want global feel |
| 3 | gopubuilders.in | ₹400–700 | Shorter alternative |
| 4 | buildwithgopu.in | ₹400–700 | Ad-friendly, memorable |

Buy from GoDaddy.com, Hostinger.in, or BigRock.in. Check availability first — buy the .in AND .com if budget allows so no one copies the brand.

---

## 2. Meta (Facebook + Instagram) Ad Templates

### Ad 1 — Lead ad (main ad, use this first)

**Primary text:**
> ఇల్లు కట్టాలనుకుంటున్నారా? రెనోవేషన్ చేయాలనుకుంటున్నారా? 🏗️
>
> Get an INSTANT cost estimate online — room-by-room, at today's Telangana rates. No hidden charges. 20+ years of trusted construction experience across Telangana.
>
> ✅ Renovation from ₹1,050/sqft
> ✅ New construction from ₹1,650/sqft
> ✅ FREE site visit & exact quotation
>
> 📞 Call/WhatsApp: +91 83328 99003

**Headline:** Know Your Construction Cost in 2 Minutes — Free
**Description:** Trusted builder · All of Telangana · Free site visit
**Call-to-action button:** Get Quote
**Destination:** your website URL

### Ad 2 — WhatsApp click ad (cheapest leads in Telangana market)

**Primary text:**
> మీ కల ఇల్లు – మా బాధ్యత 🏡
> House renovation or new construction anywhere in Telangana? Message us on WhatsApp and get a free estimate today. Transparent per-sqft pricing. Free site visit.

**Headline:** Free Construction Estimate on WhatsApp
**CTA button:** Send WhatsApp Message

### Ad image tips
- Use a real photo of a house/building your father built (real photos outperform stock in this market)
- Overlay text: "₹1,650/sqft నుండి" + company name + phone number
- Keep text under ~20% of the image area
- Size: 1080×1080 px (square works on both FB and Instagram)

### Targeting (enter exactly this in Ads Manager)
- Location: Telangana, India (or drop a pin on his district + 50 km)
- Age: 28–60
- Languages: Telugu, English
- Detailed targeting: Home improvement, Real estate, Construction, House plan, Interior design
- Placements: Advantage+ (automatic)
- Budget to start: **₹300–500/day for 7 days**, then keep whichever ad gets cheaper leads

---

## 3. Step-by-step: put the website online

The site is one Python file — any beginner-friendly host works.

**Easiest path — Render.com (free tier):**
1. Create a free account at github.com → New repository → upload `app.py` and `requirements.txt`
2. Create a free account at render.com → **New → Web Service** → connect the GitHub repo
3. Settings: Runtime = Python, Build command = `pip install -r requirements.txt`, Start command = `gunicorn app:app`
4. Click Deploy — in ~3 minutes you get a live URL like `gopuconstructions.onrender.com`
5. Buy your domain (step 1 above) → in Render go to **Settings → Custom Domains** → add `gopuconstructions.in` → copy the DNS records Render shows into your domain provider's DNS page
6. Done — site is live on your own domain with free HTTPS

**Before launch — must-dos:**
- `ADMIN_PASSWORD` and `SECRET_KEY` are read from environment variables, not hardcoded (see `.env.example`). In Render, set them under **Settings → Environment** — don't commit real values to the repo.
- Note: on Render's free tier the Excel file resets when the server restarts — download it from the Admin page regularly, or upgrade to a paid instance (~$7/mo) with a persistent disk

**Admin access:** go to `yourdomain.in/admin-login` → username `admin` + your password. You can see all customers, estimates, appointments, and download the Excel file anytime.

---

## 4. Step-by-step: launch the Meta ad

1. **Create a Facebook Page** — facebook.com → Menu → Pages → Create: name "Gopu Constructions", category "Construction Company", add phone + website + a few photos of completed projects
2. **Create an Instagram account** (optional but free reach) and link it to the Page in Page Settings
3. Go to **adsmanager.facebook.com** → click **+ Create**
4. **Objective:** choose **Leads** (for Ad 1) or **Engagement → WhatsApp** (for Ad 2)
5. **Ad set:** enter the targeting from section 2 above, set budget ₹300–500/day, schedule 7 days
6. **Ad level:** upload your 1080×1080 image, paste the Primary text / Headline / Description from the template, set CTA button, paste your website URL
7. Add payment method (UPI/card works) → click **Publish**
8. Ad goes into review — usually approved within a few hours
9. **Check results daily** in Ads Manager: you want Cost per lead under ₹150–300. Call every lead within 1 hour — speed matters most
10. After 7 days: keep the winning ad, turn off the loser, raise budget slowly (20% at a time)

**Free extras that work well in this market:**
- List the business on Google Maps (business.google.com) — free leads
- Post project photos weekly on the FB page + local Facebook groups
- Put the WhatsApp number everywhere

---

## 5. Pricing used on the website (Telangana, July 2026)

**Renovation — Residential:** Bedroom ₹1,050 · Bathroom ₹2,200 · Kitchen (modular) ₹2,800 · Living/Hall ₹1,000 · Full house ₹1,250 (per sqft)
**Renovation — Commercial:** Office ₹1,400 · Washroom ₹2,500 · Pantry ₹3,000 · Retail ₹1,600 · Full property ₹1,550 (per sqft)
**New construction — Residential:** Basic ₹1,650 · Standard ₹1,950 · Premium ₹2,600 (per sqft)
**New construction — Commercial:** Basic ₹1,850 · Standard ₹2,200 · Premium ₹2,900 (per sqft)

Rates based on July 2026 Hyderabad/Telangana market data. Edit them anytime at the top of `app.py` (`RENOVATION_PRICES` and `BUILD_PRICES`) — ask your father to confirm they match his actual quotes.

Sources: [NoBroker construction cost guide](https://www.nobroker.in/blog/construction-cost-in-hyderabad/), [Infralens Hyderabad 2026 rates](https://infralens.in/prices/hyderabad), [GharKaBudget Hyderabad 2026](https://gharkabudget.com/articles/construction-cost-hyderabad-2026/), [AECORD kitchen renovation 2026](https://aecord.com/blog/kitchen-renovation-cost-india-2026), [AECORD bathroom renovation 2026](https://aecord.com/blog/bathroom-renovation-cost-india-2026)
