"""Fill the official SIH idea-presentation template with the YatraDrishti idea.

Content sources: the team's Gemini working session (idea, modules, stack, business model, judge Q&A)
and the codebase review of the existing multi-tenant WhatsApp AI gateway (head start, risks, action plan).
Citations were checked against arXiv / Crossref; two misattributed references were corrected.

Template rules enforced (see the template's "Important Instructions" slide):
  1. max 6 slides incl. title  2. points/diagrams, no paragraphs  5. pointer texts kept verbatim
  6. upload as PDF (export the .pptx to PDF).

Usage (from SIH_Submission/):  python3 source/build_yatradrishti_ppt.py
"""
from sih_ppt_lib import *  # noqa: F401,F403  (template helpers, palette, paths)

TEAM = {
    "ps_id": "__________",           # Problem Statement ID from the SIH portal
    "ps_title": "Student Innovation",
    "theme": "Travel & Tourism",
    "category": "Software",
    "team_id": "__________",         # Team ID from the SIH portal
    "team_name": "________",         # Team name exactly as registered on the portal
    "logo_text": "YatraDrishti",     # text shown in the team-logo oval on slides 2-6
}

OUT = SRC.parent / "YatraDrishti_SIH_Idea_Presentation.pptx"
TEAL = RGBColor(0x0E, 0x7C, 0x7B)
PURPLE = RGBColor(0x6A, 0x3D, 0x9A)
T_RED = RGBColor(0xFB, 0xEA, 0xEA)
T_PURPLE = RGBColor(0xF1, 0xEB, 0xF8)

prs, S = open_template()

# ---------- Slide 1: title page
fill_title_slide(S[0], TEAM, [
    ("Idea: ", {"bold": True, "color": ORANGE}),
    ("YatraDrishti – Hyper-Personalized AI Ecosystem for Sustainable Tourism & Smart Hospitality", {"bold": True, "color": NAVY}),
])

# ---------- Slide 2: IDEA TITLE / proposed solution
s2 = S[1]
set_title(s2, "YatraDrishti: AI Ecosystem for Sustainable Tourism & Smart Hospitality", 18)
set_logo(s2, TEAM["logo_text"], 11)
hdr = find(s2, "TextBox 8")
tfp = hdr.text_frame.paragraphs
pointers2 = [p.text.strip() for p in tfp[1:] if p.text.strip()]
for p in tfp[1:]:
    p._p.getparent().remove(p._p)
tfp[0].runs[0].font.size = Pt(18)
hdr.left, hdr.top, hdr.width, hdr.height = Inches(0.35), Inches(1.22), Inches(10.2), Inches(0.4)
assert pointers2 == ["Detailed explanation of the proposed solution", "How it addresses the problem",
                     "Innovation and uniqueness of the solution"], pointers2

pointer(s2, 0.4, 1.72, 12.53, pointers2[0], h=0.34, size=12)
text(s2, 0.4, 2.1, 12.53, 0.5, [{"runs": [
    ("YatraDrishti ", {"bold": True, "color": NAVY}),
    ("is a multi-sided AI platform that unites ", {}),
    ("tourists, hoteliers, transport providers and local artisans", {"bold": True, "color": ORANGE}),
    (" in one intelligent ecosystem: a Sustainable Tourism Operating System that links demand for hotels with "
     "predictive crowd balancing for governments.", {})]}], size=12)
mods = [
    ("robot", "1  Tourist App", "GenAI itinerary (budget, pace, weather) · AR monument overlay · Green Travel Pass", BLUE, T_BLUE),
    ("hotel", "2  Hotelier Suite", "XGBoost + Prophet dynamic pricing · zero-commission direct booking", ORANGE, T_ORANGE),
    ("balance", "3  Overtourism Mitigator", "Predicts crowd peaks · nudges to time-shifted or hidden spots with credits", GREEN, T_GREEN),
    ("store", "4  Local Vendor Hub", "One-click, QR-verified onboarding of guides, artisans & homestays", PURPLE, T_PURPLE),
    ("mobile", "5  Offline Edge-AI Guide", "On-device multilingual voice guide, translation & SOS without network", RED, T_RED),
]
cw, gap, y0 = 2.33, 0.22, 2.68
for i, (ic, name, desc, col, tint) in enumerate(mods):
    x = 0.4 + i * (cw + gap)
    shape(s2, MSO_SHAPE.ROUNDED_RECTANGLE, x, y0, cw, 1.3, fill=tint, radius=0.12)
    icon(s2, ic, x + 0.1, y0 + 0.12, 0.5, col)
    text(s2, x + 0.66, y0 + 0.12, cw - 0.72, 0.5, [name], size=11, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    text(s2, x + 0.1, y0 + 0.68, cw - 0.18, 0.6, [desc], size=10, color=INK)
    if i < 4:
        arrow(s2, x + cw + 0.03, y0 + 0.5, 0.16, 0.26, color=LINE)

colw = 6.14
pointer(s2, 0.4, 4.3, colw, pointers2[1], h=0.34, size=12)
fixes = [("Overcrowded tier-1 hotspots", "predictive, time-shifted rerouting"),
         ("15–30% OTA commissions", "zero-commission direct booking engine"),
         ("Guesswork hotel pricing", "AI demand forecast & dynamic rates"),
         ("Invisible artisans & homestays", "verified marketplace inside itineraries"),
         ("Static, unsafe itineraries", "weather-aware plans + offline SOS")]
text(s2, 0.45, 4.72, colw - 0.05, 2.2, [
    {"bullet": True, "runs": [(a, {"bold": True, "color": RED}), ("  →  ", {"color": MUTED}), (b, {})]} for a, b in fixes
], size=12, space_after=6)
pointer(s2, 6.79, 4.3, colw, pointers2[2], h=0.34, size=12)
uniq = [("Tourism OS: ", "tourist actions update hotel & vendor analytics live"),
        ("Green Travel Credits: ", "crowd pressure becomes revenue for nearby spots"),
        ("Enterprise AI for small hotels: ", "5-minute setup via WhatsApp or iCal"),
        ("Offline Edge-AI: ", "safety guide works in no-network regions"),
        ("Not an aggregator or GPT wrapper: ", "real pricing models + spatial analytics")]
text(s2, 6.84, 4.72, colw - 0.05, 2.2, [
    {"bullet": True, "bullet_color": GREEN, "runs": [(a, {"bold": True, "color": NAVY}), (b, {})]} for a, b in uniq
], size=12, space_after=6)

# ---------- Slide 3: TECHNICAL APPROACH
s3 = S[2]
set_logo(s3, TEAM["logo_text"], 11)
pointers3 = pointers_of(s3)
assert pointers3[0].startswith("Technologies to be used") and pointers3[1].startswith("Methodology and process"), pointers3
pointer(s3, 0.4, 1.3, 4.55, pointers3[0], h=0.62, size=12)
table(s3, 0.4, 2.02, 4.55, [1.2, 3.35], [
    ["Layer", "Technology"],
    ["Frontend", "React Native / Flutter · Next.js dashboard"],
    ["Backend", "FastAPI (Python) · Node.js (TypeScript) gateway"],
    ["WhatsApp", [("Multi-tenant Baileys AI gateway ", {}), ("(already built)", {"bold": True, "color": GREEN})]],
    ["AI / ML", "XGBoost + Prophet · PyTorch · LangChain + Llama-3 / Gemini"],
    ["Data", "PostgreSQL + PostGIS · Redis · H3 hexagon grid"],
    ["AR / Edge", "ARCore / ARKit · TensorFlow Lite (on-device)"],
    ["APIs", "OpenStreetMap · OpenWeatherMap · data.tourism.gov.in · Razorpay (test)"],
    ["Cloud", "Docker · Kubernetes · AWS / GCP"],
    ["Hardware", "None: runs on any smartphone"],
], size=10, row_h=0.43)

rx, rw = 5.2, 7.73
pointer(s3, rx, 1.3, rw, pointers3[1], h=0.62, size=12)
inputs = ["Tourist preferences", "Bookings / iCal", "Weather API", "Footfall data", "Local events"]
iw = (rw - 4 * 0.1) / 5
for i, t in enumerate(inputs):
    shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, rx + i * (iw + 0.1), 2.03, iw, 0.36, fill=T_GREY, line=LINE, radius=0.3,
          paras=[{"runs": [(t, {})], "align": PP_ALIGN.CENTER}], size=9.5, color=INK, margin=0.03)
arrow(s3, rx + rw / 2 - 0.13, 2.42, 0.26, 0.24, kind=MSO_SHAPE.DOWN_ARROW)
steps = [("1  Ingest", "FastAPI · PostGIS", BLUE), ("2  Forecast demand", "Prophet + XGBoost", BLUE),
         ("3  Crowd density", "anonymised H3 grid", BLUE), ("4  Peak ahead?", "per site, per hour", NAVY)]
sw = (rw - 3 * 0.28) / 4
for i, (a, b, col) in enumerate(steps):
    x = rx + i * (sw + 0.28)
    shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.7, sw, 0.72, fill=col, radius=0.15,
          paras=[{"runs": [(a, {"bold": True})], "align": PP_ALIGN.CENTER},
                 {"runs": [(b, {"size": 9.5})], "align": PP_ALIGN.CENTER}], size=11, color=WHITE, margin=0.04)
    if i < 3:
        arrow(s3, x + sw + 0.04, 2.93, 0.2, 0.26)
lx = rx + 3 * (sw + 0.28)
arrow(s3, lx + sw / 2 - 0.13, 3.45, 0.26, 0.26, kind=MSO_SHAPE.DOWN_ARROW)
shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, lx - 0.3, 3.76, sw + 0.3, 0.72, fill=T_ORANGE, line=ORANGE, radius=0.15,
      paras=[{"runs": [("YES: congestion", {"bold": True, "color": RED})], "align": PP_ALIGN.CENTER},
             {"runs": [("NO → normal plan & pricing", {"size": 9})], "align": PP_ALIGN.CENTER}], size=11, color=INK)
arrow(s3, lx - 0.62, 3.99, 0.28, 0.26, kind=MSO_SHAPE.LEFT_ARROW, color=ORANGE)
shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, rx, 3.76, lx - 0.66 - rx, 0.72, fill=ORANGE, radius=0.15,
      paras=[{"runs": [("5  Rebalance: time-shifted or secondary-site nudge + Green Travel Credits; hotel rates re-optimised", {"bold": True})],
              "align": PP_ALIGN.CENTER}], size=10.5, color=WHITE, margin=0.08)
arrow(s3, rx + rw / 2 - 1.7, 4.52, 0.26, 0.24, kind=MSO_SHAPE.DOWN_ARROW)
outs = [("Tourist app", "itinerary · AR · credits · SOS"), ("Hotel dashboard + WhatsApp", "rate alerts · direct bookings"),
        ("Govt analytics", "footfall heatmap · policy switch")]
ow = (rw - 2 * 0.15) / 3
for i, (a, b) in enumerate(outs):
    shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, rx + i * (ow + 0.15), 4.8, ow, 0.6, fill=T_GREEN, line=GREEN, radius=0.15,
          paras=[{"runs": [(a, {"bold": True, "color": GREEN})], "align": PP_ALIGN.CENTER},
                 {"runs": [(b, {"size": 9.5})], "align": PP_ALIGN.CENTER}], size=11, color=INK, margin=0.04)
text(s3, rx, 5.52, rw, 0.28, [{"runs": [("36-hour prototype plan (1 pilot circuit, public + synthetic data)", {"bold": True})]}],
     size=10.5, color=NAVY)
plan = ["0–6 h  Data & schema", "6–14 h  Pricing & forecast", "14–22 h  Itinerary AI & crowd", "22–30 h  Apps & AR",
        "30–36 h  Test & demo"]
pw = (rw + 4 * 0.08) / 5
for i, t in enumerate(plan):
    shape(s3, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON, rx + i * (pw - 0.08), 5.84, pw, 0.58,
          fill=[BLUE, BLUE, NAVY, NAVY, GREEN][i],
          paras=[{"runs": [(t, {})], "align": PP_ALIGN.CENTER}], size=9.5, bold=True, color=WHITE, margin=0.14)
text(s3, 0.4, 6.5, 12.5, 0.3, [{"runs": [
    ("Emergency policy switch: ", {"bold": True, "color": RED}),
    ("one admin flag freezes discounts and reroutes tourists during flood / weather alerts (can be shown live to judges).", {})]}],
    size=10, color=MUTED)

# ---------- Slide 4: FEASIBILITY AND VIABILITY
s4 = S[3]
set_logo(s4, TEAM["logo_text"], 11)
pointers4 = pointers_of(s4)
assert pointers4 == ["Analysis of the feasibility of the idea", "Potential challenges and risks",
                     "Strategies for overcoming these challenges"], pointers4
pointer(s4, 0.4, 1.3, 12.53, pointers4[0], h=0.34, size=12)
feas = [
    ("whatsapp", "Technical", GREEN, T_GREEN, ["Multi-tenant WhatsApp AI gateway already working", "Per-tenant AI prompts defined"]),
    ("data", "Data & APIs", NAVY, T_GREY, ["data.tourism.gov.in, OSM, weather, Kaggle hotel data", "Backup: synthetic footfall + 50 hotels"]),
    ("economic", "Economic", BLUE, T_BLUE, ["Hotel SaaS ≈ ₹1,250–4,200 / month", "3–5% fee on local experiences"]),
    ("ops", "Operational", ORANGE, T_ORANGE, ["5-minute hotel setup via WhatsApp / iCal", "Fits existing hotel PMS"]),
]
fw = (12.53 - 3 * 0.2) / 4
for i, (ic, name, col, tint, pts) in enumerate(feas):
    x = 0.4 + i * (fw + 0.2)
    shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.76, fw, 1.38, fill=tint, radius=0.1)
    icon(s4, ic, x + 0.12, 1.84, 0.44, col)
    text(s4, x + 0.64, 1.84, fw - 0.72, 0.44, [f"{name} ✔"], size=13, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    text(s4, x + 0.12, 2.36, fw - 0.2, 0.78, [{"bullet": True, "bullet_color": col, "runs": [(t, {})]} for t in pts],
         size=10.5, space_after=2)
road = [("Phase 1", "Pilot: Golden Triangle, Puducherry, Himachal"), ("Phase 2", "State Tourism Dev. Corporations"),
        ("Phase 3", "Hotel associations & homestay networks"), ("National", "Incredible India integration")]
rw4 = (12.53 + 3 * 0.1) / 4
for i, (a, b) in enumerate(road):
    shape(s4, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON, 0.4 + i * (rw4 - 0.1), 3.25, rw4, 0.56,
          fill=[BLUE, NAVY, ORANGE, GREEN][i],
          paras=[{"runs": [(a + ":  ", {"bold": True}), (b, {})], "align": PP_ALIGN.CENTER}], size=10, color=WHITE, margin=0.25)
tbl_rows = [
    [pointers4[1], pointers4[2]],
    ["WhatsApp session keys exposed in code", "Revoke linked device, purge .baileys_auth from git, secrets in environment variables"],
    ["Sessions lost on every restart", "Database-backed auth store (PostgreSQL / Redis) with caching & safe key serialisation"],
    ["WhatsApp rate limits / number ban", "Outbound queue, rate limiter, randomised delays; official WhatsApp Business API in production"],
    ["Cross-tenant data leakage", "Per-tenant session factory + JWT tenant binding on every endpoint"],
    ["No live telecom crowd data", "Anonymised H3 density counts + Poisson-simulated footfall (DPDP Act 2023)"],
    ["Hotels resist new software", "Zero-friction iCal / webhook sync with existing PMS; rate tips on WhatsApp"],
]
t4 = table(s4, 0.4, 3.93, 12.53, [3.9, 8.63], tbl_rows, size=10.5, row_h=0.39)
t4.cell(0, 0).fill.fore_color.rgb = RED
t4.cell(0, 1).fill.fore_color.rgb = GREEN
for r in range(1, len(tbl_rows)):
    t4.cell(r, 0).text_frame.paragraphs[0].runs[0].font.color.rgb = RED

# ---------- Slide 5: IMPACT AND BENEFITS
s5 = S[4]
set_logo(s5, TEAM["logo_text"], 11)
pointers5 = pointers_of(s5)
assert pointers5 == ["Potential impact on the target audience",
                     "Benefits of the solution (social, economic, environmental, etc.)"], pointers5
pointer(s5, 0.4, 1.3, 12.53, pointers5[0], h=0.34, size=12)
aud = [
    ("traveller", "Tourists", BLUE, T_BLUE, "20–35%", "savings on off-peak routes*",
     ["Budget-aware custom trips", "Eco-travel rewards", "Offline SOS"]),
    ("hotel", "Small hotels", ORANGE, T_ORANGE, "+22–30%", "low-season yield*",
     ["Zero-commission bookings", "AI rate tips on WhatsApp", "Own their guest data"]),
    ("store", "Artisans & locals", PURPLE, T_PURPLE, "+35%", "direct revenue*",
     ["Verified digital badge", "Direct tourist reach", "No middlemen"]),
    ("gov", "Governments", GREEN, T_GREEN, "up to 40%", "less peak congestion*",
     ["Predictive heatmaps", "Emergency policy switch", "Greener destinations"]),
]
aw = (12.53 - 3 * 0.2) / 4
for i, (ic, name, col, tint, big, lab, pts) in enumerate(aud):
    x = 0.4 + i * (aw + 0.2)
    shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.76, aw, 2.0, fill=tint, radius=0.08)
    icon(s5, ic, x + 0.12, 1.86, 0.48, col)
    text(s5, x + 0.68, 1.86, aw - 0.75, 0.48, [name], size=13, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    text(s5, x + 0.12, 2.36, aw - 0.2, 0.38, [big], size=20, bold=True, color=NAVY)
    text(s5, x + 0.12, 2.74, aw - 0.2, 0.25, [lab], size=9.5, color=MUTED)
    text(s5, x + 0.12, 3.02, aw - 0.2, 0.72, [{"bullet": True, "bullet_color": col, "runs": [(t, {})]} for t in pts],
         size=10.5, space_after=2)
pointer(s5, 0.4, 3.9, 12.53, pointers5[1], h=0.34, size=12)
ben = [
    ("social", "Social", BLUE, ["Preserves indigenous crafts & heritage", "Brings tier-2/3 & rural areas into tourism", "Multilingual access for all"]),
    ("economic", "Economic", GREEN, ["Spend shifts from OTAs to local economies", "Hotels keep 100% of direct-booking margin", "New income for guides & homestays"]),
    ("env", "Environmental", RGBColor(0x3A, 0x7D, 0x44), ["Less strain on fragile spots (Ladakh, Shimla, Munnar)", "Rewards off-peak & eco-friendly transit", "Supports SDG 8 & SDG 12"]),
    ("kavach", "Safety & privacy", RED, ["Offline SOS in no-network areas", "DPDP Act 2023: no PII, anonymised H3 counts", "Weather-aware re-planning"]),
]
bw = (12.53 - 3 * 0.2) / 4
for i, (ic, name, col, pts) in enumerate(ben):
    x = 0.4 + i * (bw + 0.2)
    shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, x, 4.36, bw, 2.12, fill=WHITE, line=LINE, radius=0.08)
    icon(s5, ic, x + 0.12, 4.46, 0.46, col)
    text(s5, x + 0.66, 4.46, bw - 0.72, 0.46, [name], size=13, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    text(s5, x + 0.14, 5.0, bw - 0.24, 1.45, [{"bullet": True, "bullet_color": col, "runs": [(t, {})]} for t in pts],
         size=11.5, space_after=5)
text(s5, 0.4, 6.56, 12.5, 0.3, ["*Target outcomes to be validated in the pilot; microservice-based dynamic pricing has shown "
                                "+22% revenue in published research (Barua & Kaiser, 2024)."], size=8.5, color=MUTED)

# ---------- Slide 6: RESEARCH AND REFERENCES
s6 = S[5]
set_logo(s6, TEAM["logo_text"], 11)
pointers6 = pointers_of(s6)
assert pointers6 == ["Details / Links of the reference and research work"], pointers6
pointer(s6, 0.4, 1.3, 12.53, pointers6[0], h=0.34, size=12)


def ref_col(slide, x, w, ic, title, col, items):
    icon(slide, ic, x, 1.8, 0.42, col)
    text(slide, x + 0.52, 1.8, w - 0.55, 0.42, [title], size=13, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    paras = []
    for label, link_text, url in items:
        runs = [(label, {})]
        if url:
            runs += [("  ", {}), (link_text, {"color": BLUE, "link": url})]
        paras.append({"bullet": True, "bullet_color": col, "runs": runs})
    text(slide, x, 2.32, w, 3.9, paras, size=11, space_after=5)


ref_col(s6, 0.4, 6.1, "univ", "Data, platforms & policy", NAVY, [
    ("Ministry of Tourism data portal (tourist footfall)", "data.tourism.gov.in", "https://data.tourism.gov.in"),
    ("OpenStreetMap (heritage sites, transit nodes)", "openstreetmap.org", "https://www.openstreetmap.org"),
    ("OpenWeatherMap API (weather-aware plans)", "openweathermap.org", "https://openweathermap.org"),
    ("Antonio et al. (2019) Hotel booking demand datasets, Data in Brief 22", "doi:10.1016/j.dib.2018.11.126",
     "https://doi.org/10.1016/j.dib.2018.11.126"),
    ("Baileys WhatsApp Web library (our gateway)", "github.com/WhiskeySockets/Baileys", "https://github.com/WhiskeySockets/Baileys"),
    ("Digital Personal Data Protection Act 2023", "meity.gov.in", "https://www.meity.gov.in"),
    ("UN Sustainable Development Goals 8 & 12", "sdgs.un.org/goals", "https://sdgs.un.org/goals"),
])
ref_col(s6, 6.83, 6.1, "flask", "Research papers", ORANGE, [
    ("Barua & Kaiser (2024) Microservices dynamic pricing in travel: +22% revenue, 17% faster pricing", "arXiv:2411.01636",
     "https://arxiv.org/abs/2411.01636"),
    ("Gössling & Mei (2025) AI and sustainable tourism: risks & opportunities for the SDGs, Current Issues in Tourism",
     "doi:10.1080/13683500.2025.2477142", "https://doi.org/10.1080/13683500.2025.2477142"),
    ("Sidiq et al. (2025) Role of AI in transforming smart tourism, JSEIT", "doi:10.31764/jseit.v5i2.30705",
     "https://doi.org/10.31764/jseit.v5i2.30705"),
    ("Chen & Guestrin (2016) XGBoost: a scalable tree boosting system, KDD", "doi:10.1145/2939672.2939785",
     "https://doi.org/10.1145/2939672.2939785"),
    ("Taylor & Letham (2018) Forecasting at scale (Prophet), The American Statistician", "doi:10.1080/00031305.2017.1380080",
     "https://doi.org/10.1080/00031305.2017.1380080"),
])
shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.4, 5.48, 12.53, 0.6, fill=T_GREEN, radius=0.3, margin=0.2,
      paras=[{"runs": [("Our prototype base: ", {"bold": True, "color": GREEN}),
                       ("working multi-tenant WhatsApp AI gateway (Node.js + Baileys) with per-tenant AI prompts; codebase "
                        "security review done and remediation plan defined (see Feasibility).", {})]}], size=10.5, color=INK)
shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.4, 6.22, 12.53, 0.6, fill=T_BLUE, radius=0.3, margin=0.2,
      paras=[{"runs": [("Competitors studied: ", {"bold": True, "color": NAVY}),
                       ("MakeMyTrip, Booking.com, EaseMyTrip (OTAs, 15–30% commission, no crowd balancing) · Mindtrip.ai (AI planner, "
                        "no hotel/PMS link) · Airbnb Experiences · state tourism apps (information only).", {})]}], size=10.5, color=INK)

prs.save(str(OUT))
print("wrote", OUT)
