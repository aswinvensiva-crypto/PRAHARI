"""Fill the official SIH idea-presentation template with the PRAHARI idea.

Rules from the template's "Important Instructions" slide that this script enforces:
  1. Maximum 6 slides including the title slide  -> the instructions slide (7) is removed.
  2. Points / diagrams / infographics, no paragraphs.
  5. Only the provided template, idea-detail pointers unchanged -> every pointer text is kept verbatim.
  6. Upload as PDF -> convert the output .pptx to PDF (see README).

Usage (from SIH_Submission/):  python3 source/build_idea_ppt.py
Edit TEAM below before building (Team ID and Team Name come from the SIH portal).
"""
from sih_ppt_lib import *  # noqa: F401,F403  (template helpers, palette, paths)

TEAM = {
    "ps_id": "__________",           # Problem Statement ID from the SIH portal
    "ps_title": "Student Innovation",
    "theme": "Travel & Tourism",
    "category": "Software",
    "team_id": "__________",         # Team ID from the SIH portal
    "team_name": "________",         # Team name exactly as registered on the portal
    "logo_text": "PRAHARI",          # text shown in the team-logo oval on slides 2-6
}

OUT = SRC.parent / "PRAHARI_SIH_Idea_Presentation.pptx"

# ---------------------------------------------------------------- build
prs = Presentation(str(TEMPLATE))

# Rule 1: max six slides -> drop the "Important Instructions" slide (template note says it may be deleted).
sld_ids = prs.slides._sldIdLst
last = sld_ids[-1]
prs.part.drop_rel(last.rId)
sld_ids.remove(last)
S = list(prs.slides)
assert len(S) == 6

# ---------- Slide 1: title page (template pointers kept; values appended)
s1 = S[0]
tb = find(s1, "TextBox 9")
values = [TEAM["ps_id"], TEAM["ps_title"], TEAM["theme"], TEAM["category"], TEAM["team_id"], TEAM["team_name"]]
paras = [p for p in tb.text_frame.paragraphs if p.runs]
for p, val in zip(paras, values):
    label = p.runs[0]
    if label.text.startswith("PS Category"):
        label.text = "PS Category- "
    else:
        label.text = label.text.rstrip() + " "
    label.font.size = Pt(18)
    p.alignment = PP_ALIGN.LEFT
    r = copy.deepcopy(label._r)
    p._p.append(r)
    p.runs[-1].text = val
    p.runs[-1].font.bold = False
    p.runs[-1].font.color.rgb = NAVY
text(s1, 0.36, 2.2, 6.3, 0.7, [
    {"runs": [("Idea: ", {"bold": True, "color": ORANGE}),
              ("PRAHARI – AI Demand Forecasting & Crowd Rebalancing to Boost Hotels, Homestays and Travel in India", {"bold": True, "color": NAVY})]},
], size=13)

# ---------- Slide 2: IDEA TITLE / Proposed solution
s2 = S[1]
set_title(s2, "PRAHARI: AI Demand Forecasting & Crowd Rebalancing", 21)
set_logo(s2, TEAM["logo_text"])
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
    ("PRAHARI ", {"bold": True, "color": NAVY}),
    ("(Predictive Routing & AI for Hotels, Attractions, Regions & Itineraries) predicts where and when tourists will go, checks each "
     "destination's live carrying capacity, and sends surplus travellers to nearby ", {}),
    ("\"twin\" destinations and small hotels with free rooms.", {"bold": True, "color": ORANGE})]}], size=12)
mods = [
    ("pulse", "1  PULSE · Predict", "AI forecast of footfall & hotel occupancy, 1–30 days ahead", BLUE, T_BLUE),
    ("balance", "2  BALANCE · Redistribute", "Live capacity index → suggests \"twin\" places + incentives", ORANGE, T_ORANGE),
    ("setu", "3  SETU · Include", "Homestays list on WhatsApp (22 languages), sold on ONDC, paid by UPI", GREEN, T_GREEN),
    ("kavach", "4  KAVACH · Protect", "Verified stays, fair prices, disaster alerts, SOS to 112 / 1363", RED, RGBColor(0xFB, 0xEA, 0xEA)),
    ("drishti", "5  DRISHTI · Govern", "Live dashboard & what-if simulator for tourism officials", NAVY, T_GREY),
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
fixes = [("Overcrowded hotspots", "7–14 days early warning + overflow moved to twins"),
         ("Empty rooms next door", "surplus demand sent to under-used homestays"),
         ("15–25% OTA commission", "single-digit fees via government-backed ONDC"),
         ("Officials act too late", "live dashboard + permit / campaign simulator"),
         ("Tourist safety gaps", "SACHET alerts, SOS, fair-price registry")]
text(s2, 0.45, 4.72, colw - 0.05, 2.3, [
    {"bullet": True, "runs": [(a, {"bold": True, "color": RED}), ("  →  ", {"color": MUTED}), (b, {})]} for a, b in fixes
], size=12, space_after=6)
pointer(s2, 6.79, 4.3, colw, pointers2[2], h=0.34, size=12)
uniq = [("First closed loop: ", "Predict → Measure → Redistribute (others only sell, inform or restrict)"),
        ("Dynamic Carrying Capacity Index: ", "one-time studies become a daily number"),
        ("No overflow cascade: ", "twins are never pushed above 80% of their capacity"),
        ("Self-learning incentives: ", "Thompson sampling picks the nudge that works"),
        ("Built on India's DPI: ", "ONDC · Bhashini · UPI · NIDHI · SACHET")]
text(s2, 6.84, 4.72, colw - 0.05, 2.3, [
    {"bullet": True, "bullet_color": GREEN, "runs": [(a, {"bold": True, "color": NAVY}), (b, {})]} for a, b in uniq
], size=12, space_after=6)

# ---------- Slide 3: TECHNICAL APPROACH
s3 = S[2]
set_logo(s3, TEAM["logo_text"])
tb3 = find(s3, "TextBox 8")
pointers3 = [p.text.strip() for p in tb3.text_frame.paragraphs if p.text.strip()]
remove(tb3)
assert pointers3[0].startswith("Technologies to be used") and pointers3[1].startswith("Methodology and process"), pointers3

pointer(s3, 0.4, 1.3, 4.55, pointers3[0], h=0.62, size=12)
table(s3, 0.4, 2.02, 4.55, [1.2, 3.35], [
    ["Layer", "Technology"],
    ["Languages", "Python 3.11 · TypeScript · Node.js"],
    ["Frontend", "React PWA (offline) · WhatsApp bot · MapLibre + deck.gl"],
    ["Backend", "FastAPI microservices · Beckn / ONDC adapter"],
    ["Data", "PostgreSQL + PostGIS · TimescaleDB · Kafka · Redis · H3"],
    ["AI / ML", "LightGBM · Temporal Fusion Transformer · OR-Tools · MLflow"],
    ["Integrations", "ONDC · Bhashini · UPI · DigiLocker · IMD · SACHET · 112"],
    ["Infra", "Docker · Kubernetes · MeitY cloud (data in India)"],
    ["Hardware", "None: runs on any smartphone or browser"],
], size=10, row_h=0.47)

rx, rw = 5.2, 7.73
pointer(s3, rx, 1.3, rw, pointers3[1], h=0.62, size=12)
inputs = ["Permits / e-pass", "Hotel bookings", "Weather + SACHET", "Holidays & events", "Search trends"]
iw = (rw - 4 * 0.1) / 5
for i, t in enumerate(inputs):
    shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, rx + i * (iw + 0.1), 2.03, iw, 0.36, fill=T_GREY, line=LINE, radius=0.3,
          paras=[{"runs": [(t, {})], "align": PP_ALIGN.CENTER}], size=9.5, color=INK, margin=0.03)
arrow(s3, rx + rw / 2 - 0.13, 2.42, 0.26, 0.24, kind=MSO_SHAPE.DOWN_ARROW)
steps = [("1  Ingest & clean", "Kafka · PostGIS", BLUE), ("2  PULSE forecast", "LightGBM + TFT", BLUE),
         ("3  D-CCI capacity", "Cifuentes + live factors", BLUE), ("4  Load ratio", "LR = demand ÷ capacity", NAVY)]
sw = (rw - 3 * 0.28) / 4
for i, (a, b, col) in enumerate(steps):
    x = rx + i * (sw + 0.28)
    shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, x, 2.7, sw, 0.72, fill=col, radius=0.15,
          paras=[{"runs": [(a, {"bold": True})], "align": PP_ALIGN.CENTER},
                 {"runs": [(b, {"size": 9.5})], "align": PP_ALIGN.CENTER}], size=11, color=WHITE, margin=0.04)
    if i < 3:
        arrow(s3, x + sw + 0.04, 2.93, 0.2, 0.26)
# decision branch
lx = rx + 3 * (sw + 0.28)
arrow(s3, lx + sw / 2 - 0.13, 3.45, 0.26, 0.26, kind=MSO_SHAPE.DOWN_ARROW)
shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, lx - 0.3, 3.76, sw + 0.3, 0.72, fill=T_ORANGE, line=ORANGE, radius=0.15,
      paras=[{"runs": [("RED (LR > 1)?", {"bold": True, "color": RED})], "align": PP_ALIGN.CENTER},
             {"runs": [("GREEN / AMBER → normal booking", {"size": 9})], "align": PP_ALIGN.CENTER}], size=11, color=INK)
arrow(s3, lx - 0.62, 3.99, 0.28, 0.26, kind=MSO_SHAPE.LEFT_ARROW, color=ORANGE)
shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, rx, 3.76, lx - 0.66 - rx, 0.72, fill=ORANGE, radius=0.15,
      paras=[{"runs": [("5  BALANCE: twin similarity → min-cost-flow allocation (≤ 80% twin capacity) → learned incentive", {"bold": True})],
              "align": PP_ALIGN.CENTER}], size=10.5, color=WHITE, margin=0.08)
arrow(s3, rx + rw / 2 - 1.7, 4.52, 0.26, 0.24, kind=MSO_SHAPE.DOWN_ARROW)
outs = [("Traveller PWA", "forecast · twins · booking · alerts"), ("Host WhatsApp", "listing · demand outlook · price tips"),
        ("DRISHTI dashboard", "heatmap · simulator · reports")]
ow = (rw - 2 * 0.15) / 3
for i, (a, b) in enumerate(outs):
    shape(s3, MSO_SHAPE.ROUNDED_RECTANGLE, rx + i * (ow + 0.15), 4.8, ow, 0.6, fill=T_GREEN, line=GREEN, radius=0.15,
          paras=[{"runs": [(a, {"bold": True, "color": GREEN})], "align": PP_ALIGN.CENTER},
                 {"runs": [(b, {"size": 9.5})], "align": PP_ALIGN.CENTER}], size=11, color=INK, margin=0.04)
text(s3, rx, 5.52, rw, 0.28, [{"runs": [("36-hour prototype plan (Kullu-Manali cluster + 4 twins, public + simulated data)", {"bold": True})]}],
     size=10.5, color=NAVY)
plan = ["0–4 h  Data & setup", "4–12 h  Models & APIs", "12–20 h  Twins & ONDC", "20–28 h  Safety & simulator",
        "28–36 h  Test & demo"]
pw = (rw + 4 * 0.08) / 5
for i, t in enumerate(plan):
    shape(s3, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON, rx + i * (pw - 0.08), 5.84, pw, 0.58,
          fill=[BLUE, BLUE, NAVY, NAVY, GREEN][i],
          paras=[{"runs": [(t, {})], "align": PP_ALIGN.CENTER}], size=9.5, bold=True, color=WHITE, margin=0.14)
text(s3, 0.4, 6.5, 12.5, 0.3, [{"runs": [
    ("D-CCI = PCC × Π(1 − correction factors) × management capacity;  status: ", {}),
    ("GREEN < 0.8", {"bold": True, "color": GREEN}), ("  ·  ", {}), ("AMBER 0.8–1", {"bold": True, "color": AMBER}),
    ("  ·  ", {}), ("RED > 1", {"bold": True, "color": RED}), ("  ·  ", {}), ("BLACK = hazard / closed", {"bold": True, "color": DARK})]}],
    size=10, color=MUTED)

# ---------- Slide 4: FEASIBILITY AND VIABILITY
s4 = S[3]
set_logo(s4, TEAM["logo_text"])
tb4 = find(s4, "TextBox 8")
pointers4 = [p.text.strip() for p in tb4.text_frame.paragraphs if p.text.strip()]
remove(tb4)
assert pointers4 == ["Analysis of the feasibility of the idea", "Potential challenges and risks",
                     "Strategies for overcoming these challenges"], pointers4
pointer(s4, 0.4, 1.3, 12.53, pointers4[0], h=0.34, size=12)
feas = [
    ("tech", "Technical", BLUE, T_BLUE, ["Proven open-source stack", "Well-established ML models", "Prototype fits in 36 h"]),
    ("data", "Data & APIs", NAVY, T_GREY, ["Public data: Tourism Statistics, ASI, IMD, holidays", "ONDC · Bhashini · UPI sandboxes"]),
    ("economic", "Economic", GREEN, T_GREEN, ["District pilot ≈ ₹30–40 lakh / year", "5% of one peak weekend retained ≈ ₹2 Cr*"]),
    ("ops", "Operational", ORANGE, T_ORANGE, ["WhatsApp-first, zero joining fee", "Fits e-pass, NIDHI & portal workflows"]),
]
fw = (12.53 - 3 * 0.2) / 4
for i, (ic, name, col, tint, pts) in enumerate(feas):
    x = 0.4 + i * (fw + 0.2)
    shape(s4, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.78, fw, 1.5, fill=tint, radius=0.1)
    icon(s4, ic, x + 0.12, 1.88, 0.46, col)
    text(s4, x + 0.66, 1.88, fw - 0.75, 0.46, [f"{name} ✔"], size=13, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    text(s4, x + 0.12, 2.42, fw - 0.2, 0.85, [{"bullet": True, "bullet_color": col, "runs": [(t, {})]} for t in pts],
         size=10.5, space_after=2)
road = [("36 hours", "Working prototype"), ("0–9 months", "District pilot: Kullu-Manali, Nilgiris"),
        ("9–18 months", "3–5 states as SaaS tenants"), ("18–36 months", "National tourism DPI")]
rw4 = (12.53 + 3 * 0.1) / 4
for i, (a, b) in enumerate(road):
    shape(s4, MSO_SHAPE.CHEVRON if i else MSO_SHAPE.PENTAGON, 0.4 + i * (rw4 - 0.1), 3.4, rw4, 0.58,
          fill=[BLUE, NAVY, ORANGE, GREEN][i],
          paras=[{"runs": [(a + ":  ", {"bold": True}), (b, {})], "align": PP_ALIGN.CENTER}], size=10.5, color=WHITE, margin=0.25)
tbl_rows = [
    [pointers4[1], pointers4[2]],
    ["Sparse daily footfall data", "Proxy signals (permits, bookings, search) + cold-start transfer learning + state data MoUs"],
    ["Low digital adoption by small hosts", "WhatsApp & voice onboarding in local language, zero fee, camps with SHGs & district offices"],
    ["Travellers ignore suggestions", "Learned incentives, honest delay estimates, twin offers at the e-pass stage"],
    ["Crowd just shifts to a fragile village", "Hard cap at 80% of twin capacity + community consent before listing"],
    ["Govt. API dependency & privacy", "Modular adapters with cached fallbacks; aggregate-only data, DPDP Act 2023 compliant"],
]
t4 = table(s4, 0.4, 4.12, 12.53, [4.1, 8.43], tbl_rows, size=11, row_h=0.43)
for c in range(2):
    cell = t4.cell(0, c)
    cell.fill.fore_color.rgb = RED if c == 0 else GREEN
for r in range(1, len(tbl_rows)):
    t4.cell(r, 0).text_frame.paragraphs[0].runs[0].font.color.rgb = RED
text(s4, 0.4, 6.72, 12.5, 0.22, ["*Indicative estimate: 1 lakh peak-weekend visitors × ₹4,000 spend; to be validated in the pilot."],
     size=8.5, color=MUTED)

# ---------- Slide 5: IMPACT AND BENEFITS
s5 = S[4]
set_logo(s5, TEAM["logo_text"])
tb5 = find(s5, "TextBox 8")
pointers5 = [p.text.strip() for p in tb5.text_frame.paragraphs if p.text.strip()]
remove(tb5)
assert pointers5 == ["Potential impact on the target audience",
                     "Benefits of the solution (social, economic, environmental, etc.)"], pointers5
pointer(s5, 0.4, 1.3, 12.53, pointers5[0], h=0.34, size=12)
aud = [
    ("traveller", "Travellers", BLUE, T_BLUE, "7–14 days", "early warning of crowds",
     ["Less traffic & queues", "Verified stays, fair prices", "Alerts + one-tap SOS"]),
    ("hotel", "Hotels & homestays", ORANGE, T_ORANGE, "+10–15 pts", "off-season occupancy*",
     ["Listed in < 10 minutes", "Single-digit fees vs 15–25%", "Weekly demand & price tips"]),
    ("gov", "Tourism departments", GREEN, T_GREEN, "~10 min", "weekly report (was ~2 days)",
     ["Plan police, buses, sanitation early", "What-if permit simulator", "Track scheme outcomes"]),
]
aw = (12.53 - 2 * 0.25) / 3
for i, (ic, name, col, tint, big, lab, pts) in enumerate(aud):
    x = 0.4 + i * (aw + 0.25)
    shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, x, 1.76, aw, 1.95, fill=tint, radius=0.08)
    icon(s5, ic, x + 0.15, 1.88, 0.55, col)
    text(s5, x + 0.8, 1.88, 1.9, 0.55, [name], size=13, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    text(s5, x + 0.15, 2.5, 1.95, 0.5, [big], size=24, bold=True, color=NAVY)
    text(s5, x + 0.15, 3.0, 1.95, 0.6, [lab], size=10, color=MUTED)
    text(s5, x + 2.15, 2.52, aw - 2.25, 1.15, [{"bullet": True, "bullet_color": col, "runs": [(t, {})]} for t in pts],
         size=10.5, space_after=3)
pointer(s5, 0.4, 3.86, 12.53, pointers5[1], h=0.34, size=12)
ben = [
    ("social", "Social", BLUE, ["Income for rural & women-led homestays", "Digital inclusion in 22 languages", "Less resident resentment"]),
    ("economic", "Economic", GREEN, ["Visitors redirected, not turned away", "Hosts save ₹200–400 per ₹2,000 night", "More local spend & GST"]),
    ("env", "Environmental", RGBColor(0x3A, 0x7D, 0x44), ["Fragile hill & coastal spots stay under capacity", "Less traffic idling & waste peaks", "Longer destination life cycle"]),
    ("kavach", "Safety & governance", RED, ["Disaster alerts in < 5 minutes", "Evidence-based planning", "Supports SDG 8.9, 11, 12.b"]),
]
bw = (12.53 - 3 * 0.2) / 4
for i, (ic, name, col, pts) in enumerate(ben):
    x = 0.4 + i * (bw + 0.2)
    shape(s5, MSO_SHAPE.ROUNDED_RECTANGLE, x, 4.32, bw, 2.2, fill=WHITE, line=LINE, radius=0.08)
    icon(s5, ic, x + 0.12, 4.42, 0.46, col)
    text(s5, x + 0.66, 4.42, bw - 0.72, 0.46, [name], size=13, bold=True, color=col, anchor=MSO_ANCHOR.MIDDLE)
    text(s5, x + 0.14, 5.0, bw - 0.24, 1.5, [{"bullet": True, "bullet_color": col, "runs": [(t, {})]} for t in pts],
         size=12, space_after=6)
text(s5, 0.4, 6.6, 12.5, 0.25, ["*Pilot targets / indicative estimates, to be validated in a district pilot."], size=8.5, color=MUTED)

# ---------- Slide 6: RESEARCH AND REFERENCES
s6 = S[5]
set_logo(s6, TEAM["logo_text"])
tb6 = find(s6, "TextBox 8")
pointers6 = [p.text.strip() for p in tb6.text_frame.paragraphs if p.text.strip()]
remove(tb6)
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


ref_col(s6, 0.4, 6.1, "univ", "Government data & policy", NAVY, [
    ("Ministry of Tourism: India Tourism Statistics 2024 & Tourism Satellite Account", "tourism.gov.in", "https://tourism.gov.in"),
    ("Swadesh Darshan 2.0, SASCI iconic destinations (PIB releases)", "pib.gov.in", "https://pib.gov.in"),
    ("Union Budget 2025-26: 50 destinations, MUDRA loans for homestays", "indiabudget.gov.in", "https://www.indiabudget.gov.in"),
    ("Madras HC e-pass for Ooty & Kodaikanal (2024); NGT Rohtang vehicle caps", "", None),
    ("UNWTO (2018) 'Overtourism?' Understanding & managing urban tourism growth", "e-unwto.org", "https://www.e-unwto.org"),
    ("Digital Personal Data Protection Act 2023", "meity.gov.in", "https://www.meity.gov.in"),
    ("ONDC open network (Beckn protocol)", "ondc.org", "https://ondc.org"),
    ("Bhashini language platform (MeitY)", "bhashini.gov.in", "https://bhashini.gov.in"),
    ("NDMA SACHET disaster alerts", "sachet.ndma.gov.in", "https://sachet.ndma.gov.in"),
    ("NIDHI hospitality database (Ministry of Tourism)", "nidhi.tourism.gov.in", "https://nidhi.tourism.gov.in"),
])
ref_col(s6, 6.83, 6.1, "flask", "Research papers", ORANGE, [
    ("Song & Li (2008) Tourism demand forecasting review, Tourism Management 29(2)", "doi:10.1016/j.tourman.2007.07.016",
     "https://doi.org/10.1016/j.tourman.2007.07.016"),
    ("Bangwayo-Skeete & Skeete (2015) Google data improves arrival forecasts, Tourism Management 46", "", None),
    ("Law et al. (2019) Deep learning for tourism demand, Annals of Tourism Research 75", "", None),
    ("Lim et al. (2021) Temporal Fusion Transformers, Int. J. of Forecasting", "arXiv:1912.09363", "https://arxiv.org/abs/1912.09363"),
    ("Wickramasuriya et al. (2019) MinT forecast reconciliation, JASA", "doi:10.1080/01621459.2018.1448825",
     "https://doi.org/10.1080/01621459.2018.1448825"),
    ("Cifuentes (1992) Tourism carrying-capacity method, CATIE", "", None),
    ("Butler (1980) Tourist Area Life Cycle, Canadian Geographer", "doi:10.1111/j.1541-0064.1980.tb00970.x",
     "https://doi.org/10.1111/j.1541-0064.1980.tb00970.x"),
    ("Agrawal & Goyal (2012) Thompson Sampling analysis, COLT", "", None),
])
shape(s6, MSO_SHAPE.ROUNDED_RECTANGLE, 0.4, 6.32, 12.53, 0.5, fill=T_BLUE, radius=0.3, margin=0.2,
      paras=[{"runs": [("Supporting research: ", {"bold": True, "color": NAVY}),
                       ("Detailed Business Model Canvas, architecture, competitor analysis & ROI are in our Additional Document "
                        "(PRAHARI_Additional_Document.pdf).", {})]}], size=10.5, color=INK)

prs.save(str(OUT))
print("wrote", OUT)
