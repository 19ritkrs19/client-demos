#!/usr/bin/env python3
"""Generate a 15-slide PPT: Parking and Traffic Management: Issues and Challenges."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---- Theme colors ----
NAVY = RGBColor(0x0B, 0x2E, 0x4F)
BLUE = RGBColor(0x1B, 0x6C, 0xA8)
ORANGE = RGBColor(0xF2, 0x8C, 0x28)
LIGHT = RGBColor(0xF4, 0xF7, 0xFA)
GREY = RGBColor(0x55, 0x5F, 0x6B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x8B, 0x57)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def box(slide, l, t, w, h, color):
    sp = slide.shapes.add_shape(1, l, t, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def txt(slide, l, t, w, h, text, size, color, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, font="Calibri"):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font
    return tb


def bullets(slide, l, t, w, h, items, size=18, color=NAVY, gap=6):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (txt_, sub) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        r = p.add_run()
        r.text = ("•  " if not sub else "     – ") + txt_
        r.font.size = Pt(size if not sub else size - 3)
        r.font.color.rgb = color if not sub else GREY
        r.font.bold = False
        r.font.name = "Calibri"
    return tb


def header(slide, title, num):
    box(slide, 0, 0, SW, Inches(1.15), NAVY)
    box(slide, 0, Inches(1.15), SW, Inches(0.08), ORANGE)
    txt(slide, Inches(0.5), Inches(0.2), Inches(11.5), Inches(0.8), title, 28,
        WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, Inches(12.4), Inches(0.2), Inches(0.7), Inches(0.8), str(num), 16,
        ORANGE, bold=True, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


def draw_traffic_illustration(slide, l, t, w, h, caption):
    """A simple vector 'street scene' illustration so no external image is needed."""
    # sky / backdrop
    panel = box(slide, l, t, w, h, RGBColor(0xDD, 0xE8, 0xF0))
    panel.line.color.rgb = BLUE
    panel.line.width = Pt(1.5)
    road_t = t + h - Inches(1.7)
    # road
    box(slide, l, road_t, w, Inches(1.4), RGBColor(0x53, 0x59, 0x5F))
    # lane markings
    lx = l + Inches(0.3)
    while lx < l + w - Inches(0.4):
        box(slide, lx, road_t + Inches(0.65), Inches(0.35), Inches(0.08), RGBColor(0xF5, 0xD9, 0x4E))
        lx += Inches(0.75)
    # footpath
    box(slide, l, road_t - Inches(0.18), w, Inches(0.18), RGBColor(0xB9, 0xC3, 0xCC))
    # haphazardly parked "vehicles" (cars) crowding the road
    car_colors = [ORANGE, RGBColor(0xC0,0x39,0x2B), BLUE, GREEN, RGBColor(0x8E,0x44,0xAD)]
    cx = l + Inches(0.35)
    import random as _r
    _r.seed(7)
    for i in range(6):
        cy = road_t + Inches(0.12) + Inches(0.05 * (i % 3))
        cw = Inches(1.15)
        body = box(slide, cx, cy, cw, Inches(0.5), car_colors[i % len(car_colors)])
        body.line.color.rgb = WHITE
        box(slide, cx + Inches(0.18), cy - Inches(0.22), Inches(0.8), Inches(0.28),
            car_colors[i % len(car_colors)])
        cx += cw + Inches(0.12)
    # label chip
    chip = box(slide, l + Inches(0.25), t + Inches(0.25), w - Inches(0.5), Inches(0.55), NAVY)
    txt(slide, l + Inches(0.25), t + Inches(0.25), w - Inches(0.5), Inches(0.55),
        "PATNA STREET SCENE", 13, WHITE, bold=True, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE)
    # caption below inside panel
    txt(slide, l + Inches(0.2), t + h - Inches(0.55), w - Inches(0.4), Inches(0.5),
        caption, 11, NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


import os as _os
_IMG_DIR = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "ppt_images")


def img_placeholder(slide, l, t, w, h, caption, img_file=None):
    """Embed a real image if provided; else fall back to the vector scene."""
    path = _os.path.join(_IMG_DIR, img_file) if img_file else None
    if path and _os.path.exists(path):
        from PIL import Image as _PILImage
        iw, ih = _PILImage.open(path).size
        cap_h = Inches(0.55)
        inner_w = w - Inches(0.16)
        disp_h = int(inner_w * ih / iw)  # keep aspect ratio
        # frame sized to image + caption (no white gap)
        frame_h = disp_h + Inches(0.16) + cap_h
        frame = box(slide, l, t, w, frame_h, WHITE)
        frame.line.color.rgb = BLUE
        frame.line.width = Pt(1.5)
        slide.shapes.add_picture(path, l + Inches(0.08), t + Inches(0.08), width=inner_w)
        cap_t = t + Inches(0.08) + disp_h + Inches(0.04)
        box(slide, l, cap_t, w, cap_h, NAVY)
        txt(slide, l + Inches(0.1), cap_t, w - Inches(0.2), cap_h, caption, 11,
            WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    else:
        draw_traffic_illustration(slide, l, t, w, h, caption)


# ================= SLIDE 1 — Title =================
s = prs.slides.add_slide(BLANK)
bg(s, NAVY)
box(s, 0, Inches(2.7), SW, Inches(0.12), ORANGE)
txt(s, Inches(1), Inches(1.4), Inches(11.3), Inches(1.3),
    "Parking & Traffic Management", 46, WHITE, bold=True, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(2.85), Inches(11.3), Inches(0.9),
    "Issues and Challenges", 30, ORANGE, bold=True, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(4.2), Inches(11.3), Inches(0.6),
    "A Session for City Managers  •  Patna", 20, WHITE, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(6.4), Inches(11.3), Inches(0.5),
    "Urban Mobility  |  Duration: 1.5 hours", 14, RGBColor(0xAF,0xC3,0xD6),
    align=PP_ALIGN.CENTER)

# ================= SLIDE 2 — Agenda =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Session Agenda", 2)
bullets(s, Inches(0.8), Inches(1.6), Inches(11.5), Inches(5.4), [
    ("Why parking & traffic matter for cities", False),
    ("The scale of the problem in Indian cities", False),
    ("Key parking issues — on-street & off-street", False),
    ("Traffic management challenges", False),
    ("Root causes behind the chaos", False),
    ("Impact: economy, environment, safety", False),
    ("Solutions & smart interventions", False),
    ("Case study: Pune on-street parking reform", False),
    ("Action roadmap for Patna", False),
], size=20, gap=10)

# ================= SLIDE 3 — Why it matters =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Why Parking & Traffic Matter", 3)
bullets(s, Inches(0.7), Inches(1.6), Inches(7.2), Inches(5.2), [
    ("Traffic reflects a city's quality of life & governance", False),
    ("Poor parking eats scarce public road space", False),
    ("Congestion = lost productivity & fuel", False),
    ("Directly affects safety, emergency access & emissions", False),
    ("A city that manages mobility well attracts investment", False),
], size=18, gap=12)
img_placeholder(s, Inches(8.2), Inches(1.7), Inches(4.6), Inches(4.8),
                "Congested urban street — mixed traffic & on-street parking",
                img_file="real2.jpg")

# ================= SLIDE 4 — Scale of the problem =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "The Scale of the Problem in Patna", 4)
stats = [("12.5 L+", "vehicles on Patna roads", BLUE),
         ("62%+", "new vehicles are 2-wheelers", ORANGE),
         ("~8%", "city area is road (vs 25% ideal)", RGBColor(0xC0,0x39,0x2B)),
         ("798", "traffic cops for the whole city", GREEN)]
x = Inches(0.7)
for val, lab, col in stats:
    card = box(s, x, Inches(2.0), Inches(2.85), Inches(2.6), WHITE)
    card.line.color.rgb = col
    card.line.width = Pt(2)
    txt(s, x, Inches(2.3), Inches(2.85), Inches(1.0), val, 36, col, bold=True,
        align=PP_ALIGN.CENTER)
    txt(s, x, Inches(3.45), Inches(2.85), Inches(1.0), lab, 14, GREY,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    x += Inches(3.05)
txt(s, Inches(0.7), Inches(5.2), Inches(12), Inches(1.2),
    "Vehicles have grown roughly threefold in a decade, while road space and "
    "enforcement staff have not kept pace — the core of Patna's congestion.", 15, NAVY)
txt(s, Inches(0.7), Inches(6.5), Inches(12), Inches(0.5),
    "Sources: Times of India (Patna) reports; NIUA public perception study.", 11, GREY)

# ================= SLIDE 5 — Parking issues =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Key Parking Issues in Patna", 5)
bullets(s, Inches(0.7), Inches(1.6), Inches(7.2), Inches(5.4), [
    ("Illegal roadside parking on both sides of arterial roads", False),
    ("Footpaths & carriageways encroached by parked vehicles", False),
    ("Autos & e-rickshaws halting anywhere near stations/markets", False),
    ("Very little organised off-street / paid parking", False),
    ("Choke points: Boring Road, Gandhi Maidan, Station Road", False),
    ("Weak enforcement — too few cops for 12.5L+ vehicles", False),
    ("Parking seen as free & unlimited by users", False),
], size=16, gap=9)
img_placeholder(s, Inches(8.2), Inches(1.7), Inches(4.6), Inches(4.8),
                "Illegal roadside parking narrowing a busy road",
                img_file="real1.jpg")

# ================= SLIDE 6 — Traffic management challenges =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Traffic Management Challenges in Patna", 6)
bullets(s, Inches(0.8), Inches(1.6), Inches(11.6), Inches(5.4), [
    ("Mixed traffic — cars, 2-wheelers, autos, e-rickshaws, carts, pedestrians", False),
    ("Only ~8% of city area is road (well below the 25% norm)", False),
    ("Ongoing construction & broken roads slow traffic to a crawl", False),
    ("Encroachment by vendors & parked vehicles shrinks road width", False),
    ("Weak public transport pushes people to private vehicles", False),
    ("Severe manpower shortage — ~798 cops for 12.5L+ vehicles", False),
    ("Peaks worsen during school & office hours at key crossings", False),
], size=17, gap=10)

# ================= SLIDE 7 — Root causes =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Root Causes Behind the Chaos", 7)
causes = [("Rapid Motorisation", "Vehicle ownership rising faster than infra"),
          ("Weak Enforcement", "Low penalties, poor compliance culture"),
          ("No Parking Policy", "Parking treated as free & unlimited"),
          ("Land-use Mismatch", "Dense commercial zones, no parking norms"),
          ("Weak Public Transport", "Few reliable last-mile options"),
          ("Fragmented Governance", "Multiple agencies, unclear ownership")]
x, y = Inches(0.7), Inches(1.7)
for i, (h, d) in enumerate(causes):
    col = x + (Inches(4.05) * (i % 3))
    row = y + (Inches(2.35) * (i // 3))
    card = box(s, col, row, Inches(3.8), Inches(2.1), WHITE)
    card.line.color.rgb = BLUE
    txt(s, col + Inches(0.2), row + Inches(0.2), Inches(3.4), Inches(0.6),
        h, 17, BLUE, bold=True)
    txt(s, col + Inches(0.2), row + Inches(0.85), Inches(3.4), Inches(1.1),
        d, 13, GREY)

# ================= SLIDE 8 — Impact =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Impact of Poor Management", 8)
cols = [("Economic", ["Lost man-hours in jams", "Higher fuel & logistics cost",
                       "Reduced business footfall"], ORANGE),
        ("Environmental", ["Higher vehicular emissions", "Idling pollution & noise",
                           "Poor air quality (AQI)"], GREEN),
        ("Social & Safety", ["More accidents & fatalities", "Blocked emergency vehicles",
                             "Stress & road rage"], RGBColor(0xC0,0x39,0x2B))]
x = Inches(0.7)
for title, items, col in cols:
    box(s, x, Inches(1.7), Inches(3.95), Inches(0.7), col)
    txt(s, x, Inches(1.7), Inches(3.95), Inches(0.7), title, 18, WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    card = box(s, x, Inches(2.4), Inches(3.95), Inches(3.6), WHITE)
    card.line.color.rgb = col
    bullets(s, x + Inches(0.25), Inches(2.65), Inches(3.5), Inches(3.2),
            [(it, False) for it in items], size=14, gap=12)
    x += Inches(4.15)

# ================= SLIDE 9 — Solutions (parking) =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Solutions: Smarter Parking", 9)
bullets(s, Inches(0.8), Inches(1.6), Inches(11.6), Inches(5.4), [
    ("Priced, demand-based on-street parking (pay-and-park)", False),
    ("Clearly marked bays + no-parking zones", False),
    ("Multi-level & underground off-street parking", False),
    ("Digital parking — apps, sensors, cashless payment", False),
    ("Strict enforcement: towing, fines, CCTV", False),
    ("Parking norms in building bye-laws for new projects", False),
    ("Revenue from parking funds better streets & transit", False),
], size=18, gap=11)

# ================= SLIDE 10 — Solutions (traffic) =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Solutions: Traffic Management", 10)
bullets(s, Inches(0.8), Inches(1.6), Inches(11.6), Inches(5.4), [
    ("Adaptive / synchronised signal systems (ITS)", False),
    ("Dedicated lanes for buses & non-motorised transport", False),
    ("Strengthen public transport & last-mile connectivity", False),
    ("Junction redesign & pedestrian-first street design", False),
    ("Decongest by removing encroachments", False),
    ("Real-time monitoring, CCTV & data-driven decisions", False),
    ("Public awareness & lane discipline campaigns", False),
], size=18, gap=11)

# ================= SLIDE 11 — Smart tech =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "The Role of Smart Technology", 11)
tech = [("Smart Sensors", "Detect free/occupied bays in real time"),
        ("Parking Apps", "Locate, book & pay for parking digitally"),
        ("ANPR Cameras", "Automatic number-plate enforcement"),
        ("Adaptive Signals", "AI-based signal timing on traffic flow"),
        ("Data Dashboards", "Live congestion & occupancy analytics"),
        ("Cashless Payment", "Transparent, leak-proof revenue")]
x, y = Inches(0.7), Inches(1.7)
for i, (h, d) in enumerate(tech):
    col = x + (Inches(4.05) * (i % 3))
    row = y + (Inches(2.35) * (i // 3))
    card = box(s, col, row, Inches(3.8), Inches(2.1), NAVY)
    txt(s, col + Inches(0.2), row + Inches(0.2), Inches(3.4), Inches(0.6),
        h, 16, ORANGE, bold=True)
    txt(s, col + Inches(0.2), row + Inches(0.85), Inches(3.4), Inches(1.1),
        d, 13, WHITE)

# ================= SLIDE 12 — Case study: Pune =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Case Study: Pune Parking Reform", 12)
box(s, Inches(0.7), Inches(1.6), Inches(0.15), Inches(5.1), ORANGE)
txt(s, Inches(1.0), Inches(1.6), Inches(11.3), Inches(0.6),
    "Pune's on-street parking management (with ITDP India)", 20, BLUE, bold=True)
bullets(s, Inches(1.0), Inches(2.4), Inches(11.5), Inches(4.2), [
    ("Adopted a city parking policy — parking priced, not free", False),
    ("Demand-based pricing: higher rates in busy zones & peak hours", False),
    ("Clearly demarcated, managed on-street bays", False),
    ("Cashless, self-paying system for transparent revenue", False),
    ("Freed footpaths & road space for people and buses", False),
    ("Result: reduced haphazard parking & smoother flow on key roads", False),
    ("Lesson: policy + pricing + enforcement + tech together works", True),
], size=16, gap=9)

# ================= SLIDE 13 — Roadmap for Patna =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Action Roadmap for Patna", 13)
phases = [("Short Term (0-6 m)", ["Fix hotspot zones (Boring Rd, Station Rd)",
                                  "Demarcate parking bays", "Clear encroachments",
                                  "Enforcement drive + towing"], ORANGE),
          ("Medium (6-18 m)", ["Adopt a city parking policy", "Pay-and-park + app",
                               "Signal & junction fixes", "Strengthen city bus service"], BLUE),
          ("Long Term (18 m+)", ["Multi-level parking", "City-wide ITS & CCTV",
                                 "Integrated mobility plan", "Data-driven governance"], GREEN)]
x = Inches(0.7)
for title, items, col in phases:
    box(s, x, Inches(1.6), Inches(3.95), Inches(0.75), col)
    txt(s, x, Inches(1.6), Inches(3.95), Inches(0.75), title, 16, WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    card = box(s, x, Inches(2.35), Inches(3.95), Inches(3.4), WHITE)
    card.line.color.rgb = col
    bullets(s, x + Inches(0.25), Inches(2.6), Inches(3.5), Inches(3.0),
            [(it, False) for it in items], size=13, gap=9)
    x += Inches(4.15)
box(s, Inches(0.7), Inches(5.95), Inches(11.9), Inches(0.85), NAVY)
txt(s, Inches(0.9), Inches(5.95), Inches(11.5), Inches(0.85),
    "Already begun: In late 2025 Patna traffic police banned car/auto parking on "
    "Income Tax Golambar–Dak Bungalow & Budh Marg–GPO Golambar to decongest key stretches.",
    13, WHITE, anchor=MSO_ANCHOR.MIDDLE)

# ================= SLIDE 14 — Key takeaways =================
s = prs.slides.add_slide(BLANK)
bg(s, NAVY)
box(s, 0, Inches(1.15), SW, Inches(0.08), ORANGE)
txt(s, Inches(0.5), Inches(0.25), Inches(11), Inches(0.8), "Key Takeaways", 30,
    WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(0.9), Inches(1.7), Inches(11.5), Inches(5.2), [
    ("Parking is a policy problem, not just an infrastructure one", False),
    ("Free & unmanaged parking is the root of much congestion", False),
    ("Pricing + enforcement + technology must go together", False),
    ("Strong public transport reduces parking demand at source", False),
    ("Data and clear ownership drive sustainable results", False),
    ("Patna can adapt proven models like Pune's reform", False),
], size=20, color=WHITE, gap=14)

# ================= SLIDE 15 — Thank you =================
s = prs.slides.add_slide(BLANK)
bg(s, NAVY)
box(s, 0, Inches(3.2), SW, Inches(0.12), ORANGE)
txt(s, Inches(1), Inches(2.4), Inches(11.3), Inches(1.0), "Thank You", 52, WHITE,
    bold=True, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(3.6), Inches(11.3), Inches(0.7),
    "Questions & Discussion", 24, ORANGE, bold=True, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(4.8), Inches(11.3), Inches(0.6),
    "Let's build a well-managed, people-friendly Patna", 18,
    RGBColor(0xAF,0xC3,0xD6), align=PP_ALIGN.CENTER)

out = "/Users/riteshkumar/Desktop/github-enterprise/Parking_Traffic_Management_Patna.pptx"
prs.save(out)
print("Saved:", out, "| slides:", len(prs.slides._sldIdLst))
