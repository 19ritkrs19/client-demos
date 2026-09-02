#!/usr/bin/env python3
"""Generate the Patna Parking & Traffic Management presentation.

Session by Pourush Agarwal, DySP, Patna Traffic Police.
Uses real Patna photos and covers: VIP movement, metro diversion,
sign boards, real-time updates.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import os as _os

# ---- Theme colors ----
NAVY = RGBColor(0x0B, 0x2E, 0x4F)
BLUE = RGBColor(0x1B, 0x6C, 0xA8)
ORANGE = RGBColor(0xF2, 0x8C, 0x28)
LIGHT = RGBColor(0xF4, 0xF7, 0xFA)
GREY = RGBColor(0x55, 0x5F, 0x6B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREEN = RGBColor(0x2E, 0x8B, 0x57)
RED = RGBColor(0xC0, 0x39, 0x2B)
PURPLE = RGBColor(0x6C, 0x3A, 0x8E)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

_IMG_DIR = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "ppt_images")


def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def box(slide, l, t, w, h, color, line=None, lw=None):
    sp = slide.shapes.add_shape(1, l, t, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    if line is not None:
        sp.line.color.rgb = line
        sp.line.width = Pt(lw or 1.5)
    else:
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
        r.font.name = "Calibri"
    return tb


def header(slide, title, num):
    box(slide, 0, 0, SW, Inches(1.15), NAVY)
    box(slide, 0, Inches(1.15), SW, Inches(0.08), ORANGE)
    txt(slide, Inches(0.5), Inches(0.2), Inches(11.5), Inches(0.8), title, 27,
        WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    txt(slide, Inches(12.35), Inches(0.2), Inches(0.7), Inches(0.8), str(num), 15,
        ORANGE, bold=True, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)


def photo(slide, l, t, w, img_file, caption, max_h=None):
    """Embed a real photo, keep aspect ratio, add a navy caption strip."""
    from PIL import Image as _PILImage
    path = _os.path.join(_IMG_DIR, img_file)
    iw, ih = _PILImage.open(path).size
    inner_w = w - Inches(0.16)
    disp_h = int(inner_w * ih / iw)
    if max_h is not None and disp_h > (max_h - Inches(0.71)):
        # constrain by height instead
        disp_h = max_h - Inches(0.71)
        inner_w = int(disp_h * iw / ih)
    cap_h = Inches(0.55)
    frame_h = disp_h + Inches(0.16) + cap_h
    frame = box(slide, l, t, w, frame_h, WHITE, line=BLUE, lw=1.5)
    px = l + (w - inner_w) / 2
    slide.shapes.add_picture(path, px, t + Inches(0.08), width=inner_w)
    cap_t = t + Inches(0.08) + disp_h + Inches(0.04)
    box(slide, l, cap_t, w, cap_h, NAVY)
    txt(slide, l + Inches(0.1), cap_t, w - Inches(0.2), cap_h, caption, 11,
        WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return frame_h


# ================= SLIDE 1 — Title (cover photo + DySP credit) =================
s = prs.slides.add_slide(BLANK)
bg(s, NAVY)
# cover photo on the right
from PIL import Image as _PIL
_ip = _os.path.join(_IMG_DIR, "patna_cover.jpg")
_iw, _ih = _PIL.open(_ip).size
_ph = Inches(7.5)
_pw = int(_ph * _iw / _ih)
s.shapes.add_picture(_ip, SW - _pw, 0, height=_ph)
# left navy panel for text
box(s, 0, 0, Inches(7.6), SH, NAVY)
box(s, Inches(0.7), Inches(2.35), Inches(0.18), Inches(2.4), ORANGE)
txt(s, Inches(1.0), Inches(1.3), Inches(6.2), Inches(1.4),
    "Parking & Traffic Management", 34, WHITE, bold=True)
txt(s, Inches(1.0), Inches(2.55), Inches(6.2), Inches(0.8),
    "Issues and Challenges", 24, ORANGE, bold=True)
txt(s, Inches(1.0), Inches(3.5), Inches(6.2), Inches(0.6),
    "A Session for City Managers  •  Patna", 16, WHITE)
txt(s, Inches(1.0), Inches(5.5), Inches(6.2), Inches(0.5),
    "Session by", 13, RGBColor(0xAF, 0xC3, 0xD6))
txt(s, Inches(1.0), Inches(5.85), Inches(6.2), Inches(0.55),
    "Pourush Agarwal", 20, WHITE, bold=True)
txt(s, Inches(1.0), Inches(6.4), Inches(6.2), Inches(0.5),
    "DySP, Patna Traffic Police", 15, ORANGE, bold=True)

# ================= SLIDE 2 — Agenda =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Session Agenda", 2)
bullets(s, Inches(0.8), Inches(1.5), Inches(11.5), Inches(5.6), [
    ("Why parking & traffic matter for a city", False),
    ("The scale of the problem in Patna", False),
    ("Key parking issues & the role of sign boards", False),
    ("Traffic management challenges", False),
    ("Special situations: VIP movement & metro diversion", False),
    ("Root causes and their impact", False),
    ("Solutions — smarter parking, signage & real-time updates", False),
    ("Smart technology in traffic management", False),
    ("Case study: Pune parking reform", False),
    ("Action roadmap for Patna", False),
], size=19, gap=9)

# ================= SLIDE 3 — Why it matters =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Why Parking & Traffic Matter", 3)
bullets(s, Inches(0.7), Inches(1.6), Inches(11.8), Inches(5.2), [
    ("Traffic reflects a city's quality of life & governance", False),
    ("Poor parking eats into scarce public road space", False),
    ("Congestion means lost productivity, fuel and time", False),
    ("It directly affects safety, emergency access & emissions", False),
    ("Well-managed mobility builds public trust & attracts investment", False),
], size=19, gap=14)

# ================= SLIDE 4 — Scale of the problem =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "The Scale of the Problem in Patna", 4)
stats = [("12.5 L+", "vehicles on Patna roads", BLUE),
         ("62%+", "new vehicles are 2-wheelers", ORANGE),
         ("~8%", "city area is road (vs 25% ideal)", RED),
         ("798", "traffic cops for the whole city", GREEN)]
x = Inches(0.7)
for val, lab, col in stats:
    box(s, x, Inches(2.0), Inches(2.85), Inches(2.6), WHITE, line=col, lw=2)
    txt(s, x, Inches(2.3), Inches(2.85), Inches(1.0), val, 36, col, bold=True,
        align=PP_ALIGN.CENTER)
    txt(s, x, Inches(3.45), Inches(2.85), Inches(1.0), lab, 14, GREY,
        align=PP_ALIGN.CENTER)
    x += Inches(3.05)
txt(s, Inches(0.7), Inches(5.2), Inches(12), Inches(1.2),
    "Vehicles have grown roughly threefold in a decade, while road space and "
    "enforcement staff have not kept pace — the core of Patna's congestion.", 15, NAVY)
txt(s, Inches(0.7), Inches(6.5), Inches(12), Inches(0.5),
    "Sources: Times of India (Patna) reports; NIUA public perception study.", 11, GREY)

# ================= SLIDE 5 — Parking issues + No Parking board photo =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Key Parking Issues in Patna", 5)
bullets(s, Inches(0.6), Inches(1.5), Inches(6.9), Inches(5.6), [
    ("Illegal roadside parking on both sides of arterial roads", False),
    ("Vehicles ignore signage — parking right under 'No Parking' boards", False),
    ("Footpaths & carriageways encroached by parked vehicles", False),
    ("Autos & e-rickshaws halt anywhere near stations & markets", False),
    ("Very little organised, paid off-street parking", False),
    ("Weak deterrence — sign boards alone don't ensure compliance", False),
], size=15, gap=10)
photo(s, Inches(7.8), Inches(1.5), Inches(5.1), "patna_noparking.jpg",
      "Autos parked right under a 'No Parking' board — Patna", max_h=Inches(5.5))

# ================= SLIDE 6 — Traffic management challenges + regulation photo =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Traffic Management Challenges", 6)
bullets(s, Inches(0.6), Inches(1.5), Inches(6.9), Inches(5.6), [
    ("Mixed traffic — cars, 2-wheelers, autos, e-rickshaws, carts", False),
    ("Only ~8% of city area is road (well below the 25% norm)", False),
    ("Manual regulation at junctions with severe manpower shortage", False),
    ("Frequent VIP movement disrupts normal traffic flow", False),
    ("Metro construction forces route diversions & lane closures", False),
    ("Encroachment & construction shrink usable road width", False),
], size=15, gap=10)
photo(s, Inches(7.8), Inches(1.5), Inches(5.1), "patna_regulation.jpg",
      "Patna Traffic Police regulating a busy junction", max_h=Inches(5.5))

# ================= SLIDE 7 — Special situations: VIP movement & Metro diversion =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Special Traffic Situations", 7)
# Two detailed cards
box(s, Inches(0.6), Inches(1.5), Inches(6.0), Inches(2.4), WHITE, line=PURPLE, lw=2)
txt(s, Inches(0.85), Inches(1.65), Inches(5.5), Inches(0.5), "VIP Movement", 19,
    PURPLE, bold=True)
bullets(s, Inches(0.85), Inches(2.25), Inches(5.5), Inches(1.6), [
    ("Route clearance & holding of traffic during transit", False),
    ("Advance planning with escort & pilot coordination", False),
    ("Minimise public inconvenience with timed, short holds", False),
], size=13, gap=6)
box(s, Inches(6.9), Inches(1.5), Inches(6.0), Inches(2.4), WHITE, line=BLUE, lw=2)
txt(s, Inches(7.15), Inches(1.65), Inches(5.5), Inches(0.5), "Metro Diversion", 19,
    BLUE, bold=True)
bullets(s, Inches(7.15), Inches(2.25), Inches(5.5), Inches(1.6), [
    ("Lane closures around metro construction sites", False),
    ("Planned diversions on key corridors & junctions", False),
    ("Clear signage + advance public communication needed", False),
], size=13, gap=6)
# supporting photo below
photo(s, Inches(3.3), Inches(4.1), Inches(6.7), "patna_boringroad.jpg",
      "On-ground regulation at Boring Road crossing, Patna", max_h=Inches(3.2))

# ================= SLIDE 8 — Root causes =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Root Causes Behind the Chaos", 8)
causes = [("Rapid Motorisation", "Vehicle ownership rising faster than infra"),
          ("Weak Enforcement", "Low penalties, poor signage compliance"),
          ("No Parking Policy", "Parking treated as free & unlimited"),
          ("Land-use Mismatch", "Dense commercial zones, no parking norms"),
          ("Weak Public Transport", "Few reliable last-mile options"),
          ("Fragmented Governance", "Multiple agencies, unclear ownership")]
x, y = Inches(0.7), Inches(1.6)
for i, (h, d) in enumerate(causes):
    col = x + (Inches(4.05) * (i % 3))
    row = y + (Inches(2.4) * (i // 3))
    box(s, col, row, Inches(3.8), Inches(2.15), WHITE, line=BLUE)
    txt(s, col + Inches(0.2), row + Inches(0.2), Inches(3.4), Inches(0.6),
        h, 17, BLUE, bold=True)
    txt(s, col + Inches(0.2), row + Inches(0.9), Inches(3.4), Inches(1.1),
        d, 13, GREY)

# ================= SLIDE 9 — Impact =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Impact of Poor Management", 9)
cols = [("Economic", ["Lost man-hours in jams", "Higher fuel & logistics cost",
                       "Reduced business footfall"], ORANGE),
        ("Environmental", ["Higher vehicular emissions", "Idling pollution & noise",
                           "Poor air quality (AQI)"], GREEN),
        ("Social & Safety", ["More accidents & fatalities", "Blocked emergency vehicles",
                             "Stress & road rage"], RED)]
x = Inches(0.7)
for title, items, col in cols:
    box(s, x, Inches(1.7), Inches(3.95), Inches(0.7), col)
    txt(s, x, Inches(1.7), Inches(3.95), Inches(0.7), title, 18, WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, Inches(2.4), Inches(3.95), Inches(3.6), WHITE, line=col)
    bullets(s, x + Inches(0.25), Inches(2.65), Inches(3.5), Inches(3.2),
            [(it, False) for it in items], size=14, gap=12)
    x += Inches(4.15)

# ================= SLIDE 10 — Smarter Parking + proper parking photo =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Solutions: Smarter Parking", 10)
bullets(s, Inches(0.6), Inches(1.5), Inches(6.9), Inches(5.6), [
    ("Priced, demand-based on-street parking (pay-and-park)", False),
    ("Clearly marked bays with barricades & discipline", False),
    ("Multi-level & organised off-street parking", False),
    ("Designated auto/e-rickshaw stands near hubs", False),
    ("Strict enforcement: towing, fines, CCTV", False),
    ("Parking norms in building bye-laws for new projects", False),
], size=15, gap=10)
photo(s, Inches(7.8), Inches(1.5), Inches(5.1), "patna_properparking.jpg",
      "Organised auto stand with barricades — Patna", max_h=Inches(5.5))

# ================= SLIDE 11 — Traffic Mgmt solutions (sign boards + real-time updates) =====
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Solutions: Traffic Management", 11)
bullets(s, Inches(0.8), Inches(1.5), Inches(11.6), Inches(5.6), [
    ("Clear, well-placed sign boards + strict enforcement of signage", False),
    ("Real-time updates to public — social media, apps, FM & alerts", False),
    ("Advance communication for VIP movement & metro diversions", False),
    ("Adaptive / synchronised signals on key corridors (ITS)", False),
    ("Dedicated lanes for buses & non-motorised transport", False),
    ("Decongest by removing encroachments at choke points", False),
    ("Strengthen public transport & last-mile connectivity", False),
], size=16, gap=11)

# ================= SLIDE 12 — Smart technology (real-time updates reinforced) =========
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "The Role of Smart Technology", 12)
tech = [("Smart Sensors", "Detect free/occupied parking bays in real time"),
        ("Digital Sign Boards", "Dynamic messages for diversions & warnings"),
        ("ANPR Cameras", "Automatic number-plate enforcement"),
        ("Adaptive Signals", "AI-based signal timing on live traffic flow"),
        ("Real-time Updates", "Live traffic & diversion info to citizens"),
        ("Cashless Payment", "Transparent, leak-proof parking revenue")]
x, y = Inches(0.7), Inches(1.6)
for i, (h, d) in enumerate(tech):
    col = x + (Inches(4.05) * (i % 3))
    row = y + (Inches(2.4) * (i // 3))
    box(s, col, row, Inches(3.8), Inches(2.15), NAVY)
    txt(s, col + Inches(0.2), row + Inches(0.2), Inches(3.4), Inches(0.6),
        h, 16, ORANGE, bold=True)
    txt(s, col + Inches(0.2), row + Inches(0.9), Inches(3.4), Inches(1.1),
        d, 13, WHITE)

# ================= SLIDE 13 — Case study: Pune =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Case Study: Pune Parking Reform", 13)
box(s, Inches(0.7), Inches(1.6), Inches(0.15), Inches(5.1), ORANGE)
txt(s, Inches(1.0), Inches(1.6), Inches(11.3), Inches(0.6),
    "Pune's on-street parking management (with ITDP India)", 20, BLUE, bold=True)
bullets(s, Inches(1.0), Inches(2.4), Inches(11.5), Inches(4.2), [
    ("Adopted a city parking policy — parking priced, not free", False),
    ("Demand-based pricing: higher rates in busy zones & peak hours", False),
    ("Clearly demarcated, managed on-street bays with signage", False),
    ("Cashless, self-paying system for transparent revenue", False),
    ("Freed footpaths & road space for people and buses", False),
    ("Result: reduced haphazard parking & smoother flow on key roads", False),
    ("Lesson: policy + pricing + enforcement + technology together works", True),
], size=16, gap=9)

# ================= SLIDE 14 — Roadmap for Patna =================
s = prs.slides.add_slide(BLANK)
bg(s, LIGHT)
header(s, "Action Roadmap for Patna", 14)
phases = [("Short Term (0-6 m)", ["Fix hotspots (Boring Rd, Station Rd)",
                                  "Demarcate bays + clear signage",
                                  "Clear encroachments", "Enforcement drive + towing"], ORANGE),
          ("Medium (6-18 m)", ["Adopt a city parking policy", "Pay-and-park + app",
                               "Real-time updates channel", "Metro diversion planning"], BLUE),
          ("Long Term (18 m+)", ["Multi-level parking", "City-wide ITS & CCTV",
                                 "Integrated mobility plan", "Data-driven governance"], GREEN)]
x = Inches(0.7)
for title, items, col in phases:
    box(s, x, Inches(1.55), Inches(3.95), Inches(0.75), col)
    txt(s, x, Inches(1.55), Inches(3.95), Inches(0.75), title, 16, WHITE, bold=True,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    box(s, x, Inches(2.3), Inches(3.95), Inches(3.4), WHITE, line=col)
    bullets(s, x + Inches(0.22), Inches(2.55), Inches(3.55), Inches(3.0),
            [(it, False) for it in items], size=13, gap=9)
    x += Inches(4.15)
box(s, Inches(0.7), Inches(5.9), Inches(11.9), Inches(0.85), NAVY)
txt(s, Inches(0.9), Inches(5.9), Inches(11.5), Inches(0.85),
    "Already begun: In late 2025 Patna traffic police banned car/auto parking on "
    "Income Tax Golambar–Dak Bungalow & Budh Marg–GPO Golambar to decongest key stretches.",
    13, WHITE, anchor=MSO_ANCHOR.MIDDLE)

# ================= SLIDE 15 — Key takeaways =================
s = prs.slides.add_slide(BLANK)
bg(s, NAVY)
box(s, 0, Inches(1.15), SW, Inches(0.08), ORANGE)
txt(s, Inches(0.5), Inches(0.25), Inches(11), Inches(0.8), "Key Takeaways", 30,
    WHITE, bold=True, anchor=MSO_ANCHOR.MIDDLE)
bullets(s, Inches(0.9), Inches(1.6), Inches(11.5), Inches(5.4), [
    ("Parking is a policy problem, not just an infrastructure one", False),
    ("Sign boards must be backed by real enforcement", False),
    ("Plan ahead for VIP movement & metro diversions", False),
    ("Real-time updates keep citizens informed & reduce chaos", False),
    ("Technology + pricing + enforcement must go together", False),
    ("Patna can adapt proven models like Pune's reform", False),
], size=19, color=WHITE, gap=13)

# ================= SLIDE 16 — Thank you =================
s = prs.slides.add_slide(BLANK)
bg(s, NAVY)
box(s, 0, Inches(3.2), SW, Inches(0.12), ORANGE)
txt(s, Inches(1), Inches(2.3), Inches(11.3), Inches(1.0), "Thank You", 50, WHITE,
    bold=True, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(3.5), Inches(11.3), Inches(0.7),
    "Questions & Discussion", 22, ORANGE, bold=True, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(4.6), Inches(11.3), Inches(0.5),
    "Session by Pourush Agarwal, DySP, Patna Traffic Police", 15,
    WHITE, align=PP_ALIGN.CENTER)
txt(s, Inches(1), Inches(5.3), Inches(11.3), Inches(0.5),
    "Let's build a well-managed, people-friendly Patna", 15,
    RGBColor(0xAF, 0xC3, 0xD6), align=PP_ALIGN.CENTER)

out = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)),
                    "Parking_Traffic_Management_Patna.pptx")
prs.save(out)
print("Saved:", out, "| slides:", len(prs.slides._sldIdLst))
