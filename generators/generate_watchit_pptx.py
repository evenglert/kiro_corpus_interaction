"""
WatchIt — Management Metrics PPTX Template Generator
Design tokens extracted from:
  Evgeniya Englert - Data-Driven Licensing Recommendations for Watch It.pptx

Design palette
  BG_DARK   #0A132E  deep navy  (slide background / dark panels)
  BG_BLACK  #0F0A05  near-black (accent fill on title slide oval)
  GOLD      #DCA523  gold       (primary accent — bars, highlights, CTAs)
  GOLD_WARM #F89F19  amber      (secondary accent)
  WHITE     #FFFFFF
  GREY_LT   #D9D9D9  light grey (body text, dividers)
  GREY_MID  #A6A6A6  mid grey   (secondary labels)
  BLUE_ACC  #3A86FF  bright blue (optional highlight)
  CREAM     #FCFCEA  off-white  (light card backgrounds)

Typography  : Aptos Display (headings) / Aptos (body)  — Office 2024 default
Slide size  : 13.33 × 7.50 in  (widescreen 16:9)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Palette ──────────────────────────────────────────────────────────────────
BG_DARK   = RGBColor(0x0A, 0x13, 0x2E)
BG_BLACK  = RGBColor(0x0F, 0x0A, 0x05)
GOLD      = RGBColor(0xDC, 0xA5, 0x23)
GOLD_WARM = RGBColor(0xF8, 0x9F, 0x19)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GREY_LT   = RGBColor(0xD9, 0xD9, 0xD9)
GREY_MID  = RGBColor(0xA6, 0xA6, 0xA6)
BLUE_ACC  = RGBColor(0x3A, 0x86, 0xFF)
CREAM     = RGBColor(0xFC, 0xFC, 0xEA)
TRANSPARENT = None

# ── Helpers ───────────────────────────────────────────────────────────────────
W = Inches(13.33)
H = Inches(7.50)

def rgb(r, g, b):
    return RGBColor(r, g, b)

def add_rect(slide, x, y, w, h, fill_rgb=None, line_rgb=None, line_w_pt=0,
             rounding=None):
    """Add a rectangle (or rounded rectangle) shape."""
    from pptx.util import Pt as PtU
    from pptx.oxml.ns import qn
    from lxml import etree

    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    fill = shape.fill
    if fill_rgb:
        fill.solid()
        fill.fore_color.rgb = fill_rgb
    else:
        fill.background()

    line = shape.line
    if line_rgb:
        line.color.rgb = line_rgb
        line.width = Pt(line_w_pt) if line_w_pt else Pt(0.5)
    else:
        line.fill.background()

    if rounding is not None:
        # patch the prst geom to roundRect
        spPr = shape.element.find(qn('p:spPr'))
        prstGeom = spPr.find(qn('a:prstGeom'))
        if prstGeom is not None:
            prstGeom.set('prst', 'roundRect')
            avLst = prstGeom.find(qn('a:avLst'))
            if avLst is None:
                avLst = etree.SubElement(prstGeom, qn('a:avLst'))
            else:
                avLst.clear()
            gd = etree.SubElement(avLst, qn('a:gd'))
            gd.set('name', 'adj')
            gd.set('fmla', f'val {rounding}')
    return shape


def add_line(slide, x1, y1, x2, y2, color_rgb=GOLD, width_pt=1.5):
    """Add a straight connector line."""
    from pptx.util import Pt as PtU
    from pptx.oxml.ns import qn
    from lxml import etree

    cx = abs(x2 - x1)
    cy = abs(y2 - y1)
    shape = slide.shapes.add_shape(
        # We use a thin rectangle as a line (1px height or width)
        1, Inches(min(x1, x2)), Inches(min(y1, y2)),
        Inches(max(cx, 0.001)), Inches(max(cy, 0.001))
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color_rgb
    shape.line.fill.background()
    return shape


def add_textbox(slide, x, y, w, h, text, font_name="Aptos", font_size=12,
                font_color=WHITE, bold=False, italic=False,
                align=PP_ALIGN.LEFT, word_wrap=True, margin_top=0,
                margin_left=Inches(0.05)):
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    tf.margin_top = margin_top
    tf.margin_left = margin_left

    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.color.rgb = font_color
    run.font.bold = bold
    run.font.italic = italic
    return txBox


def add_textbox_multiline(slide, x, y, w, h, lines,
                          font_name="Aptos", font_size=11,
                          font_color=WHITE, bold=False,
                          align=PP_ALIGN.LEFT, line_spacing_pt=None):
    """Add a textbox with multiple paragraph lines."""
    from pptx.util import Pt as PtU
    txBox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_top = Inches(0.05)

    for idx, (line_text, line_size, line_bold, line_color) in enumerate(lines):
        if idx == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = align
        if line_spacing_pt:
            p.space_after = Pt(line_spacing_pt)
        run = p.add_run()
        run.text = line_text
        run.font.name = font_name
        run.font.size = Pt(line_size)
        run.font.color.rgb = line_color
        run.font.bold = line_bold
    return txBox


def slide_bg(slide, color_rgb):
    """Fill slide background with a solid color."""
    from pptx.oxml.ns import qn
    from lxml import etree
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color_rgb


def gold_bar_left(slide, h=7.50):
    """Thin gold vertical accent bar on left edge (Evgeniya's signature)."""
    add_rect(slide, 0, 0, 0.06, h, fill_rgb=GOLD)


def gold_bar_bottom(slide):
    """Thin gold horizontal bar at bottom (Evgeniya title slide style)."""
    add_rect(slide, 0, 7.35, 13.33, 0.06, fill_rgb=GOLD)


def footer_bar(slide, page_num):
    """Dark footer bar with page number, matching source deck style."""
    add_rect(slide, 0, 6.92, 13.33, 0.58, fill_rgb=BG_BLACK)
    add_rect(slide, 0, 6.92, 0.06, 0.58, fill_rgb=GOLD)
    add_textbox(slide, 12.0, 6.95, 1.2, 0.45,
                f"Page  {page_num}", font_size=8, font_color=GREY_MID,
                align=PP_ALIGN.RIGHT)


def section_label(slide, x, y, w, text):
    """Gold-accented section title label."""
    add_rect(slide, x, y, 0.05, 0.35, fill_rgb=GOLD)
    add_textbox(slide, x + 0.12, y, w, 0.35, text,
                font_name="Aptos Display", font_size=14,
                font_color=WHITE, bold=True)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def build_title_slide(prs):
    """Slide 1 — Cover / Title (mirrors Evgeniya's title layout)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    slide_bg(slide, BG_DARK)

    # Large dark oval (logo placeholder) — top-left, matching original
    oval_shape = slide.shapes.add_shape(
        9,  # MSO_SHAPE_TYPE.OVAL
        Inches(0.8), Inches(1.2), Inches(2.8), Inches(2.8)
    )
    oval_shape.fill.solid()
    oval_shape.fill.fore_color.rgb = BG_BLACK
    oval_shape.line.fill.background()

    # Logo text inside oval
    add_textbox(slide, 0.85, 1.8, 2.7, 1.2,
                "WatchIt", font_name="Aptos Display", font_size=28,
                font_color=GOLD, bold=True, align=PP_ALIGN.CENTER)
    add_textbox(slide, 0.85, 2.85, 2.7, 0.5,
                "🎬", font_size=20, font_color=GOLD, align=PP_ALIGN.CENTER)

    # Main title
    add_textbox(slide, 4.3, 1.4, 8.5, 1.4,
                "Management Metrics Update",
                font_name="Aptos Display", font_size=32,
                font_color=WHITE, bold=True)

    # Subtitle
    add_textbox(slide, 4.3, 2.9, 8.5, 0.55,
                "Platform Performance · User Growth · Content Strategy",
                font_name="Aptos", font_size=14, font_color=GREY_LT)

    # Presenter line
    add_textbox(slide, 4.3, 4.7, 8.5, 0.45,
                "Presented by:", font_size=10, font_color=GREY_MID)
    add_textbox(slide, 4.3, 5.1, 8.5, 0.45,
                "[Your Name]  ·  [Date]",
                font_size=12, font_color=GOLD, bold=True)

    # Gold bottom bar (title slide signature)
    gold_bar_bottom(slide)
    gold_bar_left(slide, h=7.35)

    return slide


def build_agenda_slide(prs):
    """Slide 2 — Agenda."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 2)

    section_label(slide, 0.4, 0.45, 6, "Agenda")

    items = [
        ("01", "Platform KPI Overview"),
        ("02", "User Metrics & Demographics"),
        ("03", "Content & Catalogue Performance"),
        ("04", "Marketing & Genre Insights"),
        ("05", "Engagement & Retention"),
        ("06", "Key Takeaways & Next Steps"),
    ]

    col_x = [0.55, 6.85]
    for i, (num, label) in enumerate(items):
        row = i % 3
        col = i // 3
        y = 1.4 + row * 1.55
        x = col_x[col]

        # Card background
        add_rect(slide, x, y, 5.9, 1.25, fill_rgb=BG_BLACK, rounding=15000)
        # Gold number
        add_textbox(slide, x + 0.18, y + 0.1, 0.9, 0.9,
                    num, font_name="Aptos Display", font_size=28,
                    font_color=GOLD, bold=True)
        # Label
        add_textbox(slide, x + 1.15, y + 0.3, 4.5, 0.65,
                    label, font_name="Aptos Display", font_size=14,
                    font_color=WHITE, bold=True)

    return slide


def build_kpi_overview_slide(prs):
    """Slide 3 — Platform KPI Overview (6 KPI tiles, empty values)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 3)

    section_label(slide, 0.4, 0.45, 10, "Platform KPI Overview")
    add_textbox(slide, 0.55, 0.95, 12, 0.35,
                "Core platform metrics for the reporting period",
                font_size=10, font_color=GREY_MID, italic=True)

    kpis = [
        ("Total Registered Users",    "______",  "users"),
        ("Total Ratings Submitted",   "______",  "ratings"),
        ("Movies in Catalogue",       "______",  "titles"),
        ("Avg. Rating Score",         "__ / 5",  "platform avg"),
        ("Avg. Ratings per User",     "______",  "ratings/user"),
        ("Monthly Active Users",      "__ %",    "of total base"),
    ]

    cols = 3
    card_w, card_h = 3.9, 2.2
    x_starts = [0.45, 4.72, 8.98]
    y_starts = [1.45, 3.95]

    for i, (label, value, sub) in enumerate(kpis):
        col = i % cols
        row = i // cols
        x = x_starts[col]
        y = y_starts[row]

        # Card
        add_rect(slide, x, y, card_w, card_h, fill_rgb=BG_BLACK, rounding=15000)
        # Gold accent top strip
        add_rect(slide, x, y, card_w, 0.07, fill_rgb=GOLD, rounding=0)

        # Metric label
        add_textbox(slide, x + 0.2, y + 0.15, card_w - 0.3, 0.45,
                    label, font_size=10, font_color=GREY_LT)
        # Big value (placeholder)
        add_textbox(slide, x + 0.2, y + 0.6, card_w - 0.3, 0.85,
                    value, font_name="Aptos Display", font_size=26,
                    font_color=GOLD, bold=True)
        # Sub-label
        add_textbox(slide, x + 0.2, y + 1.55, card_w - 0.3, 0.4,
                    sub, font_size=9, font_color=GREY_MID, italic=True)

    return slide


def build_users_slide(prs):
    """Slide 4 — User Metrics & Demographics."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 4)

    section_label(slide, 0.4, 0.45, 10, "User Metrics & Demographics")

    # ── Left panel: Gender split ──────────────────────────────────────
    add_rect(slide, 0.45, 1.35, 3.9, 5.35, fill_rgb=BG_BLACK, rounding=12000)
    add_textbox(slide, 0.65, 1.45, 3.5, 0.4,
                "Gender Split", font_size=11, font_color=GOLD, bold=True)
    add_line(slide, 0.65, 1.9, 4.25, 1.9, color_rgb=GOLD, width_pt=0.5)

    genders = [("Male", "__ %", GOLD), ("Female", "__ %", GREY_LT)]
    for i, (label, pct, color) in enumerate(genders):
        y = 2.05 + i * 1.4
        add_textbox(slide, 0.7, y, 3.4, 0.4,
                    label, font_size=10, font_color=GREY_MID)
        add_textbox(slide, 0.7, y + 0.4, 3.4, 0.75,
                    pct, font_name="Aptos Display", font_size=22,
                    font_color=color, bold=True)
        # Bar placeholder
        add_rect(slide, 0.7, y + 1.2, 3.1, 0.18, fill_rgb=GREY_MID, rounding=8000)
        add_rect(slide, 0.7, y + 1.2, 1.55, 0.18, fill_rgb=color, rounding=8000)

    # Avg ratings per gender
    add_textbox(slide, 0.65, 5.05, 3.5, 0.35,
                "Avg. Ratings / User", font_size=9, font_color=GREY_MID)
    add_textbox(slide, 0.65, 5.4, 1.5, 0.55,
                "___", font_size=16, font_color=GOLD, bold=True)
    add_textbox(slide, 0.65, 5.7, 1.5, 0.3,
                "Male", font_size=8, font_color=GREY_MID)
    add_textbox(slide, 2.3, 5.4, 1.5, 0.55,
                "___", font_size=16, font_color=GREY_LT, bold=True)
    add_textbox(slide, 2.3, 5.7, 1.5, 0.3,
                "Female", font_size=8, font_color=GREY_MID)

    # ── Middle panel: Age Groups ──────────────────────────────────────
    add_rect(slide, 4.62, 1.35, 4.0, 5.35, fill_rgb=BG_BLACK, rounding=12000)
    add_textbox(slide, 4.82, 1.45, 3.6, 0.4,
                "Age Group Breakdown", font_size=11, font_color=GOLD, bold=True)
    add_line(slide, 4.82, 1.9, 8.5, 1.9, color_rgb=GOLD, width_pt=0.5)

    age_groups = [
        "Under 18", "18 – 24", "25 – 34",
        "35 – 44", "45 – 49", "50 – 55", "56+"
    ]
    bar_colors = [GREY_MID, GOLD_WARM, GOLD, GOLD_WARM, GREY_LT, GREY_MID, GREY_MID]
    for i, (age, color) in enumerate(zip(age_groups, bar_colors)):
        y = 2.05 + i * 0.7
        add_textbox(slide, 4.82, y, 1.3, 0.35,
                    age, font_size=9, font_color=GREY_LT)
        # Bar background
        add_rect(slide, 6.2, y + 0.05, 2.0, 0.22, fill_rgb=BG_DARK, rounding=5000)
        # Bar fill placeholder (varying width to suggest chart)
        bar_w = [0.2, 0.9, 1.5, 0.85, 0.4, 0.25, 0.1][i]
        add_rect(slide, 6.2, y + 0.05, bar_w, 0.22, fill_rgb=color, rounding=5000)
        add_textbox(slide, 8.3, y, 0.5, 0.35,
                    "__ %", font_size=9, font_color=GREY_MID)

    # ── Right panel: Top Occupations ─────────────────────────────────
    add_rect(slide, 8.88, 1.35, 4.0, 5.35, fill_rgb=BG_BLACK, rounding=12000)
    add_textbox(slide, 9.08, 1.45, 3.6, 0.4,
                "Top Occupations", font_size=11, font_color=GOLD, bold=True)
    add_line(slide, 9.08, 1.9, 12.75, 1.9, color_rgb=GOLD, width_pt=0.5)

    occs = ["Student", "Other / None", "Educator", "Engineer",
            "Programmer", "Executive", "Scientist"]
    for i, occ in enumerate(occs):
        y = 2.05 + i * 0.65
        add_textbox(slide, 9.08, y, 2.2, 0.35,
                    occ, font_size=9, font_color=GREY_LT)
        add_rect(slide, 11.35, y + 0.06, 1.3, 0.2, fill_rgb=BG_DARK, rounding=5000)
        bar_w = [1.3, 0.75, 0.55, 0.52, 0.4, 0.28, 0.26][i]
        add_rect(slide, 11.35, y + 0.06, bar_w, 0.2, fill_rgb=GOLD, rounding=5000)
        add_textbox(slide, 12.68, y, 0.5, 0.35,
                    "__ %", font_size=9, font_color=GREY_MID)

    return slide


def build_content_slide(prs):
    """Slide 5 — Content & Catalogue Performance."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 5)

    section_label(slide, 0.4, 0.45, 10, "Content & Catalogue Performance")

    # ── Top row: 4 KPI tiles ──────────────────────────────────────────
    top_kpis = [
        ("Movies in Catalogue", "______", "titles"),
        ("Top Genre (Volume)",  "______", "by ratings"),
        ("Top Genre (Rating)",  "__ / 5", "avg score"),
        ("Film-Noir Avg Score", "__ / 5", "niche peak"),
    ]
    for i, (label, val, sub) in enumerate(top_kpis):
        x = 0.45 + i * 3.22
        add_rect(slide, x, 1.3, 3.0, 1.55, fill_rgb=BG_BLACK, rounding=12000)
        add_rect(slide, x, 1.3, 3.0, 0.06, fill_rgb=GOLD)
        add_textbox(slide, x + 0.15, 1.4, 2.7, 0.35,
                    label, font_size=9, font_color=GREY_MID)
        add_textbox(slide, x + 0.15, 1.75, 2.7, 0.65,
                    val, font_name="Aptos Display", font_size=20,
                    font_color=GOLD, bold=True)
        add_textbox(slide, x + 0.15, 2.45, 2.7, 0.3,
                    sub, font_size=8, font_color=GREY_MID, italic=True)

    # ── Genre table (left) ────────────────────────────────────────────
    add_rect(slide, 0.45, 3.05, 6.5, 3.65, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 0.65, 3.15, 6.1, 0.4,
                "Genre Performance Table", font_size=11,
                font_color=GOLD, bold=True)

    headers = ["Genre", "Ratings", "% Share", "Avg Score"]
    hx = [0.65, 2.8, 4.15, 5.5]
    hw = [2.0, 1.25, 1.25, 1.35]
    for j, (hd, hx_, hw_) in enumerate(zip(headers, hx, hw)):
        add_textbox(slide, hx_, 3.6, hw_, 0.32,
                    hd, font_size=9, font_color=GOLD, bold=True)

    add_line(slide, 0.65, 3.95, 6.8, 3.95, color_rgb=GOLD, width_pt=0.4)

    genres = ["Drama", "Comedy", "Action", "Thriller",
              "Romance", "Sci-Fi", "Crime", "Film-Noir"]
    for i, g in enumerate(genres):
        y = 4.05 + i * 0.31
        bg = BG_DARK if i % 2 == 0 else BG_BLACK
        add_textbox(slide, 0.65, y, 2.0, 0.28, g, font_size=9, font_color=GREY_LT)
        for j, (hx_, hw_) in enumerate(zip(hx[1:], hw[1:]), 1):
            add_textbox(slide, hx_, y, hw_, 0.28,
                        "______", font_size=9, font_color=GREY_MID)

    # ── Rating distribution (right) ───────────────────────────────────
    add_rect(slide, 7.22, 3.05, 5.66, 3.65, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 7.42, 3.15, 5.2, 0.4,
                "Rating Score Distribution", font_size=11,
                font_color=GOLD, bold=True)

    scores = [("5 ⭐", GOLD, 0.85), ("4 ⭐", GOLD_WARM, 0.65),
              ("3 ⭐", GREY_LT, 0.45), ("2 ⭐", GREY_MID, 0.28), ("1 ⭐", GREY_MID, 0.15)]
    max_bar = 4.5
    for i, (label, color, ratio) in enumerate(scores):
        y = 3.65 + i * 0.58
        add_textbox(slide, 7.42, y, 0.65, 0.32,
                    label, font_size=9, font_color=GREY_LT)
        add_rect(slide, 8.15, y + 0.06, max_bar, 0.2, fill_rgb=BG_DARK, rounding=5000)
        add_rect(slide, 8.15, y + 0.06, max_bar * ratio, 0.2,
                 fill_rgb=color, rounding=5000)
        add_textbox(slide, 12.72, y, 0.55, 0.32,
                    "__ %", font_size=9, font_color=GREY_MID)

    return slide


def build_marketing_slide(prs):
    """Slide 6 — Marketing & Genre Insights."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 6)

    section_label(slide, 0.4, 0.45, 10, "Marketing & Genre Insights")

    # ── Funnel (left) ─────────────────────────────────────────────────
    add_rect(slide, 0.45, 1.3, 5.9, 5.4, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 0.65, 1.4, 5.5, 0.4,
                "User Conversion Funnel", font_size=11, font_color=GOLD, bold=True)

    funnel_stages = [
        ("Visitors (est.)",         "100%",  5.2, GOLD),
        ("Registered Users",        "__ %",  4.2, GOLD_WARM),
        ("Completed Profile",       "__ %",  3.2, GREY_LT),
        ("Active (≥ 20 ratings)",   "__ %",  2.2, GREY_LT),
        ("Power Users (≥ 500 rtg)", "__ %",  1.2, GREY_MID),
    ]
    for i, (label, pct, bar_w, color) in enumerate(funnel_stages):
        y = 1.95 + i * 0.85
        add_rect(slide, 0.65, y, bar_w, 0.5, fill_rgb=color, rounding=8000)
        add_textbox(slide, 0.75, y + 0.08, 4.8, 0.35,
                    label, font_size=9, font_color=BG_DARK, bold=True)
        add_textbox(slide, 5.5, y + 0.08, 0.75, 0.35,
                    pct, font_size=9, font_color=GREY_MID)

    # ── User Growth Trend (right) ─────────────────────────────────────
    add_rect(slide, 6.62, 1.3, 6.26, 2.55, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 6.82, 1.4, 5.9, 0.4,
                "Monthly New Users — Trend", font_size=11, font_color=GOLD, bold=True)
    add_textbox(slide, 6.82, 1.88, 5.9, 0.4,
                "[ Chart Placeholder — Monthly New User Signups ]",
                font_size=9, font_color=GREY_MID, italic=True)
    # Simulated line chart bars
    bar_heights = [0.9, 0.6, 0.5, 0.65, 0.55, 0.75, 0.7, 0.78, 0.95, 0.76, 0.7,
                   0.65, 0.55, 0.5, 0.45, 0.37, 0.3]
    bar_w = 0.27
    for i, bh in enumerate(bar_heights):
        x = 6.82 + i * (bar_w + 0.03)
        if x + bar_w > 12.7:
            break
        color = GOLD if bh == max(bar_heights) else GREY_MID
        add_rect(slide, x, 2.38 + (0.95 - bh), bar_w, bh,
                 fill_rgb=color, rounding=5000)

    # Axis labels
    add_textbox(slide, 6.82, 3.4, 6.0, 0.3,
                "Apr '00          Jun '00          Sep '00          Dec '00          Mar '01",
                font_size=7, font_color=GREY_MID)

    # ── Key Marketing KPIs (right bottom) ────────────────────────────
    add_rect(slide, 6.62, 4.05, 6.26, 2.6, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 6.82, 4.15, 5.9, 0.4,
                "Marketing KPIs", font_size=11, font_color=GOLD, bold=True)

    mkt_kpis = [
        ("Peak Acquisition Month",  "______"),
        ("Peak New Users (month)",  "______"),
        ("Top Genre by Volume",     "______"),
        ("Top Genre by Avg Score",  "______"),
        ("Conversion Rate",         "__ %"),
        ("Power User Share",        "__ %"),
    ]
    for i, (label, val) in enumerate(mkt_kpis):
        y = 4.65 + i * 0.33
        alt = i % 2 == 0
        add_textbox(slide, 6.82, y, 4.2, 0.3,
                    label, font_size=9, font_color=GREY_LT)
        add_textbox(slide, 11.05, y, 1.6, 0.3,
                    val, font_size=9, font_color=GOLD, bold=True,
                    align=PP_ALIGN.RIGHT)

    return slide


def build_engagement_slide(prs):
    """Slide 7 — Engagement & Retention."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 7)

    section_label(slide, 0.4, 0.45, 10, "Engagement & Retention")

    # ── Top KPI row ───────────────────────────────────────────────────
    eng_kpis = [
        ("DAU (peak)",         "______"),
        ("MAU (avg)",          "______"),
        ("DAU / MAU Ratio",    "__ %"),
        ("M1→M2 Retention",    "__ %"),
        ("Monthly Churn",      "__ %"),
        ("NPS (est.)",         "+__"),
    ]
    for i, (label, val) in enumerate(eng_kpis):
        x = 0.45 + i * 2.08
        add_rect(slide, x, 1.3, 1.95, 1.4, fill_rgb=BG_BLACK, rounding=10000)
        add_rect(slide, x, 1.3, 1.95, 0.06, fill_rgb=GOLD)
        add_textbox(slide, x + 0.1, 1.4, 1.75, 0.35,
                    label, font_size=8, font_color=GREY_MID)
        add_textbox(slide, x + 0.1, 1.78, 1.75, 0.65,
                    val, font_name="Aptos Display", font_size=18,
                    font_color=GOLD, bold=True)

    # ── User Cohort Panel (left) ──────────────────────────────────────
    add_rect(slide, 0.45, 2.9, 5.9, 3.8, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 0.65, 3.0, 5.5, 0.4,
                "User Activity Cohorts", font_size=11, font_color=GOLD, bold=True)

    cohorts = [
        ("Casual",      "20 – 49 ratings",    GREY_MID,  0.8),
        ("Engaged",     "50 – 99 ratings",    GREY_LT,   0.65),
        ("Regular",     "100 – 199 ratings",  GREY_LT,   0.75),
        ("Active",      "200 – 499 ratings",  GOLD_WARM, 0.55),
        ("Power User",  "500 – 999 ratings",  GOLD,      0.4),
        ("Super User",  "1,000+ ratings",     GOLD,      0.15),
    ]
    for i, (seg, rng, color, ratio) in enumerate(cohorts):
        y = 3.5 + i * 0.5
        add_textbox(slide, 0.65, y, 1.8, 0.38,
                    seg, font_size=9, font_color=GREY_LT, bold=True)
        add_textbox(slide, 2.5, y, 1.6, 0.38,
                    rng, font_size=8, font_color=GREY_MID)
        add_rect(slide, 4.2, y + 0.08, 1.8, 0.2, fill_rgb=BG_DARK, rounding=5000)
        add_rect(slide, 4.2, y + 0.08, 1.8 * ratio, 0.2,
                 fill_rgb=color, rounding=5000)
        add_textbox(slide, 6.1, y, 0.55, 0.38,
                    "__ %", font_size=9, font_color=GREY_MID)

    # ── Retention & Churn Panel (right) ──────────────────────────────
    add_rect(slide, 6.62, 2.9, 6.26, 3.8, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 6.82, 3.0, 5.9, 0.4,
                "Retention Cohort Snapshot", font_size=11, font_color=GOLD, bold=True)

    add_textbox(slide, 6.82, 3.5, 5.9, 0.35,
                "[ Chart Placeholder — Retention by Month ]",
                font_size=9, font_color=GREY_MID, italic=True)

    # Simulated retention curve
    points = [1.0, 0.68, 0.55, 0.48, 0.44, 0.41, 0.39, 0.38, 0.37, 0.36, 0.35, 0.34]
    seg_w = 0.46
    for i, p in enumerate(points[:-1]):
        x1 = 6.82 + i * seg_w
        x2 = 6.82 + (i + 1) * seg_w
        y1 = 3.95 + (1 - p) * 1.5
        y2 = 3.95 + (1 - points[i + 1]) * 1.5
        # Draw as thin filled rects approximating the curve
        bar_h = max(p * 0.12, 0.04)
        add_rect(slide, x1, y1, seg_w - 0.02, bar_h,
                 fill_rgb=GOLD if i == 0 else GREY_LT, rounding=3000)

    retention_labels = [
        ("M1 → M2",  "68%",  GOLD),
        ("M1 → M6",  "41%",  GOLD_WARM),
        ("M1 → M12", "35%",  GREY_LT),
        ("Churn/mo", "4%",   GREY_MID),
    ]
    for i, (label, val, color) in enumerate(retention_labels):
        x = 6.82 + i * 1.55
        add_textbox(slide, x, 5.75, 1.4, 0.3,
                    label, font_size=8, font_color=GREY_MID)
        add_textbox(slide, x, 6.05, 1.4, 0.45,
                    val, font_name="Aptos Display", font_size=16,
                    font_color=color, bold=True)

    return slide


def build_top_movies_slide(prs):
    """Slide 8 — Top Content Highlights."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 8)

    section_label(slide, 0.4, 0.45, 10, "Top Content Highlights")
    add_textbox(slide, 0.55, 0.95, 12, 0.3,
                "Highest rated & most reviewed titles (min. 1,000 ratings)",
                font_size=9, font_color=GREY_MID, italic=True)

    # ── Top Rated (left table) ────────────────────────────────────────
    add_rect(slide, 0.45, 1.35, 6.1, 5.35, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 0.65, 1.45, 5.7, 0.4,
                "Top Rated Titles", font_size=11, font_color=GOLD, bold=True)

    t_headers = ["#", "Title", "Avg", "Ratings"]
    t_hx      = [0.65, 1.05, 4.65, 5.5]
    t_hw      = [0.35, 3.5, 0.75, 0.9]
    for j, (hd, hx_, hw_) in enumerate(zip(t_headers, t_hx, t_hw)):
        add_textbox(slide, hx_, 1.9, hw_, 0.3,
                    hd, font_size=9, font_color=GOLD, bold=True)
    add_line(slide, 0.65, 2.22, 6.4, 2.22, color_rgb=GOLD, width_pt=0.4)

    top_titles = [
        "Schindler's List",
        "Shawshank Redemption, The",
        "Godfather, The",
        "Usual Suspects, The",
        "Rear Window",
        "Star Wars: Episode IV",
        "Casablanca",
        "Silence of the Lambs, The",
        "Raiders of the Lost Ark",
        "American Beauty",
    ]
    for i, title in enumerate(top_titles):
        y = 2.3 + i * 0.46
        add_textbox(slide, 0.65, y, 0.35, 0.38,
                    str(i + 1), font_size=9, font_color=GOLD, bold=True)
        add_textbox(slide, 1.05, y, 3.5, 0.38,
                    title, font_size=9, font_color=GREY_LT)
        add_textbox(slide, 4.65, y, 0.75, 0.38,
                    "__.___", font_size=9, font_color=GOLD_WARM)
        add_textbox(slide, 5.5, y, 0.9, 0.38,
                    "_,___", font_size=9, font_color=GREY_MID)

    # ── Most Reviewed (right table) ───────────────────────────────────
    add_rect(slide, 6.8, 1.35, 6.1, 5.35, fill_rgb=BG_BLACK, rounding=10000)
    add_textbox(slide, 7.0, 1.45, 5.7, 0.4,
                "Most Reviewed Titles", font_size=11, font_color=GOLD, bold=True)

    for j, (hd, offset) in enumerate(zip(t_headers, [0, 0.4, 3.5, 4.35])):
        add_textbox(slide, 7.0 + offset, 1.9, 0.9, 0.3,
                    hd, font_size=9, font_color=GOLD, bold=True)
    add_line(slide, 7.0, 2.22, 12.75, 2.22, color_rgb=GOLD, width_pt=0.4)

    most_reviewed = [
        "American Beauty",
        "Star Wars: Episode IV",
        "Star Wars: Episode V",
        "Silence of the Lambs",
        "Fargo",
        "Saving Private Ryan",
        "Schindler's List",
        "Shawshank Redemption",
        "Pulp Fiction",
        "Toy Story",
    ]
    for i, title in enumerate(most_reviewed):
        y = 2.3 + i * 0.46
        add_textbox(slide, 7.0, y, 0.35, 0.38,
                    str(i + 1), font_size=9, font_color=GOLD, bold=True)
        add_textbox(slide, 7.4, y, 3.5, 0.38,
                    title, font_size=9, font_color=GREY_LT)
        add_textbox(slide, 10.95, y, 0.75, 0.38,
                    "__.___", font_size=9, font_color=GOLD_WARM)
        add_textbox(slide, 11.75, y, 0.95, 0.38,
                    "_,___", font_size=9, font_color=GREY_MID)

    return slide


def build_takeaways_slide(prs):
    """Slide 9 — Key Takeaways & Next Steps (mirrors Evgeniya's outcome slide style)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)
    gold_bar_left(slide, h=6.92)
    footer_bar(slide, 9)

    section_label(slide, 0.4, 0.45, 10, "Key Takeaways & Next Steps")

    takeaway_items = [
        ("01.", "User Base", "Platform has __,___ registered users. Core segment: ages 25–34 (__ %)."),
        ("02.", "Engagement", "DAU/MAU ratio at __ %. M1→M2 retention at __ %. Churn stable at __ %."),
        ("03.", "Content", "__ genres catalogued. Drama leads volume. Film-Noir leads quality score."),
        ("04.", "Marketing", "Peak acquisition in ______. Top conversion bottleneck at ______ stage."),
        ("05.", "Opportunity", "Power users (__ %) drive outsized rating volume. Upsell / loyalty lever."),
    ]

    for i, (num, title, body) in enumerate(takeaway_items):
        row = i % 3
        col = i // 3
        x = 0.45 + col * 6.6
        y = 1.35 + row * 1.75

        add_rect(slide, x, y, 6.25, 1.5, fill_rgb=BG_BLACK, rounding=12000)

        # Gold number
        add_textbox(slide, x + 0.18, y + 0.15, 0.85, 1.1,
                    num, font_name="Aptos Display", font_size=24,
                    font_color=GOLD, bold=True)

        # Title
        add_textbox(slide, x + 1.1, y + 0.15, 5.0, 0.42,
                    title, font_name="Aptos Display", font_size=13,
                    font_color=WHITE, bold=True)

        # Body text
        add_textbox(slide, x + 1.1, y + 0.6, 5.0, 0.8,
                    body, font_size=9, font_color=GREY_LT, word_wrap=True)

    # ── Next Steps box ────────────────────────────────────────────────
    add_rect(slide, 0.45, 6.6, 0.0, 0.0)   # spacer
    add_rect(slide, 0.45, 6.3, 12.44, 0.45, fill_rgb=BG_BLACK, rounding=8000)
    add_rect(slide, 0.45, 6.3, 0.06, 0.45, fill_rgb=GOLD)
    add_textbox(slide, 0.6, 6.33, 12.0, 0.38,
                "Next Steps:  [ fill in action items, owners, and deadlines ]",
                font_size=9, font_color=GREY_MID, italic=True)

    return slide


def build_thank_you_slide(prs):
    """Slide 10 — Thank You (mirrors Evgeniya's closing slide)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, BG_DARK)

    # Three decorative ovals (Evgeniya signature)
    for i, (ox, oy, ow, oh, alpha) in enumerate([
        (0.3,  0.2,  4.5, 4.5, BG_BLACK),
        (9.0,  3.0,  4.5, 4.5, BG_BLACK),
        (5.5,  2.5,  3.0, 3.0, RGBColor(0x15, 0x18, 0x35)),
    ]):
        ov = slide.shapes.add_shape(9, Inches(ox), Inches(oy), Inches(ow), Inches(oh))
        ov.fill.solid()
        ov.fill.fore_color.rgb = alpha
        ov.line.fill.background()

    add_textbox(slide, 3.0, 2.8, 7.5, 1.0,
                "THANK YOU",
                font_name="Aptos Display", font_size=48,
                font_color=WHITE, bold=True, align=PP_ALIGN.CENTER)

    add_textbox(slide, 3.0, 4.0, 7.5, 0.5,
                "Presented by:", font_size=11,
                font_color=GREY_MID, align=PP_ALIGN.CENTER)
    add_textbox(slide, 3.0, 4.5, 7.5, 0.5,
                "[Your Name]  ·  WatchIt Management Update",
                font_size=12, font_color=GOLD,
                bold=True, align=PP_ALIGN.CENTER)

    gold_bar_bottom(slide)

    return slide


# ─────────────────────────────────────────────────────────────────────────────
# ASSEMBLE
# ─────────────────────────────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.50)

build_title_slide(prs)
build_agenda_slide(prs)
build_kpi_overview_slide(prs)
build_users_slide(prs)
build_content_slide(prs)
build_marketing_slide(prs)
build_engagement_slide(prs)
build_top_movies_slide(prs)
build_takeaways_slide(prs)
build_thank_you_slide(prs)

output = "/home/Kisa/CorpusInteraction/WatchIt_Management_Metrics_Template.pptx"
prs.save(output)
print(f"✅  Saved: {output}  ({len(prs.slides)} slides)")
