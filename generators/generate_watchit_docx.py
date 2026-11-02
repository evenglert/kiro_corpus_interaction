"""
WatchIt — Monthly Management Update Word Template Generator
Design: matches the WatchIt deck palette (navy / gold / Aptos)
Max 2 pages, narrative prose with structured sections.
Placeholders use [  ] brackets — easy to search/replace programmatically.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import copy

# ── Palette ──────────────────────────────────────────────────────────────────
C_NAVY   = RGBColor(0x0A, 0x13, 0x2E)   # deep navy   — header bg
C_BLACK  = RGBColor(0x0F, 0x0A, 0x05)   # near-black  — sub-header bg
C_GOLD   = RGBColor(0xDC, 0xA5, 0x23)   # gold        — accent
C_AMBER  = RGBColor(0xF8, 0x9F, 0x19)   # amber       — secondary accent
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_LGREY  = RGBColor(0xD9, 0xD9, 0xD9)   # light grey  — alt rows
C_MGREY  = RGBColor(0xA6, 0xA6, 0xA6)   # mid grey    — meta text
C_DKGREY = RGBColor(0x2D, 0x2D, 0x2D)   # dark grey   — body text

FONT_HEADING = "Aptos Display"
FONT_BODY    = "Aptos"

# ── Low-level XML helpers ────────────────────────────────────────────────────

def rgb_hex(rgb: RGBColor) -> str:
    """Return uppercase 6-char hex string from a python-docx RGBColor."""
    return str(rgb).upper()


def set_cell_bg(cell, rgb: RGBColor):
    """Set table cell background shading."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), rgb_hex(rgb))
    tcPr.append(shd)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom),
                      ('left', left), ('right', right)]:
        if val:
            el = OxmlElement(f'w:{side}')
            el.set(qn('w:val'), val.get('val', 'single'))
            el.set(qn('w:sz'), str(val.get('sz', 4)))
            el.set(qn('w:color'), val.get('color', 'auto'))
            tcBorders.append(el)
    tcPr.append(tcBorders)


def no_border(cell):
    """Remove all borders from a cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'none')
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_para_border_bottom(para, color="DCA523", sz=12):
    """Add a bottom border (rule) under a paragraph."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def set_para_shading(para, rgb: RGBColor):
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), rgb_hex(rgb))
    pPr.append(shd)


def cell_vmid(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), 'center')
    tcPr.append(vAlign)


def set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)


# ── Style helpers ─────────────────────────────────────────────────────────────

def fmt_run(run, font_name=FONT_BODY, size_pt=10, bold=False, italic=False,
            color: RGBColor = C_DKGREY):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_para(doc, text="", style=None, align=WD_ALIGN_PARAGRAPH.LEFT,
             space_before=0, space_after=4):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(13)
    if text:
        run = p.add_run(text)
        fmt_run(run)
    return p


def add_heading_block(doc, title: str, subtitle: str = ""):
    """Dark navy band with gold title text — section header."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Inches(0)
    set_para_shading(p, C_NAVY)

    run = p.add_run(f"  {title}")
    run.font.name  = FONT_HEADING
    run.font.size  = Pt(11)
    run.font.bold  = True
    run.font.color.rgb = C_GOLD

    if subtitle:
        run2 = p.add_run(f"  —  {subtitle}")
        run2.font.name   = FONT_HEADING
        run2.font.size   = Pt(9)
        run2.font.bold   = False
        run2.font.color.rgb = C_LGREY
    return p


def add_section_rule(doc, color="DCA523"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    set_para_border_bottom(p, color=color, sz=8)
    return p


def inline_gold(para, text):
    """Add a gold bold inline run to an existing paragraph."""
    run = para.add_run(text)
    fmt_run(run, bold=True, color=C_GOLD, size_pt=10)
    return run


def inline_body(para, text, bold=False, italic=False):
    run = para.add_run(text)
    fmt_run(run, bold=bold, italic=italic, size_pt=10)
    return run


# ── KPI bar table (1 row, N cells) ───────────────────────────────────────────

def add_kpi_bar(doc, kpis):
    """
    kpis: list of (label, placeholder) tuples
    Renders as a borderless table row — KPI tiles.
    """
    n = len(kpis)
    table = doc.add_table(rows=2, cols=n)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'

    for j, (label, value) in enumerate(kpis):
        # Value cell (row 0)
        cell_v = table.cell(0, j)
        set_cell_bg(cell_v, C_NAVY)
        no_border(cell_v)
        cell_v.width = Inches(6.8 / n)
        p_v = cell_v.paragraphs[0]
        p_v.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_v.paragraph_format.space_before = Pt(3)
        p_v.paragraph_format.space_after  = Pt(1)
        r_v = p_v.add_run(value)
        r_v.font.name  = FONT_HEADING
        r_v.font.size  = Pt(14)
        r_v.font.bold  = True
        r_v.font.color.rgb = C_GOLD

        # Label cell (row 1)
        cell_l = table.cell(1, j)
        set_cell_bg(cell_l, C_BLACK)
        no_border(cell_l)
        p_l = cell_l.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.paragraph_format.space_before = Pt(1)
        p_l.paragraph_format.space_after  = Pt(3)
        r_l = p_l.add_run(label)
        r_l.font.name   = FONT_BODY
        r_l.font.size   = Pt(7.5)
        r_l.font.bold   = False
        r_l.font.color.rgb = C_LGREY

    # Gold top border on entire first row
    for j in range(n):
        set_cell_borders(
            table.cell(0, j),
            top={'val': 'single', 'sz': 12, 'color': 'DCA523'}
        )
    return table


# ── Metrics mini-table ────────────────────────────────────────────────────────

def add_metrics_table(doc, rows_data, col_widths_cm=None):
    """
    rows_data: list of (label, value, note) tuples.
    Alternating row shading matching the deck aesthetic.
    """
    n_cols = len(rows_data[0])
    table = doc.add_table(rows=len(rows_data), cols=n_cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    for i, row_data in enumerate(rows_data):
        alt = i % 2 == 0
        bg  = C_NAVY if alt else C_BLACK
        for j, cell_text in enumerate(row_data):
            cell = table.cell(i, j)
            set_cell_bg(cell, bg)
            no_border(cell)
            cell_vmid(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after  = Pt(1)
            p.paragraph_format.left_indent  = Cm(0.15)
            is_value = (j == 1)
            r = p.add_run(str(cell_text))
            r.font.name  = FONT_BODY if not is_value else FONT_HEADING
            r.font.size  = Pt(8.5)
            r.font.bold  = is_value
            r.font.color.rgb = C_GOLD if is_value else C_LGREY

    if col_widths_cm:
        for j, w in enumerate(col_widths_cm):
            set_col_width(table, j, w)
    return table


# ══════════════════════════════════════════════════════════════════════════════
# BUILD DOCUMENT
# ══════════════════════════════════════════════════════════════════════════════

doc = Document()

# ── Page margins (tight to fit 2 pages) ──────────────────────────────────────
section = doc.sections[0]
section.top_margin    = Cm(1.4)
section.bottom_margin = Cm(1.4)
section.left_margin   = Cm(1.8)
section.right_margin  = Cm(1.8)
section.page_width    = Cm(21.0)   # A4
section.page_height   = Cm(29.7)

# ══════════════════════════════════════════════════════════════════════════════
# HEADER — Document masthead
# ══════════════════════════════════════════════════════════════════════════════
p_logo = doc.add_paragraph()
p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_logo.paragraph_format.space_before = Pt(0)
p_logo.paragraph_format.space_after  = Pt(0)
set_para_shading(p_logo, C_NAVY)
r1 = p_logo.add_run("  🎬  WatchIt")
r1.font.name  = FONT_HEADING
r1.font.size  = Pt(18)
r1.font.bold  = True
r1.font.color.rgb = C_GOLD

r2 = p_logo.add_run("    Monthly Management Update")
r2.font.name  = FONT_HEADING
r2.font.size  = Pt(11)
r2.font.bold  = False
r2.font.color.rgb = C_WHITE

p_meta = doc.add_paragraph()
p_meta.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_meta.paragraph_format.space_before = Pt(0)
p_meta.paragraph_format.space_after  = Pt(6)
set_para_shading(p_meta, C_BLACK)
r_m = p_meta.add_run(
    "  Reporting Period: [ Month Year ]   ·   Prepared by: [ Name ]"
    "   ·   Audience: Management / Board"
)
r_m.font.name   = FONT_BODY
r_m.font.size   = Pt(8)
r_m.font.italic = True
r_m.font.color.rgb = C_MGREY

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — EXECUTIVE SUMMARY
# ══════════════════════════════════════════════════════════════════════════════
add_heading_block(doc, "01  EXECUTIVE SUMMARY")

p_summary = add_para(doc, space_before=3, space_after=4)
inline_body(p_summary,
    "WatchIt closed [ Month ] with ")
inline_gold(p_summary, "[ __,___ ] registered users")
inline_body(p_summary,
    " and a total catalogue of ")
inline_gold(p_summary, "[ _,___ ] titles")
inline_body(p_summary,
    ". The platform recorded ")
inline_gold(p_summary, "[ _,___,___ ] cumulative ratings")
inline_body(p_summary,
    ", reflecting a platform-wide average score of ")
inline_gold(p_summary, "[ _.__ / 5 ]")
inline_body(p_summary,
    ". Monthly active users stood at ")
inline_gold(p_summary, "[ __,___ ]")
inline_body(p_summary,
    " ([ __ ]% of the registered base), with a DAU/MAU ratio of ")
inline_gold(p_summary, "[ __ ]%")
inline_body(p_summary,
    " — consistent with a healthy, engaged community. "
    "Month-one retention held at ")
inline_gold(p_summary, "[ __ ]%")
inline_body(p_summary,
    ", and the estimated Net Promoter Score remains at ")
inline_gold(p_summary, "[ +__ ]")
inline_body(p_summary, ".")

# Key points box (bordered)
p_kp_label = add_para(doc, space_before=2, space_after=1)
r_kp = p_kp_label.add_run("  KEY POINTS THIS PERIOD")
r_kp.font.name  = FONT_HEADING
r_kp.font.size  = Pt(8)
r_kp.font.bold  = True
r_kp.font.color.rgb = C_WHITE
set_para_shading(p_kp_label, C_GOLD)

key_points = [
    "[ Highlight 1 — e.g. user growth +X% MoM / new milestone reached ]",
    "[ Highlight 2 — e.g. top content genre shift or catalogue expansion ]",
    "[ Highlight 3 — e.g. engagement or retention movement ]",
    "[ Highlight 4 — e.g. marketing / acquisition insight ]",
]
for kp in key_points:
    p_item = add_para(doc, space_before=0, space_after=1)
    set_para_shading(p_item, C_NAVY)
    r_bullet = p_item.add_run("  ▸  ")
    r_bullet.font.name  = FONT_BODY
    r_bullet.font.size  = Pt(9)
    r_bullet.font.bold  = True
    r_bullet.font.color.rgb = C_GOLD
    r_text = p_item.add_run(kp)
    r_text.font.name   = FONT_BODY
    r_text.font.size   = Pt(9)
    r_text.font.color.rgb = C_LGREY

add_para(doc, space_before=2, space_after=2)   # spacer

# ── KPI Bar ───────────────────────────────────────────────────────────────────
add_kpi_bar(doc, [
    ("Registered Users",    "[ __,___ ]"),
    ("Ratings Submitted",   "[ _,___,___ ]"),
    ("Movies in Catalogue", "[ _,___ ]"),
    ("Avg. Rating Score",   "[ _.__ / 5 ]"),
    ("DAU / MAU Ratio",     "[ __ % ]"),
    ("M1→M2 Retention",     "[ __ % ]"),
])

add_para(doc, space_before=4, space_after=2)   # spacer

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — USERS
# ══════════════════════════════════════════════════════════════════════════════
add_heading_block(doc, "02  USERS", "Who is on the platform")

p_users = add_para(doc, space_before=3, space_after=3)
inline_body(p_users,
    "Our user base is predominantly ")
inline_gold(p_users, "male ([ __ ]%)")
inline_body(p_users, " versus female ([ __ ]%). The largest age cohort remains ")
inline_gold(p_users, "[ Age Group ]")
inline_body(p_users, " at [ __ ]% of total registrations, followed by ")
inline_gold(p_users, "[ Age Group ]")
inline_body(p_users,
    " ([ __ ]%). This concentration signals a ")
inline_body(p_users, "core demographic skewing toward [ young-adult / professional ] viewers",
            italic=True)
inline_body(p_users,
    ". The dominant occupation is ")
inline_gold(p_users, "[ Occupation ]")
inline_body(p_users,
    " ([ __ ]%), suggesting strong presence of a digitally engaged, "
    "quality-conscious audience. Each user submits an average of ")
inline_gold(p_users, "[ ___ ] ratings")
inline_body(p_users,
    ", with power users (≥ 500 ratings) accounting for [ __ ]% of the base "
    "yet driving a disproportionate share of content signals.")

# Users metrics table
add_metrics_table(doc,
    rows_data=[
        ("Gender: Male",              "[ __ %  ]", "[ __,___ ] users"),
        ("Gender: Female",            "[ __ %  ]", "[ __,___ ] users"),
        ("Largest Age Group",         "[ 25–34 ]",  "[ __ ]% of total"),
        ("Avg. Ratings / User",       "[ ___._ ]",  "platform-wide"),
        ("Median Ratings / User",     "[ ___ ]",    "50th percentile"),
        ("Power Users (≥500 ratings)","[ __ %  ]",  "[ ___ ] users"),
    ],
    col_widths_cm=[6.3, 2.8, 8.0]
)

add_para(doc, space_before=4, space_after=2)   # spacer

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — USAGE & ENGAGEMENT
# ══════════════════════════════════════════════════════════════════════════════
add_heading_block(doc, "03  USAGE & ENGAGEMENT", "How the platform is being used")

p_usage = add_para(doc, space_before=3, space_after=3)
inline_body(p_usage,
    "Platform activity this period tells a story of ")
inline_gold(p_usage, "steady, quality-driven engagement")
inline_body(p_usage,
    ". Users are rating an average of ")
inline_gold(p_usage, "[ _._ ] titles per session")
inline_body(p_usage,
    ", and the distribution of scores continues to skew positive: ")
inline_gold(p_usage, "[ __ ]%")
inline_body(p_usage, " of all ratings are 4 or 5 stars, while only ")
inline_gold(p_usage, "[ __ ]%")
inline_body(p_usage,
    " are 1 or 2 stars. The top-rated genre this month is ")
inline_gold(p_usage, "[ Genre ]")
inline_body(p_usage, " (avg [ _.__ ]), while ")
inline_gold(p_usage, "[ Genre ]")
inline_body(p_usage,
    " leads in raw volume. Retention momentum is [ positive / stable / declining ]: "
    "month-one retention is ")
inline_gold(p_usage, "[ __ ]%")
inline_body(p_usage, " and six-month retention is ")
inline_gold(p_usage, "[ __ ]%")
inline_body(p_usage,
    ". Monthly churn remains at ")
inline_gold(p_usage, "[ __ ]%")
inline_body(p_usage, ", within the healthy range for a subscription-adjacent content platform.")

add_metrics_table(doc,
    rows_data=[
        ("Daily Active Users (DAU)",    "[ _,___ ]",  "peak this period"),
        ("Monthly Active Users (MAU)",  "[ _,___ ]",  "[ __ ]% of base"),
        ("DAU / MAU Ratio",             "[ __ %  ]",  "engagement health"),
        ("Avg. Session Depth",          "[ _._ ]",    "ratings per session"),
        ("Ratings: Score 4–5 Share",    "[ __ %  ]",  "positive sentiment"),
        ("Monthly Churn Rate",          "[ __ %  ]",  "steady-state"),
        ("Net Promoter Score (est.)",   "[ +__ ]",    "based on avg rating"),
    ],
    col_widths_cm=[6.3, 2.8, 8.0]
)

add_para(doc, space_before=4, space_after=2)   # spacer

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — CONTENT & MARKETING
# ══════════════════════════════════════════════════════════════════════════════
add_heading_block(doc, "04  CONTENT & MARKETING", "What users watch and how we grow")

p_content = add_para(doc, space_before=3, space_after=3)
inline_gold(p_content, "Drama")
inline_body(p_content,
    " continues to dominate audience engagement, representing ")
inline_gold(p_content, "[ __ ]%")
inline_body(p_content,
    " of all ratings this period, followed by ")
inline_gold(p_content, "Comedy ([ __ ]%)")
inline_body(p_content, " and ")
inline_gold(p_content, "Action ([ __ ]%)")
inline_body(p_content,
    ". On the quality front, niche genres such as ")
inline_gold(p_content, "Film-Noir")
inline_body(p_content,
    " and ")
inline_gold(p_content, "Documentary")
inline_body(p_content,
    " command the highest average scores ([ _.__ ] and [ _.__ ] respectively), "
    "signalling a quality-seeking audience segment worth targeting. "
    "The most reviewed title this period was ")
inline_gold(p_content, "[ Title ]")
inline_body(p_content, " ([ _,___ ] ratings, avg [ _.__ ]), while the highest-rated "
    "title (min. 1,000 ratings) was ")
inline_gold(p_content, "[ Title ]")
inline_body(p_content, " ([ _.__ ] / 5). "
    "New user acquisition peaked in ")
inline_gold(p_content, "[ Month ]")
inline_body(p_content,
    " with [ ___ ] sign-ups. The overall conversion rate from visitor to registered user "
    "is estimated at ")
inline_gold(p_content, "[ __ ]%")
inline_body(p_content, ".")

add_metrics_table(doc,
    rows_data=[
        ("Top Genre by Volume",           "[ Drama   ]", "[ __ ]% of all ratings"),
        ("Top Genre by Avg. Score",       "[ Film-Noir ]", "[ _.__ ] / 5"),
        ("Most Reviewed Title",           "[ Title   ]", "[ _,___ ] ratings"),
        ("Highest Rated Title",           "[ Title   ]", "[ _.__ ] / 5"),
        ("Peak Acquisition Month",        "[ Month   ]", "[ ___ ] new users"),
        ("Visitor → Registered Conv. %",  "[ __ %    ]", "estimated"),
    ],
    col_widths_cm=[6.3, 2.8, 8.0]
)

add_para(doc, space_before=4, space_after=2)   # spacer

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — STRATEGY RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════════════
add_heading_block(doc, "05  STRATEGY RECOMMENDATIONS")

recommendations = [
    (
        "Double down on Drama & Comedy catalogue depth.",
        "These two genres account for [ __ ]% of all ratings. "
        "Prioritise licensing titles in these categories to maximise immediate engagement uplift."
    ),
    (
        "Invest in niche quality content to raise the average score.",
        "Film-Noir and Documentary genres score [ __ ]+ / 5 despite low volume. "
        "A curated 'Critically Acclaimed' section could serve the quality-seeking segment "
        "and differentiate the platform."
    ),
    (
        "Focus retention on the 25–34 core cohort.",
        "This group is the largest ([ __ ]%) and most active. "
        "Personalisation and targeted re-engagement campaigns should prioritise this segment "
        "to reduce month-3 churn."
    ),
    (
        "Convert Casual users (20–49 ratings) to Engaged (50–99).",
        "[ __ ]% of users are in the Casual tier. A nudge programme — "
        "e.g. 'Rate 10 more titles to unlock your personalised watchlist' — "
        "could move a meaningful share up the engagement ladder."
    ),
]

for i, (title, body) in enumerate(recommendations, 1):
    p_rec = add_para(doc, space_before=2, space_after=1)
    set_para_shading(p_rec, C_NAVY)
    r_num = p_rec.add_run(f"  {i}.  ")
    r_num.font.name  = FONT_HEADING
    r_num.font.size  = Pt(10)
    r_num.font.bold  = True
    r_num.font.color.rgb = C_GOLD
    r_title = p_rec.add_run(title)
    r_title.font.name  = FONT_HEADING
    r_title.font.size  = Pt(10)
    r_title.font.bold  = True
    r_title.font.color.rgb = C_WHITE

    p_body = add_para(doc, space_before=0, space_after=2)
    set_para_shading(p_body, C_BLACK)
    p_body.paragraph_format.left_indent = Cm(0.7)
    r_b = p_body.add_run(f"  {body}")
    r_b.font.name   = FONT_BODY
    r_b.font.size   = Pt(8.5)
    r_b.font.color.rgb = C_LGREY

add_para(doc, space_before=3, space_after=1)   # spacer

# ══════════════════════════════════════════════════════════════════════════════
# FOOTER — gold rule + meta
# ══════════════════════════════════════════════════════════════════════════════
p_rule = add_para(doc, space_before=2, space_after=1)
set_para_border_bottom(p_rule, color="DCA523", sz=8)

p_footer = doc.add_paragraph()
p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_footer.paragraph_format.space_before = Pt(2)
p_footer.paragraph_format.space_after  = Pt(0)
r_f = p_footer.add_run(
    "WatchIt  ·  Monthly Management Update  ·  Confidential  ·  "
    "Data source: WatchIt_Metrics_Report.xlsx"
)
r_f.font.name   = FONT_BODY
r_f.font.size   = Pt(7)
r_f.font.italic = True
r_f.font.color.rgb = C_MGREY

# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
output = "/home/Kisa/CorpusInteraction/WatchIt_Monthly_Management_Update.docx"
doc.save(output)
print(f"✅  Saved: {output}")
