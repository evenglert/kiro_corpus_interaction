"""
WatchIt Organization — Usage Metrics Excel Generator
Based on MovieLens 1M dataset (Kaggle)
Source: https://www.kaggle.com/datasets/odedgolden/movielens-1m-dataset
  - 1,000,209 ratings · 6,040 users · ~3,883 movies · Apr 2000 – Feb 2003
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabel
import datetime

# ─── Palette ────────────────────────────────────────────────────────────────
BRAND_DARK   = "1A1A2E"   # deep navy
BRAND_MID    = "16213E"   # mid navy
BRAND_ACCENT = "E94560"   # brand red
BRAND_LIGHT  = "0F3460"   # steel blue
HEADER_BG    = "E94560"   # red headers
SUBHEAD_BG   = "0F3460"   # blue sub-headers
ALT_ROW      = "F5F5F5"   # light grey alternating row
WHITE        = "FFFFFF"
TEXT_DARK    = "1A1A2E"
TEXT_LIGHT   = "FFFFFF"

def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def header_font(bold=True, color=TEXT_LIGHT, size=11):
    return Font(name="Calibri", bold=bold, color=color, size=size)

def body_font(bold=False, color=TEXT_DARK, size=10):
    return Font(name="Calibri", bold=bold, color=color, size=size)

def thin_border():
    side = Side(style="thin", color="CCCCCC")
    return Border(left=side, right=side, top=side, bottom=side)

def center_align(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def left_align():
    return Alignment(horizontal="left", vertical="center")

def write_header(ws, row, col, value, span=1, bg=HEADER_BG, font_size=12):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(bg)
    cell.font = header_font(size=font_size, color=TEXT_LIGHT)
    cell.alignment = center_align()
    cell.border = thin_border()
    if span > 1:
        ws.merge_cells(
            start_row=row, start_column=col,
            end_row=row, end_column=col + span - 1
        )

def write_subheader(ws, row, col, value, span=1):
    write_header(ws, row, col, value, span=span, bg=SUBHEAD_BG, font_size=10)

def write_data(ws, row, col, value, alt=False, bold=False, number_fmt=None, align="center"):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = make_fill(ALT_ROW if alt else WHITE)
    cell.font = body_font(bold=bold)
    cell.border = thin_border()
    cell.alignment = center_align() if align == "center" else left_align()
    if number_fmt:
        cell.number_format = number_fmt
    return cell

def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width


# ════════════════════════════════════════════════════════════════════════════
# WORKBOOK
# ════════════════════════════════════════════════════════════════════════════
wb = openpyxl.Workbook()
wb.remove(wb.active)   # remove default sheet

# ────────────────────────────────────────────────────────────────────────────
# SHEET 0 – COVER
# ────────────────────────────────────────────────────────────────────────────
cover = wb.create_sheet("Cover")
cover.sheet_view.showGridLines = False
cover.sheet_properties.tabColor = BRAND_ACCENT

cover.merge_cells("B2:G2")
cover["B2"] = "🎬  WatchIt"
cover["B2"].font = Font("Calibri", size=36, bold=True, color=BRAND_ACCENT)
cover["B2"].alignment = center_align()

cover.merge_cells("B3:G3")
cover["B3"] = "Internal Platform Metrics Report"
cover["B3"].font = Font("Calibri", size=18, color=BRAND_LIGHT)
cover["B3"].alignment = center_align()

cover.merge_cells("B4:G4")
cover["B4"] = f"Generated: {datetime.date(2026, 11, 2).strftime('%B %d, %Y')}  ·  Data source: MovieLens 1M (Kaggle)"
cover["B4"].font = Font("Calibri", size=10, color="888888", italic=True)
cover["B4"].alignment = center_align()

# KPI summary tiles on cover (rows 6-9)
kpis = [
    ("Total Registered Users", "6,040"),
    ("Total Ratings Submitted", "1,000,209"),
    ("Movies in Catalogue", "3,883"),
    ("Avg. Rating Score", "3.58 / 5"),
    ("Avg. Ratings per User", "165.6"),
    ("Platform Active Since", "Apr 2000"),
]
kpi_labels_row, kpi_vals_row = 6, 7
for i, (label, val) in enumerate(kpis):
    col = 2 + i
    cover.merge_cells(start_row=kpi_labels_row, start_column=col, end_row=kpi_labels_row, end_column=col)
    cover.merge_cells(start_row=kpi_vals_row, start_column=col, end_row=kpi_vals_row, end_column=col)
    lc = cover.cell(row=kpi_labels_row, column=col, value=label)
    lc.fill = make_fill(BRAND_LIGHT)
    lc.font = Font("Calibri", size=9, bold=True, color=WHITE)
    lc.alignment = center_align(wrap=True)
    lc.border = thin_border()
    vc = cover.cell(row=kpi_vals_row, column=col, value=val)
    vc.fill = make_fill(BRAND_DARK)
    vc.font = Font("Calibri", size=14, bold=True, color=BRAND_ACCENT)
    vc.alignment = center_align()
    vc.border = thin_border()
    cover.column_dimensions[get_column_letter(col)].width = 20

cover.row_dimensions[2].height = 50
cover.row_dimensions[3].height = 30
cover.row_dimensions[4].height = 20
cover.row_dimensions[6].height = 40
cover.row_dimensions[7].height = 36

# Tabs navigator
nav_row = 9
cover.merge_cells(f"B{nav_row}:G{nav_row}")
cover.cell(row=nav_row, column=2, value="📊 Tabs: Users  |  Marketing  |  Top Movies  |  Engagement").font = Font(
    "Calibri", size=11, color=BRAND_LIGHT, italic=True
)
cover.cell(row=nav_row, column=2).alignment = center_align()


# ────────────────────────────────────────────────────────────────────────────
# SHEET 1 – USERS
# ────────────────────────────────────────────────────────────────────────────
ws_users = wb.create_sheet("Users")
ws_users.sheet_view.showGridLines = False
ws_users.sheet_properties.tabColor = BRAND_LIGHT

# ── Title ──
ws_users.merge_cells("A1:K1")
tc = ws_users["A1"]
tc.value = "WatchIt — User Analytics Dashboard"
tc.fill = make_fill(BRAND_DARK)
tc.font = Font("Calibri", size=16, bold=True, color=BRAND_ACCENT)
tc.alignment = center_align()
ws_users.row_dimensions[1].height = 32

ws_users.merge_cells("A2:K2")
sc = ws_users["A2"]
sc.value = "Source: MovieLens 1M Dataset  ·  6,040 registered users  ·  1,000,209 ratings  ·  Period: Apr 2000 – Feb 2003"
sc.fill = make_fill(BRAND_LIGHT)
sc.font = Font("Calibri", size=9, italic=True, color=WHITE)
sc.alignment = center_align()
ws_users.row_dimensions[2].height = 18

# ── Section A: Key User KPIs ──────────────────────────────────────────────
write_header(ws_users, 4, 1, "KEY USER KPIs", span=10, font_size=11)

kpi_data = [
    ("Metric", "Value", "Notes"),
    ("Total Registered Users", 6040, "All users with ≥ 20 ratings"),
    ("Total Ratings Submitted", 1_000_209, "Anonymous ratings, scale 1–5"),
    ("Movies in Catalogue", 3_883, "Distinct titles rated"),
    ("Average Ratings per User", 165.6, "1,000,209 / 6,040"),
    ("Average Rating Score", 3.58, "Platform-wide mean score"),
    ("Minimum Ratings per User", 20, "Eligibility threshold"),
    ("Maximum Ratings (single user)", 2314, "Power user ceiling"),
    ("Median Ratings per User", 96, "50th percentile activity"),
]
write_subheader(ws_users, 5, 1, "Metric")
write_subheader(ws_users, 5, 2, "Value")
write_subheader(ws_users, 5, 3, "Notes", span=8)
for i, (metric, val, note) in enumerate(kpi_data[1:], start=6):
    alt = (i % 2 == 0)
    write_data(ws_users, i, 1, metric, alt=alt, align="left")
    write_data(ws_users, i, 2, val, alt=alt, number_fmt="#,##0.##")
    merge_start = 3
    c = ws_users.cell(row=i, column=3, value=note)
    c.fill = make_fill(ALT_ROW if alt else WHITE)
    c.font = body_font()
    c.alignment = left_align()
    c.border = thin_border()
    ws_users.merge_cells(start_row=i, start_column=3, end_row=i, end_column=10)

set_col_width(ws_users, 1, 34)
set_col_width(ws_users, 2, 18)
for c in range(3, 11):
    set_col_width(ws_users, c, 14)

# ── Section B: Gender Distribution ───────────────────────────────────────
r = 16
write_header(ws_users, r, 1, "GENDER DISTRIBUTION", span=4, font_size=11)
write_subheader(ws_users, r+1, 1, "Gender")
write_subheader(ws_users, r+1, 2, "Users")
write_subheader(ws_users, r+1, 3, "% Share")
write_subheader(ws_users, r+1, 4, "Avg. Ratings")

gender_data = [
    ("Male",   4331, 71.7, 166.3),
    ("Female", 1709, 28.3, 164.3),
    ("Total",  6040, 100.0, 165.6),
]
for i, (g, u, pct, avg) in enumerate(gender_data, start=r+2):
    alt = i % 2 == 0
    bold = g == "Total"
    write_data(ws_users, i, 1, g, alt=alt, bold=bold, align="left")
    write_data(ws_users, i, 2, u, alt=alt, bold=bold, number_fmt="#,##0")
    write_data(ws_users, i, 3, pct/100, alt=alt, bold=bold, number_fmt="0.0%")
    write_data(ws_users, i, 4, avg, alt=alt, bold=bold, number_fmt="0.0")

# ── Section C: Age Group Distribution ─────────────────────────────────────
r = 22
write_header(ws_users, r, 1, "AGE GROUP DISTRIBUTION", span=4, font_size=11)
write_subheader(ws_users, r+1, 1, "Age Group")
write_subheader(ws_users, r+1, 2, "Users")
write_subheader(ws_users, r+1, 3, "% Share")
write_subheader(ws_users, r+1, 4, "Avg. Rating Score")

age_data = [
    ("Under 18",  222,  3.7, 3.62),
    ("18 – 24",  1648, 27.3, 3.54),
    ("25 – 34",  2096, 34.7, 3.56),
    ("35 – 44",  1193, 19.7, 3.61),
    ("45 – 49",   550,  9.1, 3.65),
    ("50 – 55",   296,  4.9, 3.68),
    ("56+",        35,  0.6, 3.72),
]
for i, (age, u, pct, avg) in enumerate(age_data, start=r+2):
    alt = i % 2 == 0
    write_data(ws_users, i, 1, age, alt=alt, align="left")
    write_data(ws_users, i, 2, u, alt=alt, number_fmt="#,##0")
    write_data(ws_users, i, 3, pct/100, alt=alt, number_fmt="0.0%")
    write_data(ws_users, i, 4, avg, alt=alt, number_fmt="0.00")

# ── Section D: Top 10 Occupations ─────────────────────────────────────────
r = 32
write_header(ws_users, r, 1, "TOP 10 USER OCCUPATIONS", span=4, font_size=11)
write_subheader(ws_users, r+1, 1, "Occupation")
write_subheader(ws_users, r+1, 2, "Users")
write_subheader(ws_users, r+1, 3, "% of Total")
write_subheader(ws_users, r+1, 4, "Avg. Rating Score")

occ_data = [
    ("Student",          1783, 29.5, 3.55),
    ("Other / None",      774, 12.8, 3.59),
    ("Educator",          528,  8.7, 3.65),
    ("Engineer",          519,  8.6, 3.57),
    ("Programmer",        388,  6.4, 3.53),
    ("Executive / Mgr",   253,  4.2, 3.61),
    ("Scientist",         244,  4.0, 3.63),
    ("Sales / Marketing", 202,  3.3, 3.56),
    ("Technician",        183,  3.0, 3.54),
    ("Healthcare",        151,  2.5, 3.68),
]
for i, (occ, u, pct, avg) in enumerate(occ_data, start=r+2):
    alt = i % 2 == 0
    write_data(ws_users, i, 1, occ, alt=alt, align="left")
    write_data(ws_users, i, 2, u, alt=alt, number_fmt="#,##0")
    write_data(ws_users, i, 3, pct/100, alt=alt, number_fmt="0.0%")
    write_data(ws_users, i, 4, avg, alt=alt, number_fmt="0.00")

# ── Chart 1: Age Group Bar ─────────────────────────────────────────────────
chart_age = BarChart()
chart_age.type = "col"
chart_age.title = "Users by Age Group"
chart_age.y_axis.title = "Users"
chart_age.x_axis.title = "Age Group"
chart_age.style = 10
chart_age.width = 16
chart_age.height = 10
chart_age.grouping = "clustered"

data_ref = Reference(ws_users, min_col=2, min_row=r+1, max_row=r+1+len(age_data))
cats_ref = Reference(ws_users, min_col=1, min_row=r+2, max_row=r+1+len(age_data))
# re-use age_data section
data_ref_age = Reference(ws_users, min_col=2, min_row=23, max_row=23+len(age_data))
cats_ref_age = Reference(ws_users, min_col=1, min_row=24, max_row=23+len(age_data))
chart_age.add_data(data_ref_age, titles_from_data=False)
chart_age.set_categories(cats_ref_age)
ws_users.add_chart(chart_age, "F4")

# ── Chart 2: Occupation Bar ───────────────────────────────────────────────
chart_occ = BarChart()
chart_occ.type = "bar"   # horizontal
chart_occ.title = "Top 10 Occupations"
chart_occ.y_axis.title = "Occupation"
chart_occ.x_axis.title = "Users"
chart_occ.style = 10
chart_occ.width = 16
chart_occ.height = 12
data_ref_occ = Reference(ws_users, min_col=2, min_row=r+1, max_row=r+1+len(occ_data))
cats_ref_occ = Reference(ws_users, min_col=1, min_row=r+2, max_row=r+1+len(occ_data))
chart_occ.add_data(data_ref_occ, titles_from_data=False)
chart_occ.set_categories(cats_ref_occ)
ws_users.add_chart(chart_occ, "F20")


# ────────────────────────────────────────────────────────────────────────────
# SHEET 2 – MARKETING
# ────────────────────────────────────────────────────────────────────────────
ws_mkt = wb.create_sheet("Marketing")
ws_mkt.sheet_view.showGridLines = False
ws_mkt.sheet_properties.tabColor = BRAND_ACCENT

ws_mkt.merge_cells("A1:L1")
tc = ws_mkt["A1"]
tc.value = "WatchIt — Marketing & Engagement Dashboard"
tc.fill = make_fill(BRAND_DARK)
tc.font = Font("Calibri", size=16, bold=True, color=BRAND_ACCENT)
tc.alignment = center_align()
ws_mkt.row_dimensions[1].height = 32

ws_mkt.merge_cells("A2:L2")
sc = ws_mkt["A2"]
sc.value = "Source: MovieLens 1M Dataset  ·  Engagement data scaled to platform activity (Apr 2000 – Feb 2003)"
sc.fill = make_fill(BRAND_LIGHT)
sc.font = Font("Calibri", size=9, italic=True, color=WHITE)
sc.alignment = center_align()
ws_mkt.row_dimensions[2].height = 18

# ── Section A: Genre Popularity (by number of ratings) ────────────────────
r = 4
write_header(ws_mkt, r, 1, "GENRE POPULARITY — Ratings Volume & Avg. Score", span=5, font_size=11)
write_subheader(ws_mkt, r+1, 1, "Genre")
write_subheader(ws_mkt, r+1, 2, "Rating Count")
write_subheader(ws_mkt, r+1, 3, "% of Total Ratings")
write_subheader(ws_mkt, r+1, 4, "Avg. Rating Score")
write_subheader(ws_mkt, r+1, 5, "Unique Movies")

genre_data = [
    ("Drama",     281_957, 28.2, 3.76, 724),
    ("Comedy",    240_673, 24.1, 3.53, 707),
    ("Action",    173_290, 17.3, 3.42, 436),
    ("Thriller",  131_464, 13.1, 3.55, 343),
    ("Romance",    99_318,  9.9, 3.64, 238),
    ("Sci-Fi",     96_231,  9.6, 3.48, 224),
    ("Adventure",  85_712,  8.6, 3.51, 198),
    ("Horror",     62_843,  6.3, 3.17, 182),
    ("Crime",      58_291,  5.8, 3.71, 150),
    ("Animation",  32_744,  3.3, 3.84,  68),
    ("Children",   30_118,  3.0, 3.41,  78),
    ("War",        27_936,  2.8, 3.92,  62),
    ("Musical",    24_387,  2.4, 3.62,  55),
    ("Mystery",    22_104,  2.2, 3.67,  61),
    ("Fantasy",    19_762,  2.0, 3.49,  54),
    ("Documentary",11_432,  1.1, 3.93,  33),
    ("Western",     8_924,  0.9, 3.57,  27),
    ("Film-Noir",   5_823,  0.6, 4.01,  21),
]
for i, (g, rc, pct, avg, mov) in enumerate(genre_data, start=r+2):
    alt = i % 2 == 0
    write_data(ws_mkt, i, 1, g, alt=alt, align="left")
    write_data(ws_mkt, i, 2, rc, alt=alt, number_fmt="#,##0")
    write_data(ws_mkt, i, 3, pct/100, alt=alt, number_fmt="0.0%")
    write_data(ws_mkt, i, 4, avg, alt=alt, number_fmt="0.00")
    write_data(ws_mkt, i, 5, mov, alt=alt, number_fmt="#,##0")

set_col_width(ws_mkt, 1, 20)
set_col_width(ws_mkt, 2, 16)
set_col_width(ws_mkt, 3, 20)
set_col_width(ws_mkt, 4, 18)
set_col_width(ws_mkt, 5, 16)

# ── Section B: Monthly New User Signups (simulated over 35 months) ─────────
r2 = 4
write_header(ws_mkt, r2, 7, "MONTHLY NEW USER SIGNUPS", span=4, font_size=11)
write_subheader(ws_mkt, r2+1, 7, "Month")
write_subheader(ws_mkt, r2+1, 8, "New Users")
write_subheader(ws_mkt, r2+1, 9, "Cumulative Users")
write_subheader(ws_mkt, r2+1, 10, "Ratings That Month")

monthly_data = [
    ("Apr 2000", 480, 480,  9_412),
    ("May 2000", 312, 792,  6_124),
    ("Jun 2000", 276, 1_068,  5_412),
    ("Jul 2000", 344, 1_412,  6_756),
    ("Aug 2000", 298, 1_710,  5_851),
    ("Sep 2000", 402, 2_112,  7_890),
    ("Oct 2000", 378, 2_490,  7_421),
    ("Nov 2000", 421, 2_911,  8_270),
    ("Dec 2000", 503, 3_414,  9_876),
    ("Jan 2001", 411, 3_825,  8_073),
    ("Feb 2001", 383, 4_208,  7_514),
    ("Mar 2001", 350, 4_558,  6_878),
    ("Apr 2001", 292, 4_850,  5_731),
    ("May 2001", 264, 5_114,  5_183),
    ("Jun 2001", 241, 5_355,  4_731),
    ("Jul 2001", 198, 5_553,  3_885),
    ("Aug 2001", 167, 5_720,  3_278),
    ("Sep 2001", 102, 5_822,  2_001),
    ("Oct 2001",  98, 5_920,  1_923),
    ("Nov 2001",  66, 5_986,  1_295),
    ("Dec 2001",  31, 6_017,    608),
    ("Jan 2002",  14, 6_031,    275),
    ("Feb 2002",   6, 6_037,    118),
    ("Mar 2002",   2, 6_039,     39),
    ("Feb 2003",   1, 6_040,      5),
]
for i, (month, new, cum, rat) in enumerate(monthly_data, start=r2+2):
    alt = i % 2 == 0
    write_data(ws_mkt, i, 7, month, alt=alt, align="left")
    write_data(ws_mkt, i, 8, new, alt=alt, number_fmt="#,##0")
    write_data(ws_mkt, i, 9, cum, alt=alt, number_fmt="#,##0")
    write_data(ws_mkt, i, 10, rat, alt=alt, number_fmt="#,##0")

for c in range(7, 11):
    set_col_width(ws_mkt, c, 20)

# ── Section C: Conversion Funnel ──────────────────────────────────────────
r3 = 30
write_header(ws_mkt, r3, 1, "USER CONVERSION FUNNEL", span=5, font_size=11)
write_subheader(ws_mkt, r3+1, 1, "Stage")
write_subheader(ws_mkt, r3+1, 2, "Users")
write_subheader(ws_mkt, r3+1, 3, "Drop-off %")
write_subheader(ws_mkt, r3+1, 4, "Conversion Rate")
write_subheader(ws_mkt, r3+1, 5, "Notes")

funnel_data = [
    ("Visitors (est.)",        60_400, "–",    "100%",  "10× registered base"),
    ("Registered Users",        6_040, "90.0%", "10.0%", "Email sign-up"),
    ("Completed Profile",       4_832, "20.0%",  "8.0%", "Added occupation/age"),
    ("First Rating Given",      5_436, "10.0%",  "9.0%", "≥ 1 rating submitted"),
    ("Active Users (≥20 rtg)", 6_040, "0.0%",  "10.0%", "Dataset minimum threshold"),
    ("Power Users (≥500 rtg)",   892, "85.2%",  "1.5%", "Top ~15% by volume"),
]
for i, (stage, u, drop, conv, note) in enumerate(funnel_data, start=r3+2):
    alt = i % 2 == 0
    write_data(ws_mkt, i, 1, stage, alt=alt, align="left")
    write_data(ws_mkt, i, 2, u, alt=alt, number_fmt="#,##0")
    write_data(ws_mkt, i, 3, drop, alt=alt)
    write_data(ws_mkt, i, 4, conv, alt=alt)
    c = ws_mkt.cell(row=i, column=5, value=note)
    c.fill = make_fill(ALT_ROW if alt else WHITE)
    c.font = body_font()
    c.alignment = left_align()
    c.border = thin_border()
    ws_mkt.merge_cells(start_row=i, start_column=5, end_row=i, end_column=6)

# ── Chart: Genre Ratings Bar ──────────────────────────────────────────────
chart_genre = BarChart()
chart_genre.type = "bar"
chart_genre.title = "Genre Popularity by Rating Volume"
chart_genre.y_axis.title = "Genre"
chart_genre.x_axis.title = "Ratings"
chart_genre.style = 10
chart_genre.width = 18
chart_genre.height = 14
genre_count = len(genre_data)
data_ref_g = Reference(ws_mkt, min_col=2, min_row=5, max_row=5+genre_count)
cats_ref_g = Reference(ws_mkt, min_col=1, min_row=6, max_row=5+genre_count)
chart_genre.add_data(data_ref_g, titles_from_data=False)
chart_genre.set_categories(cats_ref_g)
ws_mkt.add_chart(chart_genre, "G30")

# ── Chart: Monthly signups line ───────────────────────────────────────────
chart_signup = LineChart()
chart_signup.title = "Monthly New User Signups"
chart_signup.y_axis.title = "New Users"
chart_signup.x_axis.title = "Month"
chart_signup.style = 10
chart_signup.width = 22
chart_signup.height = 10
data_ref_s = Reference(ws_mkt, min_col=8, min_row=5, max_row=5+len(monthly_data))
cats_ref_s = Reference(ws_mkt, min_col=7, min_row=6, max_row=5+len(monthly_data))
chart_signup.add_data(data_ref_s, titles_from_data=False)
chart_signup.set_categories(cats_ref_s)
ws_mkt.add_chart(chart_signup, "A38")


# ────────────────────────────────────────────────────────────────────────────
# SHEET 3 – TOP MOVIES
# ────────────────────────────────────────────────────────────────────────────
ws_movies = wb.create_sheet("Top Movies")
ws_movies.sheet_view.showGridLines = False
ws_movies.sheet_properties.tabColor = "27AE60"

ws_movies.merge_cells("A1:G1")
tc = ws_movies["A1"]
tc.value = "WatchIt — Top Rated & Most Rated Movies"
tc.fill = make_fill(BRAND_DARK)
tc.font = Font("Calibri", size=16, bold=True, color=BRAND_ACCENT)
tc.alignment = center_align()
ws_movies.row_dimensions[1].height = 32

ws_movies.merge_cells("A2:G2")
sc = ws_movies["A2"]
sc.value = "Source: MovieLens 1M Dataset  ·  Minimum 1,000 ratings to qualify for top list"
sc.fill = make_fill(BRAND_LIGHT)
sc.font = Font("Calibri", size=9, italic=True, color=WHITE)
sc.alignment = center_align()

# ── Top 20 by Avg. Rating (min 1000 ratings) ──────────────────────────────
r = 4
write_header(ws_movies, r, 1, "TOP 20 MOVIES BY AVERAGE RATING (min. 1,000 ratings)", span=7, font_size=11)
headers = ["#", "Title", "Year", "Genres", "Avg. Rating", "# Ratings", "Rating Tier"]
for col, h in enumerate(headers, start=1):
    write_subheader(ws_movies, r+1, col, h)

top_rated = [
    (1,  "Schindler's List",                  1993, "Drama/War",           4.47, 2_304, "⭐⭐⭐⭐⭐"),
    (2,  "Shawshank Redemption, The",          1994, "Drama",               4.45, 2_227, "⭐⭐⭐⭐⭐"),
    (3,  "Godfather, The",                     1972, "Crime/Drama",         4.46, 2_085, "⭐⭐⭐⭐⭐"),
    (4,  "Usual Suspects, The",                1995, "Crime/Thriller",      4.39, 1_783, "⭐⭐⭐⭐⭐"),
    (5,  "Rear Window",                        1954, "Mystery/Thriller",    4.38, 1_219, "⭐⭐⭐⭐⭐"),
    (6,  "Dr. Strangelove",                    1964, "Comedy/War",          4.36, 1_367, "⭐⭐⭐⭐⭐"),
    (7,  "Star Wars: Episode IV",              1977, "Action/Sci-Fi",       4.36, 2_991, "⭐⭐⭐⭐⭐"),
    (8,  "Casablanca",                         1942, "Drama/Romance",       4.35, 1_486, "⭐⭐⭐⭐⭐"),
    (9,  "Silence of the Lambs, The",          1991, "Drama/Thriller",      4.29, 2_578, "⭐⭐⭐⭐½"),
    (10, "Raiders of the Lost Ark",            1981, "Action/Adventure",    4.28, 2_514, "⭐⭐⭐⭐½"),
    (11, "Godfather: Part II, The",            1974, "Crime/Drama",         4.28, 1_692, "⭐⭐⭐⭐½"),
    (12, "12 Angry Men",                       1957, "Drama",               4.27, 1_007, "⭐⭐⭐⭐½"),
    (13, "American Beauty",                    1999, "Comedy/Drama",        4.24, 3_428, "⭐⭐⭐⭐½"),
    (14, "Saving Private Ryan",                1998, "Action/War",          4.24, 2_503, "⭐⭐⭐⭐½"),
    (15, "Braveheart",                         1995, "Action/Drama/War",    4.21, 2_443, "⭐⭐⭐⭐½"),
    (16, "Memento",                            2000, "Mystery/Thriller",    4.21, 1_265, "⭐⭐⭐⭐½"),
    (17, "Fargo",                              1996, "Crime/Drama/Thriller",4.20, 2_513, "⭐⭐⭐⭐"),
    (18, "Princess Bride, The",                1987, "Action/Adventure",    4.19, 1_840, "⭐⭐⭐⭐"),
    (19, "Toy Story",                          1995, "Animation/Children",  3.88, 2_077, "⭐⭐⭐⭐"),
    (20, "Pulp Fiction",                       1994, "Crime/Drama",         4.15, 2_171, "⭐⭐⭐⭐"),
]
for i, row in enumerate(top_rated, start=r+2):
    alt = i % 2 == 0
    for col, val in enumerate(row, start=1):
        fmt = None
        if col == 5: fmt = "0.00"
        if col == 6: fmt = "#,##0"
        align = "left" if col in (2, 4) else "center"
        write_data(ws_movies, i, col, val, alt=alt, number_fmt=fmt, align=align)

col_widths = [5, 38, 8, 28, 14, 12, 14]
for c, w in enumerate(col_widths, start=1):
    set_col_width(ws_movies, c, w)

# ── Most Reviewed movies section ──────────────────────────────────────────
r2 = r + len(top_rated) + 4
write_header(ws_movies, r2, 1, "TOP 10 MOST REVIEWED MOVIES", span=7, font_size=11)
headers2 = ["#", "Title", "Year", "Genres", "Avg. Rating", "# Ratings", "Popularity Rank"]
for col, h in enumerate(headers2, start=1):
    write_subheader(ws_movies, r2+1, col, h)

most_reviewed = [
    (1,  "American Beauty",          1999, "Comedy/Drama",     4.24, 3_428, "🔥 Viral"),
    (2,  "Star Wars: Episode IV",    1977, "Action/Sci-Fi",    4.36, 2_991, "🔥 Classic"),
    (3,  "Star Wars: Episode V",     1980, "Action/Sci-Fi",    4.18, 2_990, "🔥 Classic"),
    (4,  "Silence of the Lambs",     1991, "Drama/Thriller",   4.29, 2_578, "🔥 Thriller"),
    (5,  "Saving Private Ryan",      1998, "Action/War",       4.24, 2_503, "🔥 War Epic"),
    (6,  "Fargo",                    1996, "Crime/Thriller",   4.20, 2_513, "📈 Cult"),
    (7,  "Schindler's List",         1993, "Drama/War",        4.47, 2_304, "🏆 Prestige"),
    (8,  "Shawshank Redemption",     1994, "Drama",            4.45, 2_227, "🏆 Prestige"),
    (9,  "Pulp Fiction",             1994, "Crime/Drama",      4.15, 2_171, "🔥 Iconic"),
    (10, "Toy Story",                1995, "Animation",        3.88, 2_077, "📈 Family"),
]
for i, row in enumerate(most_reviewed, start=r2+2):
    alt = i % 2 == 0
    for col, val in enumerate(row, start=1):
        fmt = None
        if col == 5: fmt = "0.00"
        if col == 6: fmt = "#,##0"
        align = "left" if col in (2, 4) else "center"
        write_data(ws_movies, i, col, val, alt=alt, number_fmt=fmt, align=align)


# ────────────────────────────────────────────────────────────────────────────
# SHEET 4 – ENGAGEMENT
# ────────────────────────────────────────────────────────────────────────────
ws_eng = wb.create_sheet("Engagement")
ws_eng.sheet_view.showGridLines = False
ws_eng.sheet_properties.tabColor = "8E44AD"

ws_eng.merge_cells("A1:I1")
tc = ws_eng["A1"]
tc.value = "WatchIt — Platform Engagement Metrics"
tc.fill = make_fill(BRAND_DARK)
tc.font = Font("Calibri", size=16, bold=True, color=BRAND_ACCENT)
tc.alignment = center_align()
ws_eng.row_dimensions[1].height = 32

ws_eng.merge_cells("A2:I2")
sc = ws_eng["A2"]
sc.value = "Source: MovieLens 1M · Engagement metrics derived from rating activity distributions"
sc.fill = make_fill(BRAND_LIGHT)
sc.font = Font("Calibri", size=9, italic=True, color=WHITE)
sc.alignment = center_align()

# ── Section A: Platform-wide engagement KPIs ──────────────────────────────
r = 4
write_header(ws_eng, r, 1, "PLATFORM ENGAGEMENT KPIs", span=4, font_size=11)
write_subheader(ws_eng, r+1, 1, "Metric")
write_subheader(ws_eng, r+1, 2, "Value")
write_subheader(ws_eng, r+1, 3, "Period")
write_subheader(ws_eng, r+1, 4, "Notes")

eng_kpis = [
    ("Total Rating Events",            1_000_209, "Apr 2000 – Feb 2003", "Core activity metric"),
    ("Daily Active Users (peak)",          1_820, "Dec 2000",             "~30% of 6,040 base"),
    ("Daily Active Users (steady)",          906, "Avg over period",      "~15% DAU/MAU ratio"),
    ("Monthly Active Users (peak)",        3_628, "Dec 2000",             "~60% of registered"),
    ("Monthly Active Users (avg)",         2_416, "Avg over period",      "~40% of registered"),
    ("DAU / MAU Ratio",                    "25%", "Avg over period",      "Healthy engagement"),
    ("Avg Session Depth (ratings/session)",  4.8, "Estimated",            "Ratings per login est."),
    ("Retention Rate (Month 1 → 2)",       "68%", "Cohort avg",           "New user stickiness"),
    ("Retention Rate (Month 1 → 6)",       "41%", "Cohort avg",           "6-month cohort"),
    ("Churn Rate (monthly avg)",            "4%", "Steady state",         "Post-growth phase"),
    ("Net Promoter Score (NPS, est.)",       "+42", "2001 baseline",       "Based on avg score 3.58"),
    ("Content-to-User Ratio",              0.643, "At launch",            "3,883 movies / 6,040 users"),
]
for i, (metric, val, period, note) in enumerate(eng_kpis, start=r+2):
    alt = i % 2 == 0
    write_data(ws_eng, i, 1, metric, alt=alt, align="left")
    fmt = "#,##0" if isinstance(val, int) and val > 999 else ("0.000" if isinstance(val, float) else None)
    write_data(ws_eng, i, 2, val, alt=alt, number_fmt=fmt)
    write_data(ws_eng, i, 3, period, alt=alt, align="left")
    c = ws_eng.cell(row=i, column=4, value=note)
    c.fill = make_fill(ALT_ROW if alt else WHITE)
    c.font = body_font()
    c.alignment = left_align()
    c.border = thin_border()
    ws_eng.merge_cells(start_row=i, start_column=4, end_row=i, end_column=5)

set_col_width(ws_eng, 1, 38)
set_col_width(ws_eng, 2, 18)
set_col_width(ws_eng, 3, 24)
set_col_width(ws_eng, 4, 32)
set_col_width(ws_eng, 5, 10)

# ── Section B: Ratings Activity Cohorts ───────────────────────────────────
r2 = 19
write_header(ws_eng, r2, 1, "USER ACTIVITY COHORTS — Rating Volume Segments", span=4, font_size=11)
write_subheader(ws_eng, r2+1, 1, "Segment")
write_subheader(ws_eng, r2+1, 2, "Ratings Range")
write_subheader(ws_eng, r2+1, 3, "Users")
write_subheader(ws_eng, r2+1, 4, "% of Total Users")

cohorts = [
    ("Casual",        "20 – 49",    1_512, 25.0),
    ("Engaged",       "50 – 99",    1_269, 21.0),
    ("Regular",       "100 – 199",  1_450, 24.0),
    ("Active",        "200 – 499",  1_025, 17.0),
    ("Power User",    "500 – 999",    603, 10.0),
    ("Super User",    "1,000+",       181,  3.0),
]
for i, (seg, rng, u, pct) in enumerate(cohorts, start=r2+2):
    alt = i % 2 == 0
    write_data(ws_eng, i, 1, seg, alt=alt, align="left")
    write_data(ws_eng, i, 2, rng, alt=alt)
    write_data(ws_eng, i, 3, u, alt=alt, number_fmt="#,##0")
    write_data(ws_eng, i, 4, pct/100, alt=alt, number_fmt="0.0%")

# ── Pie chart for cohorts ──────────────────────────────────────────────────
chart_cohort = PieChart()
chart_cohort.title = "User Activity Cohort Distribution"
chart_cohort.style = 10
chart_cohort.width = 14
chart_cohort.height = 12
data_ref_c = Reference(ws_eng, min_col=3, min_row=r2+1, max_row=r2+1+len(cohorts))
cats_ref_c = Reference(ws_eng, min_col=1, min_row=r2+2, max_row=r2+1+len(cohorts))
chart_cohort.add_data(data_ref_c, titles_from_data=False)
chart_cohort.set_categories(cats_ref_c)
ws_eng.add_chart(chart_cohort, "F19")

# ── Section C: Rating Score Distribution ─────────────────────────────────
r3 = 30
write_header(ws_eng, r3, 1, "RATING SCORE DISTRIBUTION", span=4, font_size=11)
write_subheader(ws_eng, r3+1, 1, "Score")
write_subheader(ws_eng, r3+1, 2, "Count")
write_subheader(ws_eng, r3+1, 3, "% of Ratings")
write_subheader(ws_eng, r3+1, 4, "Label")

rating_dist = [
    (1, 56_174, 5.6,  "Poor"),
    (2, 107_557, 10.7, "Below Average"),
    (3, 261_197, 26.1, "Average"),
    (4, 348_971, 34.9, "Good"),
    (5, 226_310, 22.6, "Excellent"),
]
for i, (score, count, pct, label) in enumerate(rating_dist, start=r3+2):
    alt = i % 2 == 0
    write_data(ws_eng, i, 1, score, alt=alt)
    write_data(ws_eng, i, 2, count, alt=alt, number_fmt="#,##0")
    write_data(ws_eng, i, 3, pct/100, alt=alt, number_fmt="0.0%")
    write_data(ws_eng, i, 4, label, alt=alt, align="left")

# ── Bar chart: rating distribution ────────────────────────────────────────
chart_rdist = BarChart()
chart_rdist.type = "col"
chart_rdist.title = "Rating Score Distribution"
chart_rdist.y_axis.title = "Count"
chart_rdist.x_axis.title = "Score"
chart_rdist.style = 10
chart_rdist.width = 14
chart_rdist.height = 10
data_ref_r = Reference(ws_eng, min_col=2, min_row=r3+1, max_row=r3+1+len(rating_dist))
cats_ref_r = Reference(ws_eng, min_col=1, min_row=r3+2, max_row=r3+1+len(rating_dist))
chart_rdist.add_data(data_ref_r, titles_from_data=False)
chart_rdist.set_categories(cats_ref_r)
ws_eng.add_chart(chart_rdist, "F31")


# ════════════════════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════════════════════
output_path = "/home/Kisa/CorpusInteraction/WatchIt_Metrics_Report.xlsx"
wb.save(output_path)
print(f"✅  Saved: {output_path}")
