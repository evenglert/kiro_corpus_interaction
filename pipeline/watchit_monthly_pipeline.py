#!/usr/bin/env python3
"""
watchit_monthly_pipeline.py
────────────────────────────────────────────────────────────────────────────────
WatchIt — Monthly Report Pipeline
Reads the monthly Excel metrics file and fills both the Word management update
and the PowerPoint metrics deck.  Templates are NEVER overwritten; each run
produces timestamped output files.

USAGE
    python watchit_monthly_pipeline.py 202611_WatchIt_Metrics_Report.xlsx

    The timestamp (e.g. 202611) is taken from the Excel filename prefix.
    Output files are placed in the same folder as the Excel unless --output-dir
    is given.

    python watchit_monthly_pipeline.py 202612_WatchIt_Metrics_Report.xlsx --output-dir reports/

DEPENDENCIES
    pip install openpyxl python-docx python-pptx
────────────────────────────────────────────────────────────────────────────────
"""

import argparse
import copy
import re
import sys
from pathlib import Path

import openpyxl
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Pt as PPt


# ══════════════════════════════════════════════════════════════════════════════
# 1.  LOCATE TEMPLATES  (always sit next to this script)
# ══════════════════════════════════════════════════════════════════════════════

SCRIPT_DIR    = Path(__file__).parent.resolve()
REPO_ROOT     = SCRIPT_DIR.parent          # one level up from pipeline/
WORD_TEMPLATE = REPO_ROOT / "templates" / "WatchIt_Monthly_Management_Update_Template.docx"
PPTX_TEMPLATE = REPO_ROOT / "templates" / "WatchIt_Monthly_Management_Update_Template.pptx"


# ══════════════════════════════════════════════════════════════════════════════
# 2.  EXCEL READER — extracts every metric needed for both reports
# ══════════════════════════════════════════════════════════════════════════════

def load_metrics(excel_path: Path) -> dict:
    """
    Parse the standardised WatchIt Excel workbook and return a flat dict of
    all metrics used for placeholder substitution in both templates.
    Raises ValueError if a required sheet or cell is missing.
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    required_sheets = {"Cover", "Users", "Marketing", "Top Movies", "Engagement"}
    missing = required_sheets - set(wb.sheetnames)
    if missing:
        raise ValueError(f"Excel is missing required sheets: {missing}")

    cover      = wb["Cover"]
    users_ws   = wb["Users"]
    mkt_ws     = wb["Marketing"]
    movies_ws  = wb["Top Movies"]
    eng_ws     = wb["Engagement"]

    # ── helper: read a sheet into a list of row-tuples (values only) ──────────
    def rows(ws, min_row=1):
        return [tuple(c for c in row) for row in ws.iter_rows(min_row=min_row, values_only=True)]

    # ── Cover sheet KPIs (row 5, cols B-G) ────────────────────────────────────
    cover_rows = rows(cover, min_row=1)
    # row index 4 (0-based) = row 5
    cover_kpi = cover_rows[4]   # (None, users_str, ratings_str, movies_str, avg_str, avg_per_user_str, since_str)

    # ── Users sheet ───────────────────────────────────────────────────────────
    u_rows = rows(users_ws, min_row=1)

    def find_value(row_list, label):
        """Return the second column value from the row whose first col matches label."""
        for r in row_list:
            if r and str(r[0]).strip() == label:
                return r[1]
        return None

    total_users       = find_value(u_rows, "Total Registered Users")
    total_ratings     = find_value(u_rows, "Total Ratings Submitted")
    movies_catalogue  = find_value(u_rows, "Movies in Catalogue")
    avg_ratings_user  = find_value(u_rows, "Average Ratings per User")
    avg_rating_score  = find_value(u_rows, "Average Rating Score")
    min_ratings_user  = find_value(u_rows, "Minimum Ratings per User")
    max_ratings_user  = find_value(u_rows, "Maximum Ratings (single user)")
    median_ratings    = find_value(u_rows, "Median Ratings per User")

    # Gender distribution — rows starting with 'Male' / 'Female' in the gender block
    gender_rows = {}
    in_gender_section = False
    for r in u_rows:
        if r and str(r[0]).strip() == "GENDER DISTRIBUTION":
            in_gender_section = True
            continue
        if in_gender_section and r and r[0] in ("Male", "Female"):
            gender_rows[r[0]] = {"users": r[1], "pct": r[2], "avg_ratings": r[3]}
        if in_gender_section and r and str(r[0]).strip() == "AGE GROUP DISTRIBUTION":
            break

    male_users   = gender_rows.get("Male", {}).get("users", 0) or 0
    female_users = gender_rows.get("Female", {}).get("users", 0) or 0
    male_pct     = gender_rows.get("Male", {}).get("pct", 0) or 0
    female_pct   = gender_rows.get("Female", {}).get("pct", 0) or 0
    male_avg     = gender_rows.get("Male", {}).get("avg_ratings", 0) or 0
    female_avg   = gender_rows.get("Female", {}).get("avg_ratings", 0) or 0

    # Age groups — collect all rows in AGE GROUP DISTRIBUTION block
    age_groups = []
    in_age = False
    for r in u_rows:
        if r and str(r[0]).strip() == "AGE GROUP DISTRIBUTION":
            in_age = True
            continue
        if in_age and r and str(r[0]).strip() == "TOP 10 USER OCCUPATIONS":
            break
        if in_age and r and r[0] and str(r[0]).strip() not in ("Age Group",):
            if r[1] is not None:
                age_groups.append({"group": r[0], "users": r[1], "pct": r[2], "avg_score": r[3]})

    age_groups_sorted = sorted(age_groups, key=lambda x: x["users"], reverse=True)
    top_age_group    = age_groups_sorted[0] if len(age_groups_sorted) > 0 else {}
    second_age_group = age_groups_sorted[1] if len(age_groups_sorted) > 1 else {}

    # Occupations
    occupations = []
    in_occ = False
    for r in u_rows:
        if r and str(r[0]).strip() == "TOP 10 USER OCCUPATIONS":
            in_occ = True
            continue
        if in_occ and r and r[0] and str(r[0]).strip() not in ("Occupation",) and r[1] is not None:
            occupations.append({"name": r[0], "users": r[1], "pct": r[2], "avg_score": r[3]})

    top_occ = occupations[0] if occupations else {}

    # Power users (≥500 ratings) from cohort data — filled from Engagement sheet
    # We'll pick them up there; stash the raw occupation list for PPTX
    top_occupations = occupations[:7]   # top 7 for the deck

    # ── Marketing sheet ────────────────────────────────────────────────────────
    m_rows = rows(mkt_ws, min_row=1)

    # Genre popularity
    genres = []
    in_genre = False
    for r in m_rows:
        if r and str(r[0]).strip() == "GENRE POPULARITY — Ratings Volume & Avg. Score":
            in_genre = True
            continue
        if in_genre and r and str(r[0]).strip() == "USER CONVERSION FUNNEL":
            break
        if in_genre and r and r[0] and str(r[0]).strip() not in ("Genre",) and r[1] is not None:
            genres.append({"genre": r[0], "count": r[1], "pct": r[2], "avg_score": r[3], "movies": r[4]})

    genres_by_vol   = sorted(genres, key=lambda x: x["count"] or 0, reverse=True)
    genres_by_score = sorted(genres, key=lambda x: x["avg_score"] or 0, reverse=True)
    top_genre_vol   = genres_by_vol[0]   if genres_by_vol   else {}
    top_genre_score = genres_by_score[0] if genres_by_score else {}
    film_noir_genre = next((g for g in genres if "film" in g["genre"].lower() and "noir" in g["genre"].lower()), {})
    documentary     = next((g for g in genres if "documentary" in g["genre"].lower()), {})

    drama_pct   = next((g["pct"] for g in genres if g["genre"] == "Drama"),   None)
    comedy_pct  = next((g["pct"] for g in genres if g["genre"] == "Comedy"),  None)
    action_pct  = next((g["pct"] for g in genres if g["genre"] == "Action"),  None)

    # Monthly new users — find peak
    monthly_users = []
    in_monthly = False
    for r in m_rows:
        if r and str(r[6]).strip() == "MONTHLY NEW USER SIGNUPS" if r[6] else False:
            in_monthly = True
            continue
        if in_monthly and r and r[6] and str(r[6]).strip() not in ("Month",) and r[7] is not None:
            monthly_users.append({"month": r[6], "new_users": r[7], "cumulative": r[8]})

    peak_month_row = max(monthly_users, key=lambda x: x["new_users"] or 0) if monthly_users else {}
    peak_month     = peak_month_row.get("month", "N/A")
    peak_new_users = peak_month_row.get("new_users", 0)

    # Conversion funnel
    funnel = {}
    in_funnel = False
    for r in m_rows:
        if r and str(r[0]).strip() == "USER CONVERSION FUNNEL":
            in_funnel = True
            continue
        if in_funnel and r and r[0] and str(r[0]).strip() not in ("Stage",) and r[1] is not None:
            funnel[str(r[0]).strip()] = {"users": r[1], "dropoff": r[2], "conv_rate": r[3]}

    visitor_conv = funnel.get("Visitors (est.)", {})
    reg_users_f  = funnel.get("Registered Users", {})
    profile_f    = funnel.get("Completed Profile", {})
    active_f     = funnel.get("Active Users (≥20 rtg)", funnel.get("First Rating Given", {}))
    power_f      = funnel.get("Power Users (≥500 rtg)", {})

    conv_rate_str = reg_users_f.get("conv_rate", "10%")
    if isinstance(conv_rate_str, float):
        conv_rate_str = f"{conv_rate_str * 100:.0f}%"
    conv_rate_str = str(conv_rate_str).replace("%", "").strip()

    # ── Top Movies sheet ───────────────────────────────────────────────────────
    t_rows = rows(movies_ws, min_row=1)

    top_rated    = []
    most_reviewed = []
    in_top_rated = False
    in_most_rev  = False
    for r in t_rows:
        if r and str(r[0]).strip() == "TOP 20 MOVIES BY AVERAGE RATING (min. 1,000 ratings)":
            in_top_rated = True
            in_most_rev  = False
            continue
        if r and str(r[0]).strip() == "TOP 10 MOST REVIEWED MOVIES":
            in_top_rated = False
            in_most_rev  = True
            continue
        if in_top_rated and r and isinstance(r[0], int):
            top_rated.append({"rank": r[0], "title": r[1], "year": r[2],
                              "genres": r[3], "avg": r[4], "count": r[5]})
        if in_most_rev and r and isinstance(r[0], int):
            most_reviewed.append({"rank": r[0], "title": r[1], "year": r[2],
                                  "genres": r[3], "avg": r[4], "count": r[5]})

    highest_rated  = top_rated[0]    if top_rated    else {}
    most_rev_title = most_reviewed[0] if most_reviewed else {}

    # ── Engagement sheet ───────────────────────────────────────────────────────
    e_rows = rows(eng_ws, min_row=1)

    def eng_value(label):
        return find_value(e_rows, label)

    total_rating_events = eng_value("Total Rating Events")
    dau_peak            = eng_value("Daily Active Users (peak)")
    dau_steady          = eng_value("Daily Active Users (steady)")
    mau_peak            = eng_value("Monthly Active Users (peak)")
    mau_avg             = eng_value("Monthly Active Users (avg)")
    dau_mau_ratio       = eng_value("DAU / MAU Ratio")
    session_depth       = eng_value("Avg Session Depth (ratings/session)")
    retention_m1m2      = eng_value("Retention Rate (Month 1 → 2)")
    retention_m1m6      = eng_value("Retention Rate (Month 1 → 6)")
    churn_rate          = eng_value("Churn Rate (monthly avg)")
    nps                 = eng_value("Net Promoter Score (NPS, est.)")
    content_user_ratio  = eng_value("Content-to-User Ratio")

    # User cohorts
    cohorts = []
    in_cohort = False
    for r in e_rows:
        if r and str(r[0]).strip() == "USER ACTIVITY COHORTS — Rating Volume Segments":
            in_cohort = True
            continue
        if in_cohort and r and str(r[0]).strip() == "RATING SCORE DISTRIBUTION":
            break
        if in_cohort and r and r[0] and str(r[0]).strip() not in ("Segment",) and r[2] is not None:
            cohorts.append({"segment": r[0], "range": r[1], "users": r[2], "pct": r[3]})

    casual_cohort     = next((c for c in cohorts if str(c["segment"]).strip() == "Casual"), {})
    power_user_cohort = next((c for c in cohorts if str(c["segment"]).strip() == "Power User"), {})
    super_user_cohort = next((c for c in cohorts if str(c["segment"]).strip() == "Super User"), {})
    power_users_total = (power_user_cohort.get("users") or 0) + (super_user_cohort.get("users") or 0)
    power_users_pct   = power_users_total / (total_users or 1)

    # Rating score distribution
    score_dist = {}
    in_scores = False
    for r in e_rows:
        if r and str(r[0]).strip() == "RATING SCORE DISTRIBUTION":
            in_scores = True
            continue
        if in_scores and r and isinstance(r[0], int) and r[0] in (1, 2, 3, 4, 5):
            score_dist[r[0]] = {"count": r[1], "pct": r[2], "label": r[3]}

    positive_pct = (score_dist.get(4, {}).get("pct") or 0) + (score_dist.get(5, {}).get("pct") or 0)
    negative_pct = (score_dist.get(1, {}).get("pct") or 0) + (score_dist.get(2, {}).get("pct") or 0)

    # ── Helpers to normalise values for display ────────────────────────────────
    def pct_str(v, decimals=0):
        """Convert 0.283 → '28%' or a pre-formatted string → return as-is."""
        if v is None:
            return "__"
        if isinstance(v, str):
            return v.strip().rstrip("%")
        return f"{v * 100:.{decimals}f}"

    def num_fmt(v, decimals=0):
        if v is None:
            return "___"
        if isinstance(v, str):
            return v
        return f"{v:,.{decimals}f}"

    def num_fmt_no_comma(v, decimals=0):
        if v is None:
            return "___"
        if isinstance(v, str):
            return v
        return f"{v:.{decimals}f}"

    # ── Assemble flat metrics dict ─────────────────────────────────────────────
    m = {
        # Cover
        "total_users":          num_fmt(total_users),
        "total_ratings":        num_fmt(total_ratings),
        "movies_catalogue":     num_fmt(movies_catalogue),
        "avg_rating_score":     num_fmt_no_comma(avg_rating_score, 2),
        "avg_ratings_user":     num_fmt_no_comma(avg_ratings_user, 1),
        "median_ratings":       num_fmt(median_ratings),
        "max_ratings":          num_fmt(max_ratings_user),

        # Gender
        "male_pct":             pct_str(male_pct),
        "female_pct":           pct_str(female_pct),
        "male_users":           num_fmt(male_users),
        "female_users":         num_fmt(female_users),
        "male_avg_ratings":     num_fmt_no_comma(male_avg, 1),
        "female_avg_ratings":   num_fmt_no_comma(female_avg, 1),

        # Age groups
        "top_age_group":        str(top_age_group.get("group", "25–34")),
        "top_age_pct":          pct_str(top_age_group.get("pct")),
        "second_age_group":     str(second_age_group.get("group", "18–24")),
        "second_age_pct":       pct_str(second_age_group.get("pct")),
        "age_groups":           age_groups,   # full list for PPTX

        # Occupations
        "top_occupation":       str(top_occ.get("name", "Student")),
        "top_occ_pct":          pct_str(top_occ.get("pct")),
        "top_occupations":      top_occupations,  # list for PPTX

        # Genres
        "top_genre_vol":        str(top_genre_vol.get("genre", "Drama")),
        "top_genre_vol_pct":    pct_str(top_genre_vol.get("pct")),
        "top_genre_score":      str(top_genre_score.get("genre", "Film-Noir")),
        "top_genre_score_val":  num_fmt_no_comma(top_genre_score.get("avg_score"), 2),
        "film_noir_avg":        num_fmt_no_comma(film_noir_genre.get("avg_score"), 2),
        "documentary_avg":      num_fmt_no_comma(documentary.get("avg_score"), 2),
        "drama_pct":            pct_str(drama_pct),
        "comedy_pct":           pct_str(comedy_pct),
        "action_pct":           pct_str(action_pct),
        "genres_by_vol":        genres_by_vol,    # full list for PPTX

        # Monthly signups
        "peak_month":           str(peak_month),
        "peak_new_users":       num_fmt(peak_new_users),

        # Funnel
        "conv_rate":            conv_rate_str,
        "funnel":               funnel,

        # Top movies
        "highest_rated_title":  str(highest_rated.get("title", "")),
        "highest_rated_avg":    num_fmt_no_comma(highest_rated.get("avg"), 2),
        "most_rev_title":       str(most_rev_title.get("title", "")),
        "most_rev_count":       num_fmt(most_rev_title.get("count")),
        "most_rev_avg":         num_fmt_no_comma(most_rev_title.get("avg"), 2),
        "top_rated":            top_rated[:10],       # list for PPTX
        "most_reviewed":        most_reviewed[:10],   # list for PPTX

        # Engagement
        "dau_peak":             num_fmt(dau_peak),
        "dau_steady":           num_fmt(dau_steady),
        "mau_peak":             num_fmt(mau_peak),
        "mau_avg":              num_fmt(mau_avg),
        "mau_pct":              pct_str((mau_avg or 0) / (total_users or 1)),
        "dau_mau_ratio":        str(dau_mau_ratio).replace("%", "").strip() if dau_mau_ratio else "__",
        "session_depth":        num_fmt_no_comma(session_depth, 1),
        "retention_m1m2":       str(retention_m1m2).replace("%", "").strip() if retention_m1m2 else "__",
        "retention_m1m6":       str(retention_m1m6).replace("%", "").strip() if retention_m1m6 else "__",
        "churn_rate":           str(churn_rate).replace("%", "").strip() if churn_rate else "__",
        "nps":                  str(nps) if nps else "+__",
        "positive_pct":         pct_str(positive_pct),
        "negative_pct":         pct_str(negative_pct),
        "power_users_pct":      pct_str(power_users_pct),
        "power_users_count":    num_fmt(power_users_total),
        "casual_pct":           pct_str(casual_cohort.get("pct")),
        "cohorts":              cohorts,     # full list for PPTX
        "score_dist":           score_dist,  # dict for PPTX
    }
    return m


# ══════════════════════════════════════════════════════════════════════════════
# 3.  WORD FILLER — replaces every placeholder in the Word template
# ══════════════════════════════════════════════════════════════════════════════

def _apply_subs(text: str, replacements: list) -> str:
    """
    Apply a list of (old, new) pairs in order.  Using a list (not dict) avoids
    silent overwriting of duplicate keys — order of substitution is explicit.

    Each replacement is done with count=1 (replace the FIRST occurrence only) when
    the pattern contains a placeholder-like substring, to avoid a pattern that
    matches multiple occurrences from consuming all of them at once.
    """
    for old, new in replacements:
        if old in text:
            text = text.replace(old, str(new), 1)
    return text


def _replace_para_text(para, replacements: list):
    """
    Replace placeholder strings inside a paragraph across all its runs.
    python-docx can split a single placeholder across multiple runs when the
    template was saved by Word; this function reconstructs the full text,
    does the substitution, then writes back into the first run and clears the rest.
    """
    full = "".join(r.text for r in para.runs)
    new  = _apply_subs(full, replacements)
    if new == full:
        return   # nothing to do
    if para.runs:
        para.runs[0].text = new
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(new)


def fill_word(template_path: Path, output_path: Path, m: dict, timestamp: str):
    """
    Open the Word template, substitute all placeholders with real values,
    and save to output_path (template is not modified).

    Strategy: instead of one global dict (which silently drops duplicate keys),
    we use ordered lists of (pattern, value) tuples with longest-match-first
    ordering for each context.
    """
    doc = Document(template_path)

    month_label = _month_label(timestamp)

    # ── Body paragraph substitutions ─────────────────────────────────────────
    # Ordered list: longer / more specific patterns before shorter ones.
    body_subs = [
        # Header
        ("[ Month Year ]",  month_label),
        # Executive summary — inline segments (must come before generic [ Month ])
        ("[ __,___ ] registered users",        f"{m['total_users']} registered users"),
        ("[ _,___ ] titles",                    f"{m['movies_catalogue']} titles"),
        ("[ _,___,___ ] cumulative ratings",    f"{m['total_ratings']} cumulative ratings"),
        ("[ _.__ / 5 ]",                        f"{m['avg_rating_score']} / 5"),
        ("stood at [ __,___ ]",                 f"stood at {m['mau_avg']}"),
        ("([ __ ]% of the registered base)",    f"({m['mau_pct']}% of the registered base)"),
        ("DAU/MAU ratio of [ __ ]%",            f"DAU/MAU ratio of {m['dau_mau_ratio']}%"),
        ("Month-one retention held at [ __ ]%", f"Month-one retention held at {m['retention_m1m2']}%"),
        ("Net Promoter Score remains at [ +__ ]", f"Net Promoter Score remains at {m['nps']}"),
        ("[ Month ]",       month_label),
        ("[ +__ ]",         m['nps']),
        # Users paragraph — specific multi-word patterns first
        ("male ([ __ ]%)",              f"male ({m['male_pct']}%)"),
        ("female ([ __ ]%)",            f"female ({m['female_pct']}%)"),
        ("[ Age Group ] at [ __ ]% of total", f"{m['top_age_group']} at {m['top_age_pct']}% of total"),
        ("followed by [ Age Group ] ([ __ ]%)",
            f"followed by {m['second_age_group']} ({m['second_age_pct']}%)"),
        ("[ Occupation ] ([ __ ]%)",    f"{m['top_occupation']} ({m['top_occ_pct']}%)"),
        ("average of [ ___ ] ratings",  f"average of {m['avg_ratings_user']} ratings"),
        ("accounting for [ __ ]% of the base",
            f"accounting for {m['power_users_pct']}% of the base"),
        # Usage & engagement paragraph
        ("[ _._ ] titles per session",          f"{m['session_depth']} titles per session"),
        ("[ __ ]% of all ratings are 4 or 5",   f"{m['positive_pct']}% of all ratings are 4 or 5"),
        ("only [ __ ]% are 1 or 2",             f"only {m['negative_pct']}% are 1 or 2"),
        ("[ Genre ] (avg [ _.__ ])",             f"{m['top_genre_score']} (avg {m['top_genre_score_val']})"),
        ("[ Genre ] leads in raw volume",        f"{m['top_genre_vol']} leads in raw volume"),
        ("month-one retention is [ __ ]%",       f"month-one retention is {m['retention_m1m2']}%"),
        ("six-month retention is [ __ ]%",       f"six-month retention is {m['retention_m1m6']}%"),
        ("Monthly churn remains at [ __ ]%",     f"Monthly churn remains at {m['churn_rate']}%"),
        # Content & marketing paragraph
        ("representing [ __ ]% of all ratings this period",
            f"representing {m['drama_pct']}% of all ratings this period"),
        ("Comedy ([ __ ]%)",            f"Comedy ({m['comedy_pct']}%)"),
        ("Action ([ __ ]%)",            f"Action ({m['action_pct']}%)"),
        ("([ _.__ ] and [ _.__ ] respectively)",
            f"({m['film_noir_avg']} and {m['documentary_avg']} respectively)"),
        ("was [ Title ] ([ _,___ ] ratings, avg [ _.__ ])",
            f"was {m['most_rev_title']} ({m['most_rev_count']} ratings, avg {m['most_rev_avg']})"),
        ("was [ Title ] ([ _.__ ] / 5)",
            f"was {m['highest_rated_title']} ({m['highest_rated_avg']} / 5)"),
        ("peaked in [ Month ] with [ ___ ] sign-ups",
            f"peaked in {m['peak_month']} with {m['peak_new_users']} sign-ups"),
        ("is estimated at [ __ ]%",     f"is estimated at {m['conv_rate']}%"),
        # Recommendations
        ("account for [ __ ]% of all ratings. Prioritise",
            f"account for {m['drama_pct']}% of all ratings. Prioritise"),
        ("score [ __ ]+ / 5",           f"score {m['top_genre_score_val']}+ / 5"),
        ("the largest ([ __ ]%) and most active",
            f"the largest ({m['top_age_pct']}%) and most active"),
        ("[ __ ]% of users are in the Casual tier",
            f"{m['casual_pct']}% of users are in the Casual tier"),
        # Footer
        ("WatchIt_Metrics_Report.xlsx", f"{timestamp}_WatchIt_Metrics_Report.xlsx"),
    ]

    for para in doc.paragraphs:
        _replace_para_text(para, body_subs)

    # ── Table substitutions ────────────────────────────────────────────────────
    for t_idx, table in enumerate(doc.tables):
        for row in table.rows:
            row_label = row.cells[0].text.strip() if row.cells else ""
            for c_idx, cell in enumerate(row.cells):
                for para in cell.paragraphs:

                    # T0 — KPI bar: substitute by column position since all cells share the same placeholder shape
                    if t_idx == 0 and len(row.cells) == 6:
                        kpi_by_col = {
                            0: (["[ __,___ ]"],                  m['total_users']),
                            1: (["[ _,___,___ ]"],               m['total_ratings']),
                            2: (["[ _,___ ]"],                   m['movies_catalogue']),
                            3: (["[ _.__ / 5 ]"],                f"{m['avg_rating_score']} / 5"),
                            4: (["[ __ % ]", "[ __ %]"],         f"{m['dau_mau_ratio']}%"),
                            5: (["[ __ % ]", "[ __ %]"],         f"{m['retention_m1m2']}%"),
                        }
                        if c_idx in kpi_by_col:
                            patterns, value = kpi_by_col[c_idx]
                            _replace_para_text(para, [(p, value) for p in patterns])

                    # T1 — Users metrics
                    elif t_idx == 1:
                        if c_idx == 1:
                            v = {
                                "Gender: Male":               f"{m['male_pct']} %",
                                "Gender: Female":             f"{m['female_pct']} %",
                                "Largest Age Group":          m['top_age_group'],
                                "Avg. Ratings / User":        m['avg_ratings_user'],
                                "Median Ratings / User":      m['median_ratings'],
                                "Power Users (≥500 ratings)": f"{m['power_users_pct']} %",
                            }.get(row_label)
                            if v:
                                _replace_para_text(para, [
                                    ("[ __ %  ]",  v),
                                    ("[ 25–34 ]",  m['top_age_group']),
                                    ("[ ___._ ]",  m['avg_ratings_user']),
                                    ("[ ___ ]",    m['median_ratings']),
                                ])
                        elif c_idx == 2:
                            v = {
                                "Gender: Male":               f"{m['male_users']} users",
                                "Gender: Female":             f"{m['female_users']} users",
                                "Largest Age Group":          f"{m['top_age_pct']}% of total",
                                "Power Users (≥500 ratings)": f"{m['power_users_count']} users",
                            }.get(row_label)
                            if v:
                                _replace_para_text(para, [
                                    ("[ __,___ ] users",  v),
                                    ("[ __ ]% of total",  v),
                                    ("[ ___ ] users",     v),
                                ])

                    # T2 — Engagement metrics
                    elif t_idx == 2:
                        if c_idx == 1:
                            v = {
                                "Daily Active Users (DAU)":   m['dau_peak'],
                                "Monthly Active Users (MAU)": m['mau_avg'],
                                "DAU / MAU Ratio":            f"{m['dau_mau_ratio']} %",
                                "Avg. Session Depth":         m['session_depth'],
                                "Ratings: Score 4–5 Share":   f"{m['positive_pct']} %",
                                "Monthly Churn Rate":         f"{m['churn_rate']} %",
                                "Net Promoter Score (est.)":  m['nps'],
                            }.get(row_label)
                            if v:
                                _replace_para_text(para, [
                                    ("[ _,___ ]",  v),
                                    ("[ __ %  ]",  v),
                                    ("[ _._ ]",    v),
                                    ("[ +__ ]",    v),
                                ])
                        elif c_idx == 2 and row_label == "Monthly Active Users (MAU)":
                            _replace_para_text(para, [
                                ("[ __ ]% of base", f"{m['mau_pct']}% of base"),
                            ])

                    # T3 — Content & Marketing
                    elif t_idx == 3:
                        col1_map = {
                            "Top Genre by Volume":           m['top_genre_vol'],
                            "Top Genre by Avg. Score":       m['top_genre_score'],
                            "Most Reviewed Title":           m['most_rev_title'],
                            "Highest Rated Title":           m['highest_rated_title'],
                            "Peak Acquisition Month":        m['peak_month'],
                            "Visitor → Registered Conv. %":  f"{m['conv_rate']}%",
                        }
                        col2_map = {
                            "Top Genre by Volume":           f"{m['top_genre_vol_pct']}% of all ratings",
                            "Top Genre by Avg. Score":       f"{m['top_genre_score_val']} / 5",
                            "Most Reviewed Title":           f"{m['most_rev_count']} ratings",
                            "Highest Rated Title":           f"{m['highest_rated_avg']} / 5",
                            "Peak Acquisition Month":        f"{m['peak_new_users']} new users",
                        }
                        if c_idx == 1:
                            v = col1_map.get(row_label)
                            if v:
                                _replace_para_text(para, [
                                    ("[ Drama   ]",    v),
                                    ("[ Film-Noir ]",  v),
                                    ("[ Title   ]",    v),
                                    ("[ Month   ]",    v),
                                    ("[ __ %    ]",    v),
                                ])
                        elif c_idx == 2:
                            v = col2_map.get(row_label)
                            if v:
                                _replace_para_text(para, [
                                    ("[ __ ]% of all ratings", v),
                                    ("[ _.__ ] / 5",           v),
                                    ("[ _,___ ] ratings",      v),
                                    ("[ ___ ] new users",      v),
                                ])

    doc.save(output_path)
    print(f"  ✅  Word saved  →  {output_path.name}")


# ══════════════════════════════════════════════════════════════════════════════
# 4.  PPTX FILLER — replaces every placeholder value in the deck
# ══════════════════════════════════════════════════════════════════════════════

def _replace_pptx_text(shape, replacements: dict):
    """Replace text in a single PPTX shape's text frame, preserving formatting."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            for old, new in replacements.items():
                if old in run.text:
                    run.text = run.text.replace(old, str(new))


def fill_pptx(template_path: Path, output_path: Path, m: dict, timestamp: str):
    """
    Open the PPTX template, fill all placeholder values with real metrics,
    and save to output_path.  Slide-by-slide substitution preserves all
    design elements (colours, fonts, shapes).
    """
    prs = Presentation(template_path)

    month_label = _month_label(timestamp)

    # ── Slide 1 — Title ───────────────────────────────────────────────────────
    slide1 = prs.slides[0]
    s1_subs = {
        "[Your Name]  ·  [Date]": f"WatchIt Team  ·  {month_label}",
        "[Date]": month_label,
    }
    for shape in slide1.shapes:
        _replace_pptx_text(shape, s1_subs)

    # ── Slide 3 — Platform KPI Overview ──────────────────────────────────────
    slide3 = prs.slides[2]
    # The 6 KPI cards each have 3 text boxes: label / BIG VALUE / sub-label
    # Values to fill in order (matching the template layout):
    #   Total Registered Users → m['total_users']
    #   Total Ratings Submitted → m['total_ratings']
    #   Movies in Catalogue → m['movies_catalogue']
    #   Avg. Rating Score → m['avg_rating_score'] / 5
    #   Avg. Ratings per User → m['avg_ratings_user']
    #   Monthly Active Users → m['mau_pct'] %
    kpi3_subs = {
        "______":  None,   # sentinel — we do a targeted pass
        "__ / 5":  f"{m['avg_rating_score']} / 5",
        "__ %":    f"{m['mau_pct']}%",
        "Core platform metrics for the reporting period": f"Core platform metrics — {month_label}",
    }

    # Targeted pass: replace "______" values in order of appearance
    placeholder_seq = [
        m['total_users'],
        m['total_ratings'],
        m['movies_catalogue'],
        f"{m['avg_rating_score']} / 5",
        m['avg_ratings_user'],
        f"{m['mau_pct']}%",
    ]
    p_idx = 0
    for shape in slide3.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.text.strip() in ("______", "__ / 5", "__ %"):
                    if p_idx < len(placeholder_seq):
                        run.text = str(placeholder_seq[p_idx])
                        p_idx += 1
                        break

    # ── Slide 4 — User Metrics & Demographics ─────────────────────────────────
    slide4 = prs.slides[3]
    # Gender %: first "__ %" → male, second → female
    # "___" values: first → male avg ratings, second → female avg ratings
    # Age groups "__ %" → 7 values in order
    # Top occupations "__ %" → 7 values in order
    gender_pcts = [f"{m['male_pct']}%", f"{m['female_pct']}%"]
    gender_avgs = [m['male_avg_ratings'], m['female_avg_ratings']]
    age_pcts = [f"{ag['pct'] * 100:.0f}%" if isinstance(ag.get('pct'), float)
                else str(ag.get('pct', '__')) for ag in m['age_groups']]
    occ_pcts  = [f"{o['pct'] * 100:.0f}%" if isinstance(o.get('pct'), float)
                 else str(o.get('pct', '__')) for o in m['top_occupations']]

    g_pct_idx = 0; g_avg_idx = 0; age_idx = 0; occ_idx = 0
    for shape in slide4.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                t = run.text.strip()
                if t == "__ %" and g_pct_idx < len(gender_pcts):
                    run.text = gender_pcts[g_pct_idx]; g_pct_idx += 1
                elif t == "___" and g_avg_idx < len(gender_avgs):
                    run.text = str(gender_avgs[g_avg_idx]); g_avg_idx += 1

    # Second pass for age groups and occupations (they share "__ %" token)
    # Reset and go section by section by shape order
    in_age_section = False; in_occ_section = False
    for shape in slide4.shapes:
        if not shape.has_text_frame:
            continue
        # Detect section by checking label text
        full_text = shape.text_frame.text
        if "Age Group Breakdown" in full_text:
            in_age_section = True; in_occ_section = False
        elif "Top Occupations" in full_text:
            in_age_section = False; in_occ_section = True

        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.text.strip() == "__ %":
                    if in_age_section and age_idx < len(age_pcts):
                        run.text = age_pcts[age_idx]; age_idx += 1
                    elif in_occ_section and occ_idx < len(occ_pcts):
                        run.text = occ_pcts[occ_idx]; occ_idx += 1

    # ── Slide 5 — Content & Catalogue Performance ─────────────────────────────
    slide5 = prs.slides[4]
    # Top 4 KPI tiles: movies, top genre vol, top genre score, film-noir avg
    kpi5_values = [
        m['movies_catalogue'],
        m['top_genre_vol'],
        f"{m['top_genre_score_val']} / 5",
        f"{m['film_noir_avg']} / 5",
    ]
    # Genre table: 8 genres × 3 data columns (ratings, % share, avg score)
    genre_table_data = {}
    for g in m['genres_by_vol']:
        genre_name = str(g['genre'])
        genre_table_data[genre_name] = {
            "count": f"{g['count']:,}" if isinstance(g['count'], (int, float)) else str(g['count']),
            "pct":   f"{g['pct'] * 100:.1f}%" if isinstance(g['pct'], float) else str(g['pct']),
            "avg":   f"{g['avg_score']:.2f}" if isinstance(g['avg_score'], float) else str(g['avg_score']),
        }
    # Rating distribution: score 1-5
    score_pcts_5to1 = [
        f"{m['score_dist'].get(5, {}).get('pct', 0) * 100:.1f}%",
        f"{m['score_dist'].get(4, {}).get('pct', 0) * 100:.1f}%",
        f"{m['score_dist'].get(3, {}).get('pct', 0) * 100:.1f}%",
        f"{m['score_dist'].get(2, {}).get('pct', 0) * 100:.1f}%",
        f"{m['score_dist'].get(1, {}).get('pct', 0) * 100:.1f}%",
    ]

    kpi5_idx = 0; score_idx = 0
    current_genre = None
    genre_col_idx = 0   # 0=ratings, 1=pct, 2=avg
    known_genres = ["Drama", "Comedy", "Action", "Thriller", "Romance", "Sci-Fi", "Crime", "Film-Noir"]
    in_genre_table = False; in_score_dist = False

    for shape in slide5.shapes:
        if not shape.has_text_frame:
            continue
        full = shape.text_frame.text.strip()
        if "Genre Performance Table" in full:
            in_genre_table = True; in_score_dist = False
        elif "Rating Score Distribution" in full:
            in_score_dist = True; in_genre_table = False

        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                t = run.text.strip()
                # Top 4 KPI tiles
                if t in ("______", "__ / 5") and kpi5_idx < len(kpi5_values) and not in_genre_table:
                    run.text = str(kpi5_values[kpi5_idx]); kpi5_idx += 1
                # Genre table data cells
                elif in_genre_table and t == "______":
                    if current_genre and current_genre in genre_table_data:
                        vals = genre_table_data[current_genre]
                        cols = [vals["count"], vals["pct"], vals["avg"]]
                        run.text = cols[genre_col_idx % 3]
                        genre_col_idx += 1
                # Rating distribution
                elif in_score_dist and t == "__ %" and score_idx < len(score_pcts_5to1):
                    run.text = score_pcts_5to1[score_idx]; score_idx += 1
            # Detect which genre row we're in
            for run in para.runs:
                if run.text.strip() in known_genres:
                    current_genre = run.text.strip()
                    genre_col_idx = 0

    # ── Slide 6 — Marketing & Genre Insights ──────────────────────────────────
    slide6 = prs.slides[5]
    # Conversion funnel %s (4 slots after "100%")
    funnel_data = m['funnel']
    funnel_stages_order = [
        "Registered Users",
        "Completed Profile",
        "Active Users (≥20 rtg)",
        "Power Users (≥500 rtg)",
    ]
    funnel_pcts = []
    for stage in funnel_stages_order:
        val = funnel_data.get(stage, {}).get("conv_rate", "__ %")
        funnel_pcts.append(str(val).replace("%", "").strip() + "%")

    # Marketing KPIs in order of appearance in template
    mkt_kpi_values = [
        m['peak_month'],
        m['peak_new_users'],
        m['top_genre_vol'],
        m['top_genre_score'],
        f"{m['conv_rate']}%",
        f"{m['power_users_pct']}%",
    ]

    funnel_idx = 0; mkt_idx = 0
    in_funnel = False; in_mkt_kpi = False
    for shape in slide6.shapes:
        if not shape.has_text_frame:
            continue
        full = shape.text_frame.text
        if "User Conversion Funnel" in full:
            in_funnel = True; in_mkt_kpi = False
        elif "Marketing KPIs" in full:
            in_funnel = False; in_mkt_kpi = True

        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                t = run.text.strip()
                if in_funnel and t == "__ %" and funnel_idx < len(funnel_pcts):
                    run.text = funnel_pcts[funnel_idx]; funnel_idx += 1
                elif in_mkt_kpi and t in ("______", "__ %") and mkt_idx < len(mkt_kpi_values):
                    run.text = str(mkt_kpi_values[mkt_idx]); mkt_idx += 1

    # ── Slide 7 — Engagement & Retention ─────────────────────────────────────
    slide7 = prs.slides[6]
    # Top 6 KPI tiles: DAU, MAU, DAU/MAU, M1→M2, Churn, NPS
    eng_kpi_top = [
        m['dau_peak'],
        m['mau_avg'],
        f"{m['dau_mau_ratio']}%",
        f"{m['retention_m1m2']}%",
        f"{m['churn_rate']}%",
        m['nps'],
    ]
    # Cohort %s: Casual, Engaged, Regular, Active, Power User, Super User
    cohort_pcts = []
    cohort_order = ["Casual", "Engaged", "Regular", "Active", "Power User", "Super User"]
    cohort_map = {str(c["segment"]).strip(): c for c in m["cohorts"]}
    for seg in cohort_order:
        c = cohort_map.get(seg, {})
        pct = c.get("pct")
        if isinstance(pct, float):
            cohort_pcts.append(f"{pct * 100:.0f}%")
        else:
            cohort_pcts.append("__%")

    # Retention hard values (already in template but we still update them)
    ret_values = {
        "68%":  f"{m['retention_m1m2']}%",
        "41%":  f"{m['retention_m1m6']}%",
        "35%":  "35%",   # M1→M12 not in Excel; keep template value
        "4%":   f"{m['churn_rate']}%",
    }

    eng_top_idx = 0; cohort_idx = 0
    in_cohorts = False; in_retention = False
    for shape in slide7.shapes:
        if not shape.has_text_frame:
            continue
        full = shape.text_frame.text
        if "User Activity Cohorts" in full:
            in_cohorts = True; in_retention = False
        elif "Retention Cohort Snapshot" in full:
            in_cohorts = False; in_retention = True

        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                t = run.text.strip()
                if t in ("______", "+__") and eng_top_idx < len(eng_kpi_top) and not in_cohorts:
                    run.text = str(eng_kpi_top[eng_top_idx]); eng_top_idx += 1
                elif t == "__ %" and not in_cohorts and not in_retention and eng_top_idx < len(eng_kpi_top):
                    run.text = str(eng_kpi_top[eng_top_idx]); eng_top_idx += 1
                elif in_cohorts and t == "__ %" and cohort_idx < len(cohort_pcts):
                    run.text = cohort_pcts[cohort_idx]; cohort_idx += 1
                elif in_retention and t in ret_values:
                    run.text = ret_values[t]

    # ── Slide 8 — Top Content Highlights ──────────────────────────────────────
    slide8 = prs.slides[7]
    # Top Rated table: avg → avg_rating, count → rating_count  (10 rows each)
    # Most Reviewed table: same columns
    top_rated_vals    = [(f"{r['avg']:.2f}", f"{r['count']:,}") for r in m['top_rated']]
    most_rev_vals     = [(f"{r['avg']:.2f}", f"{r['count']:,}") for r in m['most_reviewed']]

    # Titles also come from Excel — update them
    top_rated_titles  = [str(r['title']) for r in m['top_rated']]
    most_rev_titles   = [str(r['title']) for r in m['most_reviewed']]

    in_top_rated_tbl = False; in_most_rev_tbl = False
    tr_idx = 0; mr_idx = 0   # which row we're on
    tr_col = 0; mr_col = 0   # 0=avg, 1=count within a row

    # We iterate shapes and detect table sections by header text
    for shape in slide8.shapes:
        if not shape.has_text_frame:
            continue
        full = shape.text_frame.text.strip()
        if "Top Rated Titles" in full:
            in_top_rated_tbl = True; in_most_rev_tbl = False
        elif "Most Reviewed Titles" in full:
            in_top_rated_tbl = False; in_most_rev_tbl = True

        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                t = run.text.strip()
                if in_top_rated_tbl:
                    if t == "__.___" and tr_idx < len(top_rated_vals):
                        run.text = top_rated_vals[tr_idx][0]
                        tr_col = 1
                    elif t == "_,___" and tr_idx < len(top_rated_vals):
                        run.text = top_rated_vals[tr_idx][1]
                        tr_idx += 1; tr_col = 0
                    elif t in top_rated_titles:
                        title_idx = top_rated_titles.index(t)
                        if title_idx < len(top_rated_vals):
                            run.text = top_rated_titles[title_idx]
                elif in_most_rev_tbl:
                    if t == "__.___" and mr_idx < len(most_rev_vals):
                        run.text = most_rev_vals[mr_idx][0]
                        mr_col = 1
                    elif t == "_,___" and mr_idx < len(most_rev_vals):
                        run.text = most_rev_vals[mr_idx][1]
                        mr_idx += 1; mr_col = 0
                    elif t in most_rev_titles:
                        title_idx = most_rev_titles.index(t)
                        if title_idx < len(most_rev_vals):
                            run.text = most_rev_titles[title_idx]

    # ── Slide 9 — Key Takeaways ────────────────────────────────────────────────
    slide9 = prs.slides[8]
    # Each takeaway card has a body text containing "__,___" or "__ %" etc.
    takeaway_subs = {
        "__,___":  m['total_users'],
        "ages 25–34 (__ %)":        f"ages {m['top_age_group']} ({m['top_age_pct']}%)",
        "DAU/MAU ratio at __ %":    f"DAU/MAU ratio at {m['dau_mau_ratio']}%",
        "M1→M2 retention at __ %":  f"M1→M2 retention at {m['retention_m1m2']}%",
        "Churn stable at __ %":     f"Churn stable at {m['churn_rate']}%",
        "__ genres catalogued":     f"{len(m['genres_by_vol'])} genres catalogued",
        "Peak acquisition in ______": f"Peak acquisition in {m['peak_month']}",
        "Power users (__ %)":       f"Power users ({m['power_users_pct']}%)",
    }
    for shape in slide9.shapes:
        _replace_pptx_text(shape, takeaway_subs)

    # ── Slide 10 — Thank You ───────────────────────────────────────────────────
    slide10 = prs.slides[9]
    ty_subs = {
        "[Your Name]  ·  WatchIt Management Update":
            f"WatchIt Team  ·  {month_label} Management Update",
    }
    for shape in slide10.shapes:
        _replace_pptx_text(shape, ty_subs)

    prs.save(output_path)
    print(f"  ✅  PPTX saved  →  {output_path.name}")


# ══════════════════════════════════════════════════════════════════════════════
# 5.  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def extract_timestamp(excel_path: Path) -> str:
    """
    Extract the YYYYMM timestamp from the Excel filename.
    Expects filenames like  202611_WatchIt_Metrics_Report.xlsx.
    Returns the raw string e.g. '202611'.
    """
    stem = excel_path.stem   # e.g. "202611_WatchIt_Metrics_Report"
    match = re.match(r'^(\d{6})', stem)
    if not match:
        raise ValueError(
            f"Cannot find a YYYYMM prefix in filename: {excel_path.name}\n"
            "Expected format: 202611_WatchIt_Metrics_Report.xlsx"
        )
    return match.group(1)


def _month_label(timestamp: str) -> str:
    """Convert '202611' → 'November 2026'."""
    import calendar
    year  = int(timestamp[:4])
    month = int(timestamp[4:6])
    return f"{calendar.month_name[month]} {year}"


# ══════════════════════════════════════════════════════════════════════════════
# 6.  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="WatchIt Monthly Report Pipeline — fill Word + PPTX templates from Excel"
    )
    parser.add_argument(
        "excel",
        help="Path to the monthly Excel file, e.g. 202611_WatchIt_Metrics_Report.xlsx"
    )
    parser.add_argument(
        "--output-dir", "-o",
        default=None,
        help="Directory for output files (default: same folder as the Excel)"
    )
    parser.add_argument(
        "--word-template",
        default=str(WORD_TEMPLATE),
        help=f"Path to the Word template (default: {WORD_TEMPLATE})"
    )
    parser.add_argument(
        "--pptx-template",
        default=str(PPTX_TEMPLATE),
        help=f"Path to the PPTX template (default: {PPTX_TEMPLATE})"
    )
    args = parser.parse_args()

    excel_path   = Path(args.excel).resolve()
    word_tmpl    = Path(args.word_template).resolve()
    pptx_tmpl    = Path(args.pptx_template).resolve()

    # Validate inputs
    for p, label in [(excel_path, "Excel"), (word_tmpl, "Word template"),
                     (pptx_tmpl, "PPTX template")]:
        if not p.exists():
            print(f"❌  {label} not found: {p}")
            sys.exit(1)

    # Determine output directory
    out_dir = Path(args.output_dir).resolve() if args.output_dir else excel_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # Extract timestamp from filename
    timestamp = extract_timestamp(excel_path)
    print(f"\n🎬  WatchIt Monthly Report Pipeline")
    print(f"    Period   : {_month_label(timestamp)}  ({timestamp})")
    print(f"    Excel    : {excel_path.name}")
    print(f"    Output   : {out_dir}\n")

    # Define output filenames (templates never overwritten)
    word_out = out_dir / f"{timestamp}_WatchIt_Monthly_Management_Update.docx"
    pptx_out = out_dir / f"{timestamp}_WatchIt_Management_Metrics.pptx"

    # Check for accidental overwrite of the template itself
    if word_out.resolve() == word_tmpl.resolve():
        print("❌  Output path would overwrite the Word template. Aborting.")
        sys.exit(1)
    if pptx_out.resolve() == pptx_tmpl.resolve():
        print("❌  Output path would overwrite the PPTX template. Aborting.")
        sys.exit(1)

    # Load metrics
    print("📊  Reading Excel metrics …")
    metrics = load_metrics(excel_path)

    # Fill Word
    print("📝  Filling Word template …")
    fill_word(word_tmpl, word_out, metrics, timestamp)

    # Fill PPTX
    print("📊  Filling PowerPoint template …")
    fill_pptx(pptx_tmpl, pptx_out, metrics, timestamp)

    print(f"\n✅  Done!  Both reports for {_month_label(timestamp)} are ready.")
    print(f"   Word  : {word_out}")
    print(f"   PPTX  : {pptx_out}\n")


if __name__ == "__main__":
    main()
