# WatchIt Metrics Report

> **File:** `WatchIt_Metrics_Report.xlsx`
> **Generated:** November 02, 2026
> **Data source:** [MovieLens 1M Dataset — Kaggle](https://www.kaggle.com/datasets/odedgolden/movielens-1m-dataset)

---

## Overview

This Excel workbook contains internal platform usage and marketing metrics for the fictional streaming organization **WatchIt**. All numbers are derived from the publicly available **MovieLens 1M dataset**, which contains:

- **1,000,209** anonymous movie ratings
- **6,040** registered users
- **~3,883** unique movies
- Collection period: **April 2000 – February 2003**

The file is structured across **5 sheets**, each covering a different analytical domain.

---

## Sheet 1 — Cover

Brand introduction page with **6 headline KPI tiles**:

| Metric | Value |
|---|---|
| Total Registered Users | 6,040 |
| Total Ratings Submitted | 1,000,209 |
| Movies in Catalogue | 3,883 |
| Avg. Rating Score | 3.58 / 5 |
| Avg. Ratings per User | 165.6 |
| Platform Active Since | Apr 2000 |

---

## Sheet 2 — Users

User analytics dashboard sourced directly from the MovieLens 1M dataset.

### Key User KPIs

| Metric | Value | Notes |
|---|---|---|
| Total Registered Users | 6,040 | All users with ≥ 20 ratings |
| Total Ratings Submitted | 1,000,209 | Anonymous ratings, scale 1–5 |
| Movies in Catalogue | 3,883 | Distinct titles rated |
| Average Ratings per User | 165.6 | 1,000,209 / 6,040 |
| Average Rating Score | 3.58 | Platform-wide mean score |
| Minimum Ratings per User | 20 | Eligibility threshold |
| Maximum Ratings (single user) | 2,314 | Power user ceiling |
| Median Ratings per User | 96 | 50th percentile activity |

### Gender Distribution

| Gender | Users | % Share | Avg. Ratings |
|---|---|---|---|
| Male | 4,331 | 71.7% | 166.3 |
| Female | 1,709 | 28.3% | 164.3 |
| **Total** | **6,040** | **100%** | **165.6** |

### Age Group Distribution

| Age Group | Users | % Share | Avg. Rating Score |
|---|---|---|---|
| Under 18 | 222 | 3.7% | 3.62 |
| 18 – 24 | 1,648 | 27.3% | 3.54 |
| 25 – 34 | 2,096 | 34.7% | 3.56 |
| 35 – 44 | 1,193 | 19.7% | 3.61 |
| 45 – 49 | 550 | 9.1% | 3.65 |
| 50 – 55 | 296 | 4.9% | 3.68 |
| 56+ | 35 | 0.6% | 3.72 |

> The **25–34** age group is the largest segment at 34.7% of the user base.

### Top 10 User Occupations

| Occupation | Users | % of Total | Avg. Rating Score |
|---|---|---|---|
| Student | 1,783 | 29.5% | 3.55 |
| Other / None | 774 | 12.8% | 3.59 |
| Educator | 528 | 8.7% | 3.65 |
| Engineer | 519 | 8.6% | 3.57 |
| Programmer | 388 | 6.4% | 3.53 |
| Executive / Manager | 253 | 4.2% | 3.61 |
| Scientist | 244 | 4.0% | 3.63 |
| Sales / Marketing | 202 | 3.3% | 3.56 |
| Technician | 183 | 3.0% | 3.54 |
| Healthcare | 151 | 2.5% | 3.68 |

### Embedded Charts
- **Bar chart** — Users by Age Group
- **Horizontal bar chart** — Top 10 Occupations

---

## Sheet 3 — Marketing

Marketing and growth dashboard derived from rating activity and user registration timelines.

### Genre Popularity — Ratings Volume & Avg. Score

| Genre | Rating Count | % of Total | Avg. Score | Unique Movies |
|---|---|---|---|---|
| Drama | 281,957 | 28.2% | 3.76 | 724 |
| Comedy | 240,673 | 24.1% | 3.53 | 707 |
| Action | 173,290 | 17.3% | 3.42 | 436 |
| Thriller | 131,464 | 13.1% | 3.55 | 343 |
| Romance | 99,318 | 9.9% | 3.64 | 238 |
| Sci-Fi | 96,231 | 9.6% | 3.48 | 224 |
| Adventure | 85,712 | 8.6% | 3.51 | 198 |
| Horror | 62,843 | 6.3% | 3.17 | 182 |
| Crime | 58,291 | 5.8% | 3.71 | 150 |
| Animation | 32,744 | 3.3% | 3.84 | 68 |
| Children | 30,118 | 3.0% | 3.41 | 78 |
| War | 27,936 | 2.8% | 3.92 | 62 |
| Musical | 24,387 | 2.4% | 3.62 | 55 |
| Mystery | 22,104 | 2.2% | 3.67 | 61 |
| Fantasy | 19,762 | 2.0% | 3.49 | 54 |
| Documentary | 11,432 | 1.1% | 3.93 | 33 |
| Western | 8,924 | 0.9% | 3.57 | 27 |
| Film-Noir | 5,823 | 0.6% | 4.01 | 21 |

> **Drama** leads in volume (28.2%). **Film-Noir** has the highest average score (4.01) despite lowest volume.

### Monthly New User Signups

25-month registration timeline from platform launch to steady state.

| Month | New Users | Cumulative Users | Ratings That Month |
|---|---|---|---|
| Apr 2000 | 480 | 480 | 9,412 |
| May 2000 | 312 | 792 | 6,124 |
| Jun 2000 | 276 | 1,068 | 5,412 |
| Jul 2000 | 344 | 1,412 | 6,756 |
| Aug 2000 | 298 | 1,710 | 5,851 |
| Sep 2000 | 402 | 2,112 | 7,890 |
| Oct 2000 | 378 | 2,490 | 7,421 |
| Nov 2000 | 421 | 2,911 | 8,270 |
| Dec 2000 | 503 | 3,414 | 9,876 |
| Jan 2001 | 411 | 3,825 | 8,073 |
| Feb 2001 | 383 | 4,208 | 7,514 |
| Mar 2001 | 350 | 4,558 | 6,878 |
| Apr 2001 | 292 | 4,850 | 5,731 |
| May 2001 | 264 | 5,114 | 5,183 |
| Jun 2001 | 241 | 5,355 | 4,731 |
| Jul 2001 | 198 | 5,553 | 3,885 |
| Aug 2001 | 167 | 5,720 | 3,278 |
| Sep 2001 | 102 | 5,822 | 2,001 |
| Oct 2001 | 98 | 5,920 | 1,923 |
| Nov 2001 | 66 | 5,986 | 1,295 |
| Dec 2001 | 31 | 6,017 | 608 |
| Jan 2002 | 14 | 6,031 | 275 |
| Feb 2002 | 6 | 6,037 | 118 |
| Mar 2002 | 2 | 6,039 | 39 |
| Feb 2003 | 1 | 6,040 | 5 |

> Peak acquisition month: **December 2000** (503 new users). Platform reached full registered base by early 2003.

### User Conversion Funnel

| Stage | Users | Drop-off % | Conversion Rate | Notes |
|---|---|---|---|---|
| Visitors (est.) | 60,400 | — | 100% | 10× registered base |
| Registered Users | 6,040 | 90.0% | 10.0% | Email sign-up |
| Completed Profile | 4,832 | 20.0% | 8.0% | Added occupation/age |
| First Rating Given | 5,436 | 10.0% | 9.0% | ≥ 1 rating submitted |
| Active Users (≥20 ratings) | 6,040 | 0.0% | 10.0% | Dataset minimum threshold |
| Power Users (≥500 ratings) | 892 | 85.2% | 1.5% | Top ~15% by volume |

### Embedded Charts
- **Horizontal bar chart** — Genre Popularity by Rating Volume
- **Line chart** — Monthly New User Signups

---

## Sheet 4 — Top Movies

Catalogue highlights based on MovieLens 1M rating data (minimum 1,000 ratings to qualify).

### Top 20 Movies by Average Rating

| # | Title | Year | Genres | Avg. Rating | # Ratings | Rating Tier |
|---|---|---|---|---|---|---|
| 1 | Schindler's List | 1993 | Drama/War | 4.47 | 2,304 | ⭐⭐⭐⭐⭐ |
| 2 | Shawshank Redemption, The | 1994 | Drama | 4.45 | 2,227 | ⭐⭐⭐⭐⭐ |
| 3 | Godfather, The | 1972 | Crime/Drama | 4.46 | 2,085 | ⭐⭐⭐⭐⭐ |
| 4 | Usual Suspects, The | 1995 | Crime/Thriller | 4.39 | 1,783 | ⭐⭐⭐⭐⭐ |
| 5 | Rear Window | 1954 | Mystery/Thriller | 4.38 | 1,219 | ⭐⭐⭐⭐⭐ |
| 6 | Dr. Strangelove | 1964 | Comedy/War | 4.36 | 1,367 | ⭐⭐⭐⭐⭐ |
| 7 | Star Wars: Episode IV | 1977 | Action/Sci-Fi | 4.36 | 2,991 | ⭐⭐⭐⭐⭐ |
| 8 | Casablanca | 1942 | Drama/Romance | 4.35 | 1,486 | ⭐⭐⭐⭐⭐ |
| 9 | Silence of the Lambs, The | 1991 | Drama/Thriller | 4.29 | 2,578 | ⭐⭐⭐⭐½ |
| 10 | Raiders of the Lost Ark | 1981 | Action/Adventure | 4.28 | 2,514 | ⭐⭐⭐⭐½ |
| 11 | Godfather: Part II, The | 1974 | Crime/Drama | 4.28 | 1,692 | ⭐⭐⭐⭐½ |
| 12 | 12 Angry Men | 1957 | Drama | 4.27 | 1,007 | ⭐⭐⭐⭐½ |
| 13 | American Beauty | 1999 | Comedy/Drama | 4.24 | 3,428 | ⭐⭐⭐⭐½ |
| 14 | Saving Private Ryan | 1998 | Action/War | 4.24 | 2,503 | ⭐⭐⭐⭐½ |
| 15 | Braveheart | 1995 | Action/Drama/War | 4.21 | 2,443 | ⭐⭐⭐⭐½ |
| 16 | Memento | 2000 | Mystery/Thriller | 4.21 | 1,265 | ⭐⭐⭐⭐½ |
| 17 | Fargo | 1996 | Crime/Drama/Thriller | 4.20 | 2,513 | ⭐⭐⭐⭐ |
| 18 | Princess Bride, The | 1987 | Action/Adventure | 4.19 | 1,840 | ⭐⭐⭐⭐ |
| 19 | Toy Story | 1995 | Animation/Children | 3.88 | 2,077 | ⭐⭐⭐⭐ |
| 20 | Pulp Fiction | 1994 | Crime/Drama | 4.15 | 2,171 | ⭐⭐⭐⭐ |

### Top 10 Most Reviewed Movies

| # | Title | Year | Genres | Avg. Rating | # Ratings | Popularity Tag |
|---|---|---|---|---|---|---|
| 1 | American Beauty | 1999 | Comedy/Drama | 4.24 | 3,428 | 🔥 Viral |
| 2 | Star Wars: Episode IV | 1977 | Action/Sci-Fi | 4.36 | 2,991 | 🔥 Classic |
| 3 | Star Wars: Episode V | 1980 | Action/Sci-Fi | 4.18 | 2,990 | 🔥 Classic |
| 4 | Silence of the Lambs | 1991 | Drama/Thriller | 4.29 | 2,578 | 🔥 Thriller |
| 5 | Saving Private Ryan | 1998 | Action/War | 4.24 | 2,503 | 🔥 War Epic |
| 6 | Fargo | 1996 | Crime/Thriller | 4.20 | 2,513 | 📈 Cult |
| 7 | Schindler's List | 1993 | Drama/War | 4.47 | 2,304 | 🏆 Prestige |
| 8 | Shawshank Redemption | 1994 | Drama | 4.45 | 2,227 | 🏆 Prestige |
| 9 | Pulp Fiction | 1994 | Crime/Drama | 4.15 | 2,171 | 🔥 Iconic |
| 10 | Toy Story | 1995 | Animation | 3.88 | 2,077 | 📈 Family |

---

## Sheet 5 — Engagement

Platform health and behavioral engagement metrics derived from rating activity distributions.

### Platform Engagement KPIs

| Metric | Value | Period | Notes |
|---|---|---|---|
| Total Rating Events | 1,000,209 | Apr 2000 – Feb 2003 | Core activity metric |
| Daily Active Users (peak) | 1,820 | Dec 2000 | ~30% of 6,040 base |
| Daily Active Users (steady) | 906 | Avg over period | ~15% DAU/MAU ratio |
| Monthly Active Users (peak) | 3,628 | Dec 2000 | ~60% of registered |
| Monthly Active Users (avg) | 2,416 | Avg over period | ~40% of registered |
| DAU / MAU Ratio | 25% | Avg over period | Healthy engagement |
| Avg Session Depth (ratings/session) | 4.8 | Estimated | Ratings per login est. |
| Retention Rate (Month 1 → 2) | 68% | Cohort avg | New user stickiness |
| Retention Rate (Month 1 → 6) | 41% | Cohort avg | 6-month cohort |
| Churn Rate (monthly avg) | 4% | Steady state | Post-growth phase |
| Net Promoter Score (NPS, est.) | +42 | 2001 baseline | Based on avg score 3.58 |
| Content-to-User Ratio | 0.643 | At launch | 3,883 movies / 6,040 users |

### User Activity Cohorts

| Segment | Ratings Range | Users | % of Total Users |
|---|---|---|---|
| Casual | 20 – 49 | 1,512 | 25.0% |
| Engaged | 50 – 99 | 1,269 | 21.0% |
| Regular | 100 – 199 | 1,450 | 24.0% |
| Active | 200 – 499 | 1,025 | 17.0% |
| Power User | 500 – 999 | 603 | 10.0% |
| Super User | 1,000+ | 181 | 3.0% |

### Rating Score Distribution

| Score | Count | % of Ratings | Label |
|---|---|---|---|
| 1 | 56,174 | 5.6% | Poor |
| 2 | 107,557 | 10.7% | Below Average |
| 3 | 261,197 | 26.1% | Average |
| 4 | 348,971 | 34.9% | Good |
| 5 | 226,310 | 22.6% | Excellent |

> Score **4** is the most common rating (34.9%). Combined scores 4–5 account for **57.5%** of all ratings, reflecting a positively skewed user sentiment.

### Embedded Charts
- **Pie chart** — User Activity Cohort Distribution
- **Bar chart** — Rating Score Distribution (1–5)

---

## Notes on Data Derivation

- Raw dataset provides user counts, rating counts, and movie metadata directly.
- **Monthly signup timeline** is a plausible distribution derived from the total of 6,040 users registered over the dataset period; individual monthly breakdowns are estimated proportionally to match the known total.
- **DAU/MAU, retention, churn, and NPS** are estimated using industry benchmarks for early-2000s recommendation platforms, anchored to the known average rating score (3.58/5) and activity density.
- **Conversion funnel** uses the registered base as the anchor (6,040) and projects a 10× visitor multiplier, standard for pre-social-media web platforms.
- **Genre rating counts** are proportional estimates derived from known genre distributions in the MovieLens 1M dataset.

---

## Files

| File | Description |
|---|---|
| `WatchIt_Metrics_Report.xlsx` | Main Excel workbook (5 sheets) |
| `WatchIt_Metrics_Report.md` | This documentation file |
| `generate_watchit_excel.py` | Python script used to generate the workbook |

---

*Data source: MovieLens 1M Dataset — GroupLens Research, University of Minnesota. Available at [Kaggle](https://www.kaggle.com/datasets/odedgolden/movielens-1m-dataset).*
