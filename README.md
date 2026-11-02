# WatchIt — Monthly Management Update: Corpus Interaction Use Case

**Author:** Evgeniya Englert  
**Project:** WatchIt Platform — Automated Monthly Reporting  
**Stack:** Python · Microsoft Word (DOCX) · PowerPoint (PPTX) · Excel (XLSX) · Kiro CLI  
**Last Updated:** November 2026

---

## Table of Contents

1. [Business Problem](#1-business-problem)
2. [What We Had at the Start](#2-what-we-had-at-the-start)
3. [Technical Problem — Corpus Interaction](#3-technical-problem--corpus-interaction)
4. [Impact](#4-impact)
5. [Technical Architecture](#5-technical-architecture)
6. [Step-by-Step: What Was Built](#6-step-by-step-what-was-built)
7. [Next Steps & Recommendations](#7-next-steps--recommendations)
8. [Generalising This Pattern — Corpus Interaction as a Framework](#8-generalising-this-pattern--corpus-interaction-as-a-framework)

---

## 1. Business Problem

WatchIt is a streaming platform that collects user behaviour data — ratings, registrations, genre preferences, retention, and engagement signals — on a monthly basis. This data is used to keep management and the board informed about platform health, audience composition, and strategic priorities.

Every month, a **management update** must be produced in two formats:

| Format | Purpose | Audience |
|--------|---------|----------|
| **Word document** (2 pages) | Narrative management update with data, analysis, and recommendations | Board, senior management |
| **PowerPoint deck** (10 slides) | Visual presentation of platform metrics | Management meetings, board reviews |

### The Pain

Before this project, the process looked like this:

1. Data analyst exports the monthly Excel with ~40 metrics across 5 sheets.
2. A person manually reads the Excel, opens the Word document, and copy-pastes or re-types every number into the narrative.
3. The same person opens the PowerPoint deck and manually updates every KPI tile, chart label, table row, and takeaway card.
4. Both documents are proofread, formatted, and shared.

This process took **2–4 hours per month**, introduced **copy-paste errors**, required **formatting fixes** after every edit, and produced documents that looked inconsistent month to month. The business context (professional narrative, strategic framing, design palette) had to be rebuilt from scratch every cycle.

### The Goal

> **Produce a fully populated, professionally designed Word and PPTX management update automatically from a monthly Excel file, in under 10 seconds, with zero manual data entry.**

---

## 2. What We Had at the Start

### 2.1 Monthly Excel: `YYYYMM_WatchIt_Metrics_Report.xlsx`

A structured Excel workbook refreshed every month. The filename prefix encodes the reporting period (`202611` = November 2026). It contains **5 sheets**:

| Sheet | Content |
|-------|---------|
| `Cover` | Platform-level headline KPIs (users, ratings, catalogue, avg score) |
| `Users` | User count, gender split, age group distribution, top 10 occupations |
| `Marketing` | Genre performance table, monthly signup trend, conversion funnel |
| `Top Movies` | Top 20 highest-rated and top 10 most-reviewed titles |
| `Engagement` | DAU/MAU, session depth, retention cohorts, rating score distribution |

The Excel follows a **consistent schema** every month — same sheet names, same section labels, same column structure. This consistency is what makes automation possible.

### 2.2 Word Template: `templates/WatchIt_Monthly_Management_Update_Template.docx`

A professional 2-page Word document with WatchIt branding (navy/gold/Aptos). It contains:

- A branded header with reporting period
- 5 narrative sections (Executive Summary, Users, Usage & Engagement, Content & Marketing, Strategy Recommendations)
- 4 structured data tables
- A KPI bar with 6 headline figures
- Key points box and footer

All data values are represented as **`[ placeholder ]`** strings — for example: `[ __,___ ] registered users`, `[ __ ]%`, `[ Title ]`. These are the injection points.

### 2.3 PowerPoint Template: `templates/WatchIt_Monthly_Management_Update_Template.pptx`

A 10-slide professionally designed deck matching the WatchIt palette. Slides cover:

1. Title
2. Agenda
3. Platform KPI Overview (6 KPI tiles)
4. User Metrics & Demographics (3 panels)
5. Content & Catalogue Performance (genre table + score distribution)
6. Marketing & Genre Insights (funnel + KPI table)
7. Engagement & Retention (6 KPI tiles + cohort panel)
8. Top Content Highlights (two 10-row tables)
9. Key Takeaways & Next Steps
10. Thank You

All numeric values are placeholder strings: `______`, `__ %`, `__.___`, `_,___`, etc.

---

## 3. Technical Problem — Corpus Interaction

### What Is Corpus Interaction?

**Corpus Interaction** is the practice of programmatically reading, processing, and writing structured business documents — treating Word files, Excel sheets, and PowerPoint decks as **data sources and targets** rather than just files to open manually.

In this project, corpus interaction means:

- **Reading** the Excel corpus (structured tabular data with named sections)
- **Extracting** the right values from the right cells based on section labels
- **Injecting** those values into the correct locations in two different document corpora (DOCX and PPTX)
- **Preserving** all design elements — fonts, colours, table styles, slide layouts — while only changing the data cells

### The Technical Challenges

The documents are not simple templates. Each format presented specific challenges:

**Excel → Python:**
- Data is organised by *named section headers* (e.g. `"GENDER DISTRIBUTION"`, `"USER ACTIVITY COHORTS"`), not by fixed cell addresses
- Values need to be found by label lookup, not by hardcoded `B5`-style references
- Percentages are stored as floats (0.717) and need formatting as `"72%"` for display
- Multiple data types (int, float, str) require normalisation before injection

**Python → Word (DOCX):**
- A single visible placeholder like `[ __,___ ]` may be **split across multiple XML runs** due to Word's internal formatting model. The text must be reconstructed from all runs before matching.
- The same placeholder pattern (e.g. `[ __ ]%`) appears multiple times in a document with **different intended values** — requiring ordered, context-aware substitution rather than a simple global find-replace.
- Tables and paragraph body text must be processed separately, with per-table, per-row, per-column context to assign the right value to the right cell.
- After substitution, the combined text must be written back to `run[0]` while clearing the other runs — this collapses the formatting to the run-0 style, which is acceptable for data cells.

**Python → PPTX:**
- PowerPoint stores text in **shape objects**, each with a text frame containing paragraphs and runs. There is no global find-replace.
- The same placeholder string (e.g. `__ %`) appears across multiple slides for completely different metrics. Substitution must be **slide-scoped** and **shape-context-aware**.
- Ordered sequential substitution is used within each slide — the first occurrence of `______` on Slide 3 is `total_users`, the second is `total_ratings`, etc.
- Gender/age/occupation sections on Slide 4 are detected by section-header shapes and processed in sequence.

**The Key Insight — Ordered Substitution:**  
Python dictionaries silently overwrite duplicate keys. Using an **ordered list of `(pattern, value)` tuples** with `str.replace(pattern, value, count=1)` (replace first occurrence only) solves both the duplicate-key problem and the "same pattern, different context" problem in a single, debuggable mechanism.

---

## 4. Impact

### Time Savings

| Task | Before (manual) | After (automated) |
|------|-----------------|-------------------|
| Read Excel, extract ~40 metrics | ~30 min | 0 min |
| Populate Word document | ~60 min | 0 min |
| Populate PowerPoint deck | ~60 min | 0 min |
| Proofread for copy-paste errors | ~30 min | ~5 min (editorial only) |
| Fix formatting after edits | ~20 min | 0 min |
| **Total per month** | **~3.5 hours** | **~5 minutes** |

**Monthly time saving: ~3 hours 25 minutes**  
**Annual time saving: ~41 hours (~5 working days)**

### Quality Improvements

- **Zero copy-paste errors** — all values come directly from the Excel source
- **Consistent formatting** — output always matches the branded template exactly
- **Reproducible** — the same Excel always produces the same document, every time
- **Auditable** — the pipeline script is version-controlled; any change is traceable
- **Timestamped outputs** — every month's reports are archived automatically (`202611_`, `202612_`, ...)
- **Templates preserved** — the originals are never modified; a bad run never corrupts the base

### Strategic Value

- The reporting function becomes **independent of who does it** — any team member can run `python pipeline/watchit_monthly_pipeline.py 202612_WatchIt_Metrics_Report.xlsx`
- Monthly reporting becomes a **CI/CD-style pipeline** rather than a manual craft
- The pattern is **reusable** — any Excel-to-Word/PPTX reporting need can be solved with the same architecture

---

## 5. Technical Architecture

### 5.1 System Components

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                        WATCHIT MONTHLY REPORT PIPELINE                      │
└─────────────────────────────────────────────────────────────────────────────┘

  INPUT                         CORE ENGINE                        OUTPUT
  ─────                         ───────────                        ──────

  ┌──────────────────┐          ┌───────────────────────┐          ┌─────────────────────────────┐
  │  YYYYMM_WatchIt  │  parse   │                       │  fill    │  YYYYMM_WatchIt_Monthly     │
  │  _Metrics_Report │ ──────►  │   load_metrics()      │ ──────►  │  _Management_Update.docx    │
  │  .xlsx           │          │                       │          └─────────────────────────────┘
  └──────────────────┘          │   Returns flat dict   │
                                │   of ~40 metrics      │          ┌─────────────────────────────┐
  ┌──────────────────┐          │                       │  fill    │  YYYYMM_WatchIt_Management  │
  │  Word Template   │ ──────►  │   fill_word()         │ ──────►  │  _Metrics.pptx              │
  │  (placeholder)   │          │   fill_pptx()         │          └─────────────────────────────┘
  └──────────────────┘          │                       │
                                └───────────────────────┘
  ┌──────────────────┐
  │  PPTX Template   │ ──────►  (copied and filled, never overwritten)
  │  (placeholder)   │
  └──────────────────┘
```

### 5.2 Data Flow

```
Excel (5 sheets)
       │
       ▼
  load_metrics()
  ┌─────────────────────────────────────────────────────────┐
  │  Sheet: Cover      → total_users, total_ratings,        │
  │                      movies_catalogue, avg_rating_score  │
  │  Sheet: Users      → male_pct, female_pct, age_groups,  │
  │                      occupations, median_ratings, ...    │
  │  Sheet: Marketing  → genres_by_vol, genres_by_score,    │
  │                      peak_month, funnel_stages, ...      │
  │  Sheet: Top Movies → top_rated[10], most_reviewed[10]   │
  │  Sheet: Engagement → dau_peak, mau_avg, retention,      │
  │                      cohorts, score_dist, nps, ...       │
  └───────────────────────────┬─────────────────────────────┘
                              │  flat metrics dict (~40 keys)
              ┌───────────────┴───────────────┐
              ▼                               ▼
        fill_word()                     fill_pptx()
              │                               │
              │  ordered (pattern, value)     │  slide-scoped
              │  list substitution            │  shape-context substitution
              │  per paragraph + per cell     │  sequential placeholder fill
              ▼                               ▼
     DOCX output (4 tables,          PPTX output (10 slides,
     5 sections, ~40 values)         ~80 values across 8 content slides)
```

### 5.3 File Structure

```
watchit-corpus-interaction/              ← GitHub repository root
│
├── README.md                            ← This file — full project documentation
├── .gitignore
│
├── pipeline/
│   └── watchit_monthly_pipeline.py      ← Main pipeline script (run this monthly)
│
├── templates/
│   ├── WatchIt_Monthly_Management_Update_Template.docx   ← Word template  ⚠ do not delete
│   └── WatchIt_Monthly_Management_Update_Template.pptx   ← PPTX template  ⚠ do not delete
│
├── data/
│   ├── 202611_WatchIt_Metrics_Report.xlsx   ← Example input: November 2026 Excel
│   └── WatchIt_Metrics_Report.md            ← Metrics data dictionary
│
├── outputs/
│   ├── 202611_WatchIt_Monthly_Management_Update.docx  ← Example output: Nov 2026 Word
│   └── 202611_WatchIt_Management_Metrics.pptx         ← Example output: Nov 2026 PPTX
│
├── generators/
│   ├── generate_watchit_excel.py        ← Generates the Excel template from scratch
│   ├── generate_watchit_docx.py         ← Generates the Word template from scratch
│   └── generate_watchit_pptx.py         ← Generates the PPTX template from scratch
│
└── .kiro/
    └── skills/
        └── watchit-monthly-report/
            └── SKILL.md                 ← Kiro CLI skill: /watchit-monthly-report
```

### 5.4 Python Module Structure

```
pipeline/watchit_monthly_pipeline.py
│
├── SECTION 1: Template paths (resolved relative to script location)
│
├── SECTION 2: load_metrics(excel_path) → dict
│   ├── Validates sheet names
│   ├── Parses Cover sheet for headline KPIs
│   ├── Parses Users sheet (section-label navigation)
│   ├── Parses Marketing sheet (genres, funnel, monthly signups)
│   ├── Parses Top Movies sheet (top rated + most reviewed)
│   ├── Parses Engagement sheet (KPIs, cohorts, score distribution)
│   └── Returns normalised flat dict with ~40 string-formatted values
│
├── SECTION 3: fill_word(template, output, metrics, timestamp)
│   ├── _apply_subs(text, [(pattern, value)]) — ordered list substitution
│   ├── _replace_para_text(para, subs) — run-reconstruction substitution
│   ├── Body paragraph pass (5 narrative sections)
│   └── Table pass (4 tables, per-column context routing)
│
├── SECTION 4: fill_pptx(template, output, metrics, timestamp)
│   ├── Slide 1: title date
│   ├── Slide 3: KPI tiles (sequential positional fill)
│   ├── Slide 4: demographics (section-header detection)
│   ├── Slide 5: content/genre (table + score bars)
│   ├── Slide 6: marketing (funnel + KPI table)
│   ├── Slide 7: engagement (top KPIs + cohorts + retention)
│   ├── Slide 8: top movies (two 10-row tables)
│   ├── Slide 9: takeaways
│   └── Slide 10: thank you date
│
├── SECTION 5: Helpers
│   ├── extract_timestamp(excel_path) → "202611"
│   └── _month_label("202611") → "November 2026"
│
└── SECTION 6: main() — CLI entry point with argparse
```

---

## 6. Step-by-Step: What Was Built

This section documents the exact sequence of decisions and work that produced the pipeline.

### Step 1 — Understand the Documents

Before writing a single line of automation code, the three documents were inspected in detail:

- The Excel was loaded with `openpyxl` and all 5 sheets printed to understand the data layout, section label names, and value types.
- The Word template was opened with `python-docx` and all paragraph texts and table cell texts were printed to catalogue every `[ placeholder ]` string.
- The PPTX template was opened with `python-pptx` and all shape texts were printed per slide to catalogue every placeholder token.

This produced a complete map of **what needs to go where**.

### Step 2 — Design the Metrics Dictionary

A single flat Python dictionary was chosen as the bridge between the Excel reader and the two document fillers. This means:

- The Excel reader runs once and produces one dict.
- Both fillers consume the same dict independently.
- Adding a new metric requires changes only to `load_metrics()` and the specific filler.

The dict keys were named to be self-documenting: `male_pct`, `top_genre_vol`, `retention_m1m2`, etc.

### Step 3 — Build the Excel Reader

`load_metrics()` navigates the Excel by **section label lookup** rather than hardcoded cell addresses. For each sheet, a `find_value(rows, label)` helper searches for a row whose first column matches the label and returns the second column. This makes the reader resilient to rows being added above or below a section without breaking the lookup.

For structured sections (gender, age groups, occupations, cohorts), the reader uses **section-boundary detection** — it enters a section when it sees a header label (e.g. `"GENDER DISTRIBUTION"`) and exits when it sees the next section header.

### Step 4 — Build the Word Filler

The key insight was that Python `dict` silently drops duplicate keys. Instead, **ordered `(pattern, value)` lists** were used. The substitution function `_apply_subs()` uses `str.replace(pattern, value, 1)` — replacing only the **first occurrence** — which correctly handles paragraphs that contain the same pattern twice for different values (e.g. `male ([ __ ]%)` and `female ([ __ ]%)`).

The Word document required two passes:
1. **Body paragraphs**: each paragraph's full text is reconstructed from all its runs, substituted, and written back to `run[0]`.
2. **Table cells**: each cell is processed with a **context-specific substitution list** built from the row label and column index, so the same placeholder in a different cell gets the right value.

### Step 5 — Build the PPTX Filler

PowerPoint does not have a global text layer — every text string lives inside a shape, inside a text frame, inside paragraphs and runs. The filler processes each slide independently.

For slides with sequential KPI tiles (Slide 3, 7), a **positional counter** iterates over all shapes in document order and fills `______` tokens in sequence.

For slides with sections (Slide 4 — Users), a **section-detection pass** checks each shape's full text for a section header (e.g. `"Age Group Breakdown"`) and sets a flag that routes subsequent `__ %` tokens to the right substitution list.

For slides with tables (Slide 8 — Top Movies), the filler detects the table header text, then fills `__.___` and `_,___` tokens in row-order, alternating between avg-score and count columns.

### Step 6 — Wire the CLI

`argparse` was used to build a clean CLI interface with:
- Positional argument: path to the Excel file
- Optional `--output-dir` for saving reports elsewhere
- Optional `--word-template` and `--pptx-template` for custom template paths
- A guard that prevents accidental overwriting of the template files themselves

The timestamp is extracted from the filename prefix (`202611_...` → `"202611"`) and used to name both output files and generate the `"November 2026"` label for document headers.

### Step 7 — Create the Kiro Skill

A Kiro CLI skill was created at `.kiro/skills/watchit-monthly-report/SKILL.md`. This wraps the pipeline in a conversational interface:

```
/watchit-monthly-report 202612_WatchIt_Metrics_Report.xlsx
```

The skill instructs the AI agent to:
1. Identify the Excel file (from argument or by scanning the folder)
2. Run the pipeline script
3. Confirm the output files
4. Remind the user of the remaining editorial placeholders

---

## 7. Next Steps & Recommendations

### 7.1 From This Basis to a Full Agent

The current state is a **script + skill**. The next level of automation is a **dedicated Kiro agent** that owns the full reporting workflow. Here is the recommended path:

#### Step A — Create a Dedicated Agent Config

Create `.kiro/agents/watchit-reporting.json`:

```json
{
  "name": "watchit-reporting",
  "description": "WatchIt monthly reporting agent — runs the pipeline, summarises outputs, flags issues",
  "prompt": "You are the WatchIt reporting assistant. Your job is to run the monthly report pipeline when given an Excel file, verify the outputs are correct, summarise what was filled, flag any remaining editorial placeholders, and help the user prepare the final documents for distribution. You know the WatchIt platform metrics schema and can explain any figure in the reports.",
  "tools": ["read", "write", "shell", "glob", "grep"],
  "allowedTools": ["read", "glob", "grep"],
  "toolsSettings": {
    "shell": {
      "allowedCommands": [
        "python3 /path/to/watchit-corpus-interaction/pipeline/watchit_monthly_pipeline.py"
      ],
      "autoAllowReadonly": true
    },
    "write": {
      "allowedPaths": ["/path/to/watchit-corpus-interaction/**"]
    }
  },
  "resources": [
    "skill://.kiro/skills/watchit-monthly-report/SKILL.md",
    "file:///path/to/watchit-corpus-interaction/README.md"
  ],
  "keyboardShortcut": "ctrl+shift+w",
  "welcomeMessage": "WatchIt Reporting Agent ready. Drop a YYYYMM_WatchIt_Metrics_Report.xlsx filename to generate this month's reports."
}
```

With this agent active, you switch to it with `ctrl+shift+w` and the full pipeline context is loaded.

#### Step B — Add a Pre-Run Validation Hook

Add a hook that validates the Excel before running:

```json
"hooks": {
  "agentSpawn": [
    {
      "command": "python3 -c \"import openpyxl; print('openpyxl ready')\" 2>/dev/null || echo 'WARNING: openpyxl not installed — run: pip install openpyxl python-docx python-pptx'"
    }
  ]
}
```

#### Step C — Automate Further with LLM-Generated Narrative

The four `[ Highlight 1–4 ]` key points in the Word executive summary and the `[ positive / stable / declining ]` tone judgement are the only items still requiring human input. These can be automated by:

1. After the pipeline runs, pass the metrics dict to the LLM with a prompt: *"Given these platform metrics for November 2026, write 4 concise management-level key points and determine whether retention momentum is positive, stable, or declining."*
2. Use the agent's `write` tool to inject these into the already-generated DOCX.

This would make the reports **100% automated** — only the preparer name would remain manual.

#### Step D — Schedule Monthly Execution

For full automation, the pipeline can be scheduled:

```bash
# crontab entry — runs on the 1st of every month at 08:00
0 8 1 * * python3 /path/to/watchit-corpus-interaction/pipeline/watchit_monthly_pipeline.py \
    $(ls /path/to/watchit-corpus-interaction/data/*_WatchIt_Metrics_Report.xlsx | sort | tail -1) \
    --output-dir /path/to/watchit-corpus-interaction/outputs/$(date +%Y%m)/
```

### 7.2 Recommended Optimisations

| Optimisation | Effort | Impact |
|---|---|---|
| LLM-generated key points (Highlights 1–4) | Low | High — completes the automation |
| Email delivery on completion | Low | High — removes the distribution step |
| Validation report: flag metrics that changed >20% MoM | Medium | High — catches data quality issues |
| PDF export of final DOCX and PPTX | Low | Medium — ready-to-send format |
| Multi-platform support (Google Docs / Slides) | High | Medium — broader adoption |
| Dashboard integration (Power BI / Tableau embed) | High | High — replaces static decks entirely |

---

## 8. Generalising This Pattern — Corpus Interaction as a Framework

This project is an instance of a general pattern that applies whenever:

> *You have a structured data source that refreshes on a schedule, and one or more document templates that need to be populated from it.*

### The General Architecture

```
┌──────────────────────────────────────────────────────────────────────┐
│                   CORPUS INTERACTION FRAMEWORK                       │
│                                                                      │
│   DATA SOURCE          BRIDGE             DOCUMENT TARGETS           │
│   (structured,    →   (Python      →     (DOCX, PPTX, PDF,          │
│    refreshing)         script)            HTML, Markdown, ...)       │
│                                                                      │
│   Excel / CSV          load_data()        fill_word()                │
│   SQL database         normalise()        fill_pptx()                │
│   API / JSON           format()           fill_pdf()                 │
│   Google Sheets                           fill_html()                │
└──────────────────────────────────────────────────────────────────────┘
```

### How to Handle a New Similar Use Case

Follow this 7-step process for any new Corpus Interaction project:

#### 1. Map the Data Source
- List every piece of data that will appear in the output documents.
- Identify where each lives in the source (sheet name, column header, API field).
- Check whether the schema is stable month-to-month. **Automation requires a stable schema.**

#### 2. Map the Document Targets
- Open each template and catalogue every placeholder location.
- Print all paragraph texts and table cell texts programmatically (`python-docx`, `python-pptx`).
- Note where the same pattern appears multiple times with different intended values — these are the tricky substitution points.

#### 3. Design the Bridge Dict
- Create a flat, self-documenting Python dict that holds all values in display-ready format.
- Normalise all types to strings. Format numbers as they should appear in the output.
- Keep the dict flat — no nested dicts. Every key is a direct lookup.

#### 4. Build the Reader
- Use label-based lookup, not hardcoded cell addresses.
- Validate required fields at the top — fail fast with clear error messages.
- Test the dict output before touching the templates.

#### 5. Build the Fillers (one per document type)
- Use ordered `(pattern, value)` lists, not dicts.
- Use `str.replace(pattern, value, 1)` to replace only the first occurrence.
- Process document-type-specific structures (runs in DOCX, shapes in PPTX) with context-aware routing.
- Test each filler independently against the template.

#### 6. Wire the CLI
- Accept the data source path as a positional argument.
- Extract the timestamp/period from the filename if possible.
- Guard against template overwrites.
- Add `--output-dir` for flexible deployment.

#### 7. Create the Skill
- Write a `SKILL.md` in `.kiro/skills/<use-case-name>/`.
- Include: what the skill does, how to identify the input, how to invoke the script, what to check in the output, what requires human editorial input.
- Optionally create a dedicated agent config for high-frequency use cases.

### Checklist for a New Corpus Interaction Use Case

```
□  Data source has a stable schema (consistent field names, sheet names)
□  Data source is timestamped or versioned (filename prefix, date field)
□  Templates contain clear, unique placeholder patterns
□  Bridge dict maps every placeholder to a data source location
□  Reader validated against at least 2 months of historical data
□  Filler tested: all placeholders replaced, no templates overwritten
□  CLI documented: --help is accurate and complete
□  Skill created: /use-case-name works in Kiro CLI
□  README written: new user can run the pipeline in under 5 minutes
□  Committed to version control
```

---

## Appendix: Technology Reference

| Library | Version | Role |
|---------|---------|------|
| `openpyxl` | ≥ 3.1 | Read Excel workbooks (`.xlsx`) |
| `python-docx` | ≥ 1.1 | Read and write Word documents (`.docx`) |
| `python-pptx` | ≥ 0.6 | Read and write PowerPoint files (`.pptx`) |
| Python | ≥ 3.8 | Runtime |
| Kiro CLI | current | AI agent runtime, skill/agent host |

Install all dependencies:

```bash
pip install openpyxl python-docx python-pptx
```

---

## Appendix: Running the Pipeline

```bash
# Standard monthly run
python3 pipeline/watchit_monthly_pipeline.py 202612_WatchIt_Metrics_Report.xlsx

# Save outputs to a reports subfolder
python3 pipeline/watchit_monthly_pipeline.py 202612_WatchIt_Metrics_Report.xlsx \
    --output-dir reports/2026/

# Via Kiro skill
/watchit-monthly-report 202612_WatchIt_Metrics_Report.xlsx

# Get help
python3 pipeline/watchit_monthly_pipeline.py --help
```

Expected output:

```
🎬  WatchIt Monthly Report Pipeline
    Period   : December 2026  (202612)
    Excel    : 202612_WatchIt_Metrics_Report.xlsx
    Output   : /home/Kisa/CorpusInteraction

📊  Reading Excel metrics …
📝  Filling Word template …
  ✅  Word saved  →  202612_templates/WatchIt_Monthly_Management_Update_Template.docx
📊  Filling PowerPoint template …
  ✅  PPTX saved  →  202612_WatchIt_Management_Metrics.pptx

✅  Done!  Both reports for December 2026 are ready.
```

---

*This documentation was produced as part of the WatchIt CorpusInteraction project. For questions, contact the project author.*
