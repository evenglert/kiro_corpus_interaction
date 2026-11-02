---
name: watchit-monthly-report
description: Run the WatchIt monthly report pipeline to generate timestamped Word and PPTX reports from a monthly Excel metrics file.
---

# WatchIt Monthly Report Pipeline

Run the WatchIt monthly report pipeline using the script at `/home/Kisa/CorpusInteraction/watchit_monthly_pipeline.py`.

## Task

$ARGUMENTS

## What to do

1. Identify the Excel file to process. The filename must start with a 6-digit YYYYMM timestamp (e.g. `202611_WatchIt_Metrics_Report.xlsx`). If the user has not specified a file, look for the most recent `*_WatchIt_Metrics_Report.xlsx` in `/home/Kisa/CorpusInteraction/`.

2. Run the pipeline:
   ```
   python3 /home/Kisa/CorpusInteraction/watchit_monthly_pipeline.py <excel_file>
   ```
   If the user wants outputs in a specific folder, add `--output-dir <folder>`.

3. Confirm which two files were created:
   - `<YYYYMM>_WatchIt_Monthly_Management_Update.docx`
   - `<YYYYMM>_WatchIt_Management_Metrics.pptx`

4. Remind the user of the editorial placeholders still requiring manual input:
   - `[ Name ]` — preparer name in the Word header
   - `[ Highlight 1–4 ]` — key points bullets in the executive summary
   - `[ young-adult / professional ]` and `[ positive / stable / declining ]` — tone choices in the body
   - **Next Steps** box on PPTX slide 9

## Notes

- Templates (`WatchIt_Monthly_Management_Update.docx` and `WatchIt_Management_Metrics_Template.pptx`) are never overwritten.
- The timestamp is extracted automatically from the Excel filename prefix.
- Required Python packages: `openpyxl`, `python-docx`, `python-pptx`. If missing, install with `pip install openpyxl python-docx python-pptx`.
