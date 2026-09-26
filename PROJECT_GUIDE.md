# CO–PO Attainment Website Project

## Dynamic template outcomes

The website reads the uploaded official Excel template before displaying attainment settings. CO, PO and PSO labels are detected from the workbook, and the target and mapping tables are rebuilt automatically. The interface is therefore not limited to PO1–PO6: a UG template containing PO1–PO12 and PSO labels will display those outcomes, while an M.Tech template containing PO1–PO6 will display six POs.

Developed by **Omkar S. Powar**  
Assistant Professor, School of Electrical Engineering, MIT, Manipal

## Purpose

This local/public web application accepts the approved academic files, validates the data, calculates CO–PO attainment, and produces a completed official Excel workbook.

## Live website

https://co-po-attainment-ezoc.onrender.com

## Files accepted

1. Official M.Tech CO–PO Excel template (required)
2. CO-wise assessment-marks Excel file (required)
3. Grade and CGPA Excel file (optional)
4. Course-end survey Excel file (optional)
5. Course-plan Word file (optional)

## Current calculations

- Registration-number and marks validation
- CIE attainment from IA and mid-semester marks
- SEE attainment from final-examination marks
- Direct attainment using editable CIE/SEE weights
- Indirect attainment from course-end survey ratings
- Overall CO attainment using editable direct/indirect weights
- PO attainment through the editable CO–PO mapping
- Individual previous target, previous attainment and current target for each CO and PO
- Separate target-achievement status for every CO and mapped PO
- Official Excel workbook generation and one-time download

## Target-setting recommendation

- If previous attainment is below the previous target, retain the previous target.
- If previous attainment meets or exceeds the previous target, suggest a moderately higher rounded target.
- Every suggested target remains editable because target approval involves faculty/programme judgment.

## Main project files

- `server.py` — upload/download web server and temporary-file handling
- `engine.py` — validation, calculations and Excel-template filling
- `xlsx_raw.py` — lightweight Excel reading support
- `static/index.html` — website interface
- `static/app.js` — dynamic mapping and target tables
- `static/styles.css` — website styling
- `render.yaml`, `Procfile`, `requirements.txt` — Render deployment configuration

## Privacy

Uploaded files are processed temporarily and removed immediately after generation. Generated downloads are one-time links and expire after 15 minutes.

## Deploying an update

Upload the changed files to the `omkarspowar/co-po-attainment` GitHub repository using the same paths and filenames. Commit directly to `main`. Render will redeploy automatically, normally within a few minutes.

## Run locally

Install Python 3.11 or newer, open a terminal in the project folder, and run:

```bash
pip install -r requirements.txt
python server.py
```

Then open `http://127.0.0.1:8765`.

## Current supported structure

The present calculation engine is configured for the supplied M.Tech theory template with four COs and six PO columns. The interface and engine should be extended together before using a template with a different number or arrangement of CO/PO/PSO columns.
