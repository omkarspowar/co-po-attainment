# CO–PO Attainment Automation

CO, PO and PSO columns are detected dynamically from the uploaded official template; they are not fixed in the webpage.

## General assessment workflow

The hosted application accepts a QP Analysis document, one combined IA workbook or multiple IA files, question-wise Midsem marks, question/subquestion-wise final marks, final grades, course-end survey and course plan. Student rows are matched and sorted by registration number, so source files may be in different orders. The number of internal assessments is read from the QP Analysis rather than fixed at four.

If a marks column combines subparts mapped to different COs, the application displays a mandatory student-wise split table. Generation remains blocked until every split equals the original combined mark.

CO and PO/PSO action plans are based on each outcome's actual attainment, target, gap, CO statement and Bloom's level. Outcomes that attain the target receive a continuation statement; only below-target outcomes receive corrective actions and RCA.

RCA and action plans are written automatically by the application. Faculty do not need to type them. The RCA identifies the weakest available evidence source (CIE, SEE or course-end survey), while the action plan selects a suitable academic intervention and a measurable follow-up assessment.

A private local web application for completing the M.Tech theory CO–PO attainment template.

## Start on Windows

1. Install Python 3.10 or newer from https://www.python.org/downloads/ and select **Add Python to PATH**.
2. Extract this application folder.
3. Double-click `start.bat`.
4. The browser opens at `http://127.0.0.1:8765`.

## Required uploads

- Official unfilled M.Tech theory CO–PO template (`.xlsx`)
- CO-wise marks analysis (`.xlsx`) with a `CO Summary` sheet

## Optional uploads

- Grade and CGPA export (`.xlsx`)
- Course-end survey export (`.xlsx`)
- Course plan (`.docx`)

## Expected marks format

The `CO Summary` sheet must include:

- Registration No. and Student Name
- IA CO1–CO4
- Mid-semester CO1–CO4
- Final-examination CO1–CO4
- A `Maximum Marks` row

The current `AIMI_CO_Wise_Marks_Analysis` format is supported directly.

## Calculations

- Level 3: score ≥ 75%
- Level 2: score ≥ 50% and < 75%
- Level 1: score < 50%
- Default direct attainment: 60% CIE + 40% SEE
- Default overall attainment: 80% direct + 20% indirect
- Survey: rating 5 → Level 3; rating 3 or 4 → Level 2; rating 1 or 2 → Level 1

All weights, the target and CO–PO mapping can be changed in the web interface.

## Privacy

Temporary uploads are deleted immediately after generation. Completed workbooks use a one-time download link that expires after 15 minutes; the server deletes the generated file after download or expiry.

## Deploy on Render

1. Push this folder to a GitHub repository.
2. In Render, choose **New → Blueprint** and connect the repository.
3. Render reads `render.yaml` and creates the web service.
4. Open the generated `onrender.com` URL.

No password is enabled. Anyone with the public URL can use the website.

## Stop the application

Return to the command window and press `Ctrl+C`.
