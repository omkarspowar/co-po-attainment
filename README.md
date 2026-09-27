# CO–PO Attainment Automation

The interface is branded as **OSP Academic Quality Automation**. Every upload card shows a red status light before file selection and a green status light after selection, with a responsive four-step workflow for upload, review, configuration and generation.

Version 6.0 presents a product-ready faculty experience with a guided landing area, standard/prepared upload routes, clear required and optional file labels, automatic workflow navigation, review checkpoints, responsive sticky-header tables, generation progress and a structured success/download screen.

Interface refreshes are idempotent: uploading the template, course plan or QP repeatedly rebuilds each target/mapping table once and preserves compatible values already entered by the faculty.

When the course plan contains a CO–PO/PSO articulation matrix, supported mapping values are loaded automatically into the interface. Missing CO rows remain blank for faculty review, and extra outcomes that do not exist in the uploaded official template are reported and ignored.

CO, PO and PSO columns are detected dynamically from the uploaded official template; they are not fixed in the webpage.

Outcome detection is restricted to the official target table so labels appearing elsewhere in formulas or unused sheets are not mistakenly added. CO slots without target data remain inactive until confirmed by the QP Analysis or course plan. Unmapped CO–PO/PSO cells are displayed as blank; blank is interpreted internally as mapping value zero.

## General assessment workflow

The hosted application accepts a QP Analysis document, one combined IA workbook or multiple IA files, question-wise Midsem marks, question/subquestion-wise final marks, final grades, course-end survey and course plan. Student rows are matched and sorted by registration number, so source files may be in different orders. The number of internal assessments is read from the QP Analysis rather than fixed at four. Registration columns are recognized from common variants (`Registration Number`, `Reg No.`, `Reg. No`, `Enrollment ID`, `USN`, `Roll No`, `Student ID`) or inferred from student-number values when the heading is unconventional.

If a marks column combines subparts mapped to different COs, the application displays a mandatory student-wise split table. Generation remains blocked until every split equals the original combined mark.

When an examination export contains its own CO-wise totals, the application compares that mapping with the QP Analysis. If the maxima differ, the Assessment Review section shows both versions and requires the faculty to select the approved source before generation.

For multi-sheet IA exports, the application selects the complete table matching the detected number of internal assessments. Missing raw marks remain pending unless a complete consolidated table is present.

CO and PO/PSO action plans are based on each outcome's actual attainment, target, gap, CO statement and Bloom's level. Corrective-action cells are populated only for outcomes below target. Attained and unmapped outcomes remain blank in the action-plan column, while the RCA records the complete outcome analysis.

RCA and action plans are written automatically by the application. Faculty do not need to type them. The RCA identifies the weakest available evidence source (CIE, SEE or course-end survey), while the action plan selects a suitable academic intervention and a measurable follow-up assessment.

A private local web application for completing the M.Tech theory CO–PO attainment template.

## Start on Windows

1. Install Python 3.10 or newer from https://www.python.org/downloads/ and select **Add Python to PATH**.
2. Extract this application folder.
3. Double-click `start.bat`.
4. The browser opens at `http://127.0.0.1:8765`.

## Required uploads

- Official unfilled UG theory CO–PO template (`.xlsx`)
- Either source assessment files or a prepared CO-wise marks workbook

## Source assessment uploads

- QP Analysis (`.docx`)
- One combined IA workbook or any number of separate IA workbooks
- Mid-semester question-wise marks
- End-term question/subquestion-wise marks
- Final-grade export (`.xlsx`)
- Course-end survey export (`.xlsx`)
- Course plan (`.docx`)

## Expected marks format

The `CO Summary` sheet must include:

- Registration No. and Student Name
- IA CO1–CO4
- Mid-semester CO1–CO4
- Final-examination CO1–CO4
- A `Maximum Marks` row

Students are matched and sorted by registration number. A student without a confirmed final grade remains pending and is excluded from attainment rather than being treated as zero.

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
