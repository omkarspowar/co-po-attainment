from __future__ import annotations

import json
import math
import re
import shutil
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

from xlsx_raw import first_nonempty_sheet, read_xlsx
from assessment_import import normalize_assessments


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def number(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def canonical_reg(value: Any) -> str:
    raw = text(value)
    return raw[:-2] if raw.endswith(".0") else raw


@dataclass
class Student:
    reg: str
    name: str
    ia: list[float]
    mid: list[float]
    see: list[float]
    grade: str = ""


@dataclass
class Course:
    code: str = ""
    name: str = ""
    faculty: str = ""
    school: str = "Department of Biomedical Engineering"
    program: str = "M.Tech Biomedical Engineering (Medical Informatics)"
    semester: str = "II"
    year: str = ""
    odd_even: str = "Even"
    section: str = "Medical Informatics"
    course_type: str = "Core"
    cos: list[str] = field(default_factory=lambda: ["", "", "", ""])
    bloom: list[str] = field(default_factory=lambda: ["L3", "L3", "L5", "L6"])


DEFAULT_MAPPING = [[2, 1, 3, 0, 0, 0], [3, 1, 3, 0, 0, 0], [3, 2, 3, 0, 0, 0], [3, 3, 3, 0, 0, 0]]

OUTCOME_RE = re.compile(r"^(CO|PO|PSO)\s*[-_ ]?\s*(\d+)$", re.I)


def inspect_template(path: Path) -> dict[str, list[str]]:
    """Return active COs and defined PO/PSO labels from the official TARGET sheet."""
    wb = load_workbook(path, data_only=False, read_only=True)
    target_ws=next((ws for ws in wb.worksheets if "TARGET" in ws.title.upper()),None)
    if target_ws:
        co_labels=[];active_co_labels=[];outcome_labels=[]
        for row in target_ws.iter_rows():
            first=text(row[0].value).upper().replace(" ","") if row else ""
            if first=="CO":
                candidates=[]
                for cell in row[1:]:
                    match=OUTCOME_RE.fullmatch(text(cell.value))
                    if match and match.group(1).upper()=="CO": candidates.append((cell.column,f"CO{int(match.group(2))}"))
                active=[]
                for col,label in candidates:
                    if any(text(target_ws.cell(row[0].row+offset,col).value) for offset in (1,2,3)): active.append(label)
                co_labels=[label for _,label in candidates];active_co_labels=active
            if first in {"PO&PSO","PO/PSO","POPSO"}:
                for cell in row[1:]:
                    match=OUTCOME_RE.fullmatch(text(cell.value))
                    if match and match.group(1).upper() in {"PO","PSO"}: outcome_labels.append(f"{match.group(1).upper()}{int(match.group(2))}")
        if co_labels and outcome_labels: return {"cos":co_labels,"active_cos":active_co_labels,"outcomes":outcome_labels}
    # Fallback for older templates without a recognizable target sheet.
    found: dict[str, set[int]] = {"CO": set(), "PO": set(), "PSO": set()}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                match = OUTCOME_RE.fullmatch(text(cell.value))
                if match:
                    found[match.group(1).upper()].add(int(match.group(2)))
    labels = {"cos": [f"CO{i}" for i in sorted(found["CO"])], "outcomes": [f"PO{i}" for i in sorted(found["PO"])] + [f"PSO{i}" for i in sorted(found["PSO"])]}
    if not labels["cos"] or not labels["outcomes"]:
        raise ValueError("The uploaded template must contain CO and PO/PSO labels.")
    labels["active_cos"]=labels["cos"]
    return labels


def parse_marks(path: Path) -> tuple[list[Student], list[float], list[float], list[float]]:
    sheets = read_xlsx(path)
    rows = sheets.get("CO Summary") or next(iter(sheets.values()), [])
    header_index = next((i for i, row in enumerate(rows) if len(row) > 1 and text(row[0]).lower() == "registration no."), -1)
    if header_index < 0 or header_index + 1 >= len(rows):
        raise ValueError("Marks file must contain a 'CO Summary' table with Registration No. and CO1–CO4 columns.")
    maximum = rows[header_index + 1]
    ia_max = [number(x) for x in maximum[2:6]]
    mid_max = [number(x) for x in maximum[7:11]]
    see_max = [number(x) for x in maximum[12:16]]
    students: list[Student] = []
    for row in rows[header_index + 2:]:
        reg = canonical_reg(row[0] if row else None)
        if not reg or not re.fullmatch(r"\d{6,}", reg):
            continue
        padded = list(row) + [None] * (17 - len(row))
        students.append(Student(reg, text(padded[1]), [number(x) for x in padded[2:6]], [number(x) for x in padded[7:11]], [number(x) for x in padded[12:16]]))
    if not students:
        raise ValueError("No student rows were found in the marks file.")
    return students, ia_max, mid_max, see_max


def apply_grades(path: Path | None, students: list[Student], course_code: str) -> list[str]:
    warnings: list[str] = []
    if not path:
        warnings.append("Grade file was not supplied; grade cells will remain blank.")
        return warnings
    selected=None
    for sheet_name,rows in read_xlsx(path).items():
        for header_i,row in enumerate(rows[:20]):
            header=[text(x).lower() for x in row]
            reg_i=next((i for i,x in enumerate(header) if any(k in re.sub(r"[^a-z]","",x) for k in ("regnumber","registrationnumber","enrollmentid","orgdefinedid","rollnum"))),None)
            grade_i=next((i for i,x in enumerate(header) if x.strip() in {"grade","see grade","final grade","letter grade"}),None)
            if reg_i is not None and grade_i is not None:
                selected=(sheet_name,rows,header_i,header,reg_i,grade_i);break
        if selected: break
    if not selected:
        return ["Could not find a worksheet containing both registration numbers and a final-grade column."]
    sheet_name,rows,header_i,header,reg_i,grade_i=selected
    name_i=next((i for i,x in enumerate(header) if x in {"name","student name","participant account: name","first name"}),None)
    code_i=next((i for i,x in enumerate(header) if "course code" in x),None)
    by_reg: dict[str, tuple[str, str]] = {}
    duplicates=[]
    required=[reg_i,grade_i]+([name_i] if name_i is not None else [])+([code_i] if code_i is not None else [])
    for row in rows[header_i+1:]:
        padded = list(row) + [None] * (max(required) + 1 - len(row))
        if code_i is not None and course_code and text(padded[code_i]).replace(" ", "").lower() != course_code.replace(" ", "").lower():
            continue
        reg = canonical_reg(padded[reg_i])
        if reg:
            if reg in by_reg: duplicates.append(reg)
            by_reg[reg] = (text(padded[name_i]) if name_i is not None else "", text(padded[grade_i]))
    if duplicates: warnings.append(f"Duplicate grade rows found for: {', '.join(sorted(set(duplicates)))}.")
    blank=[reg for reg,(name,grade) in by_reg.items() if not grade]
    if blank: warnings.append(f"Blank final grades found for: {', '.join(blank)}.")
    for student in students:
        if student.reg in by_reg:
            if by_reg[student.reg][0]: student.name = by_reg[student.reg][0].title()
            student.grade = by_reg[student.reg][1]
        else:
            warnings.append(f"No grade matched registration number {student.reg}.")
    return warnings


def parse_course_plan(path: Path | None, fallback: Course) -> Course:
    if not path:
        return fallback
    doc = Document(path)
    content = "\n".join([p.text for p in doc.paragraphs] + [" | ".join(text(c.text) for c in row.cells) for table in doc.tables for row in table.rows])
    patterns = {
        "code": r"Course\s*Code\s*[:|]\s*([A-Z]{2,}\s*\d{3,})",
        "name": r"Course\s*(?:Name|Title)\s*[:|]\s*([^\n|]+)",
        "program": r"(?:Programme|Program)\s*(?:Name)?\s*[:|]\s*([^\n|]+)",
        "semester": r"Semester\s*[:|]\s*([^\n|]+)",
        "year": r"Academic\s*Year\s*[:|]\s*([^\n|]+)",
        "school": r"(?:School|Department)\s*[:|]\s*([^\n|]+)",
        "course_type": r"Course\s*Type\s*[:|]\s*([^\n|]+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, content, flags=re.I)
        if match and not getattr(fallback, key):
            setattr(fallback, key, match.group(1).strip())
    if not fallback.name:
        first_heading = next((p.text.strip() for p in doc.paragraphs if p.text.strip()), "")
        if first_heading and len(first_heading) < 150: fallback.name = first_heading
    co_matches = re.findall(r"\bCO\s*([1-6])\s*[:.\-|]\s*([^\n|]+)", content, flags=re.I)
    for index, statement in co_matches:
        position = int(index) - 1
        if position < len(fallback.cos) and len(statement.strip()) > 15:
            fallback.cos[position] = statement.strip()
    for co in range(1, len(fallback.cos)+1):
        bloom_match = re.search(rf"\bCO\s*{co}\b[^\n|]*?\bL\s*([1-6])\b", content, flags=re.I)
        if bloom_match: fallback.bloom[co-1] = f"L{bloom_match.group(1)}"
    # Institutional table format: CO/CLO | Statements | ... | BL.
    for table in doc.tables:
        rows=[[text(c.text) for c in row.cells] for row in table.rows]
        if not rows: continue
        headers=[re.sub(r"\s+"," ",x).strip().lower() for x in rows[0]]
        co_i=next((i for i,x in enumerate(headers) if "co/clo" in x or x in {"co","clo"}),None)
        statement_i=next((i for i,x in enumerate(headers) if "statement" in x),None)
        bloom_i=next((i for i,x in enumerate(headers) if x.startswith("bl") or "bloom" in x),None)
        if co_i is not None and statement_i is not None:
            for row in rows[1:]:
                if max(co_i,statement_i)>=len(row): continue
                match=re.search(r"(?:CO\s*)?(\d+)",row[co_i],re.I)
                if not match: continue
                position=int(match.group(1))-1
                if position<len(fallback.cos) and row[statement_i]: fallback.cos[position]=row[statement_i]
                if bloom_i is not None and bloom_i<len(row) and row[bloom_i]: fallback.bloom[position]=re.sub(r"\s+"," ",row[bloom_i].replace("/"," ")).strip()
        # Label/value metadata table.
        for row in rows:
            for i,value in enumerate(row[:-1]):
                key=re.sub(r"[^a-z]","",value.lower())
                candidate=row[i+1].strip()
                if not candidate: continue
                if key=="schoolname" and not fallback.school: fallback.school=candidate
                elif key in {"nameofthefaculty","facultyname"} and not fallback.faculty: fallback.faculty=candidate
                elif key in {"corepeoe","coursetype"} and not fallback.course_type: fallback.course_type=candidate
    return fallback


def inspect_course_plan(path: Path) -> dict[str, Any]:
    blank = Course(school="", program="", semester="", odd_even="", section="", course_type="", bloom=["", "", "", ""])
    course = parse_course_plan(path, blank)
    articulation: dict[str,dict[str,float]]={};mapping_outcomes=[];warnings=[]
    if path.suffix.lower()==".docx":
        doc=Document(path)
        for table in doc.tables:
            rows=[[text(c.text) for c in row.cells] for row in table.rows]
            if not rows: continue
            header_i=next((i for i,row in enumerate(rows) if sum(bool(OUTCOME_RE.fullmatch(x)) and not x.upper().startswith("CO") for x in row)>=2),None)
            if header_i is None: continue
            headers=[]
            for column,value in enumerate(rows[header_i]):
                match=OUTCOME_RE.fullmatch(value)
                if match and match.group(1).upper() in {"PO","PSO"}: headers.append((column,f"{match.group(1).upper()}{int(match.group(2))}"))
            for row in rows[header_i+1:]:
                co_match=next((re.fullmatch(r"\s*CO\s*(\d+)\s*",value,re.I) for value in row if re.fullmatch(r"\s*CO\s*(\d+)\s*",value,re.I)),None)
                if not co_match: continue
                label=f"CO{int(co_match.group(1))}";values={}
                for column,outcome in headers:
                    raw=row[column] if column<len(row) else ""
                    if raw!="" and 0<=number(raw,-1)<=3: values[outcome]=number(raw)
                articulation[label]=values
            mapping_outcomes=[label for _,label in headers]
            if articulation: break
    active=[f"CO{i+1}" for i,value in enumerate(course.cos) if value]
    missing=[co for co in active if co not in articulation]
    if missing: warnings.append("No articulation-matrix row was found for "+", ".join(missing)+"; those mappings remain blank for faculty review.")
    return {"course_code": course.code, "course_name": course.name, "faculty":course.faculty, "school": course.school, "program": course.program, "semester": course.semester, "year": course.year, "course_type": course.course_type, "cos": course.cos, "bloom": course.bloom,"mapping":articulation,"mapping_outcomes":mapping_outcomes,"warnings":warnings}


def parse_survey(path: Path | None, co_count: int=4) -> list[dict[str, Any]]:
    if not path:
        return []
    rows = first_nonempty_sheet(path)
    if not rows:
        return []
    header = [text(x).lower().replace("\xa0", " ") for x in rows[0]]
    name_i = next((i for i, x in enumerate(header) if "student's full name" in x or "student full name" in x), 6)
    reg_i = next((i for i, x in enumerate(header) if "reg. no" in x or "registration" in x), 7)
    co_indices = []
    for co in range(1, co_count+1):
        pattern = re.compile(rf"\bco\s*{co}\b")
        co_indices.append(next((i for i, x in enumerate(header) if pattern.search(x)), 7 + co))
    # Keep the latest response for each registration number so duplicate form
    # submissions never give one student extra weight in indirect attainment.
    responses_by_reg: dict[str, dict[str, Any]] = {}
    for row in rows[1:]:
        padded = list(row) + [None] * (max([name_i, reg_i] + co_indices) + 1 - len(row))
        reg = canonical_reg(padded[reg_i])
        if not reg:
            continue
        ratings = [number(padded[i]) for i in co_indices]
        if any(ratings):
            responses_by_reg[reg]={"name": text(padded[name_i]), "reg": reg, "ratings": ratings}
    return list(responses_by_reg.values())


def attainment(mark: float, maximum: float, grade: str) -> int:
    if maximum <= 0 or grade.upper() in {"DT", "I"}:
        return 0
    percentage = mark / maximum * 100
    return 3 if percentage >= 75 else 2 if percentage >= 50 else 1


def weighted_level(levels: list[int]) -> float:
    valid = [x for x in levels if x > 0]
    return sum(valid) / len(valid) if valid else 0.0


def compute(students: list[Student], ia_max: list[float], mid_max: list[float], see_max: list[float], responses: list[dict[str, Any]], direct_weights: tuple[float, float], overall_weights: tuple[float, float], mapping: list[list[float]], target: float) -> dict[str, Any]:
    n=len(ia_max)
    cie_max = [ia_max[i] + mid_max[i] for i in range(n)]
    cie_marks = [[s.ia[i] + s.mid[i] for i in range(n)] for s in students]
    cie_levels = [[attainment(cie_marks[r][c], cie_max[c], students[r].grade) for c in range(n)] for r in range(len(students))]
    see_levels = [[attainment(students[r].see[c], see_max[c], students[r].grade) for c in range(n)] for r in range(len(students))]
    cie = [weighted_level([row[c] for row in cie_levels]) for c in range(n)]
    see = [weighted_level([row[c] for row in see_levels]) for c in range(n)]
    direct = [(cie[i] * direct_weights[0] + see[i] * direct_weights[1]) / 100 for i in range(n)]
    indirect = []
    for c in range(n):
        levels = [3 if x["ratings"][c] == 5 else 2 if x["ratings"][c] >= 3 else 1 for x in responses]
        indirect.append(weighted_level(levels))
    overall = [((direct[i] * overall_weights[0] + indirect[i] * overall_weights[1]) / 100) if responses else direct[i] for i in range(n)]
    def po(co_values: list[float]) -> list[float]:
        result = []
        for p in range(len(mapping[0]) if mapping else 0):
            denominator = sum(mapping[c][p] for c in range(n))
            result.append(sum(co_values[c] * mapping[c][p] for c in range(n)) / denominator if denominator else 0.0)
        return result
    return {"cie_max": cie_max, "cie_marks": cie_marks, "cie_levels": cie_levels, "see_levels": see_levels, "cie": cie, "see": see, "direct": direct, "indirect": indirect, "overall": overall, "direct_po": po(direct), "overall_po": po(overall), "target": target}


def _sheet(workbook, title: str):
    normalized=re.sub(r"[^A-Z0-9]","",title.upper())
    exact=next((ws for ws in workbook.worksheets if re.sub(r"[^A-Z0-9]","",ws.title.upper())==normalized),None)
    if exact: return exact
    number=re.match(r"\s*(\d+)\s*\.",title)
    if number:
        prefix=number.group(1)+"."
        return next((ws for ws in workbook.worksheets if ws.title.strip().startswith(prefix)),None)
    return None


def _label_cells(ws, labels: list[str]) -> dict[str, list[Any]]:
    result = {x.upper(): [] for x in labels}
    for row in ws.iter_rows():
        for cell in row:
            key = text(cell.value).upper().replace(" ", "")
            if key in result: result[key].append(cell)
    return result


def _co_analysis(i: int, course: Course, result: dict[str, Any], target: float) -> tuple[str, str]:
    label=f"CO{i+1}"; statement=course.cos[i] or label; bloom=course.bloom[i] or "the mapped cognitive level"
    overall=result["overall"][i]; direct=result["direct"][i]; indirect=result["indirect"][i]
    components={"continuous internal assessment":result["cie"][i],"semester-end examination":result["see"][i]}
    if indirect>0: components["course-end survey"]=indirect
    weak_name,weak_value=min(components.items(),key=lambda x:x[1])
    if overall>=target:
        rca=f"{label} attained the target ({overall:.2f} against {target:.2f}). The comparatively lowest evidence was from the {weak_name} ({weak_value:.2f}); this is a monitoring point rather than a current attainment failure."
        action=""
    else:
        gap=target-overall
        rca=f"{label} did not attain the target: overall attainment {overall:.2f}, target {target:.2f}, gap {gap:.2f}. The weakest evidence was the {weak_name} ({weak_value:.2f}); direct attainment was {direct:.2f}"+(f" and indirect attainment was {indirect:.2f}." if indirect>0 else ".")+f" This indicates insufficient achievement of “{statement}”, particularly at {bloom}."
        if "survey" in weak_name: intervention="clarify the CO and its learning relevance, increase guided demonstrations and collect mid-course feedback"
        elif "semester-end" in weak_name: intervention="conduct additional exam-oriented application/case-analysis practice, discuss model answers and use a pre-SEE formative test"
        else: intervention="introduce short diagnostic quizzes, guided problem-solving/tutorial sessions and timely question-level feedback before the next internal assessment"
        action=f"For {label}, {intervention}. Align the remedial activity and rubric explicitly with “{statement}” at {bloom}. Verify effectiveness through a documented reassessment and compare its attainment with the present {gap:.2f} gap in the next cycle."
    return rca,action


def _po_analysis(label: str, value: float, target: float | None, contributors: list[str]) -> tuple[str,str]:
    source=", ".join(contributors) if contributors else "no mapped CO"
    if value<=0 or target is None: return (f"{label} is not mapped to an active CO in this course; attainment analysis is not applicable.","")
    if value>=target:
        return (f"{label} attained the target ({value:.2f} against {target:.2f}) through contributions from {source}.","")
    gap=target-value
    return (f"{label} did not attain the target: attainment {value:.2f}, target {target:.2f}, gap {gap:.2f}. The result is driven by the mapped outcomes {source}.",f"Strengthen the learning activities and assessment rubrics of {source}, include explicit evidence for {label}, and verify improvement in the next cycle against the present {gap:.2f} gap.")


def fill_template(template: Path, output: Path, course: Course, students: list[Student], maxima: tuple[list[float], list[float], list[float]], survey: list[dict[str, Any]], result: dict[str, Any], mapping: list[list[float]], co_targets: dict[str,list[float]], po_targets: dict[str,list[float]], direct_weights: tuple[float, float], overall_weights: tuple[float, float], co_labels: list[str], outcome_labels: list[str]) -> None:
    shutil.copy2(template, output)
    wb = load_workbook(output)
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    ia_max, mid_max, see_max = maxima

    ws = _sheet(wb, "2. TARGET")
    for cell, value in {"E3":course.school,"M3":course.program,"N3":course.year,"B4":course.semester,"E4":course.code,"I4":course.name,"M4":course.faculty,"M5":course.course_type}.items(): ws[cell] = value
    target_cells = _label_cells(ws, co_labels + outcome_labels)
    for i, label in enumerate(co_labels):
        for header in target_cells.get(label, []):
            if header.row <= 20:
                ws.cell(header.row+1,header.column,co_targets["previous"][i]); ws.cell(header.row+2,header.column,co_targets["attained"][i]); ws.cell(header.row+3,header.column,co_targets["current"][i]); break
    for i, label in enumerate(outcome_labels):
        for header in target_cells.get(label, []):
            if header.row >= 15:
                ws.cell(header.row+1,header.column,po_targets["previous"][i]); ws.cell(header.row+2,header.column,po_targets["attained"][i]); ws.cell(header.row+3,header.column,po_targets["current"][i]); break

    ws = _sheet(wb, "3. CO-PO Mapping")
    ws["Q4"], ws["Q5"] = course.odd_even, course.section
    map_headers = {label:[] for label in outcome_labels}
    for cell in ws[31]:
        key=text(cell.value).upper().replace(" ","")
        if key in map_headers: map_headers[key].append(cell)
    for i in range(len(co_labels)):
        ws.cell(13+i, 4, course.cos[i]); ws.cell(13+i, 17, course.bloom[i]); ws.cell(32+i, 1, f"{course.code}.{i+1}" if course.code else f"CO{i+1}")
        for p,label in enumerate(outcome_labels):
            headers=map_headers.get(label,[])
            if headers: ws.cell(32+i, headers[0].column, mapping[i][p] or None)

    ws = _sheet(wb, "4. CIE Assessment Marks")
    for row in range(14, 49):
        for col in range(2, 54): ws.cell(row, col, None)
    max_cells = ["D12","L12","T12","AB12"]
    for i, cell in enumerate(max_cells): ws[cell] = ia_max[i]
    for i, col in enumerate(range(32,36)): ws.cell(12,col,mid_max[i])
    total_cols = [46,47,48,49,50,51]
    for i in range(4): ws.cell(12,total_cols[i],ia_max[i]+mid_max[i])
    ws["AZ12"] = sum(ia_max)+sum(mid_max)
    for index, student in enumerate(students, start=14):
        ws.cell(index,2,student.name); ws.cell(index,3,student.reg); ws.cell(index,4,student.ia[0]); ws.cell(index,12,student.ia[1]); ws.cell(index,20,student.ia[2]); ws.cell(index,28,student.ia[3])
        for i in range(4): ws.cell(index,32+i,student.mid[i]); ws.cell(index,46+i,student.ia[i]+student.mid[i])
        ws.cell(index,52,sum(student.ia)+sum(student.mid)); ws.cell(index,53,student.grade)

    ws = _sheet(wb, "5. CIE- CO Attainment")
    for i in range(4): ws.cell(21,4+i,result["cie_max"][i])
    for r, student in enumerate(students, start=22):
        idx=r-22; ws.cell(r,2,student.name); ws.cell(r,3,student.reg); ws.cell(r,22,student.grade)
        for c in range(4):
            ws.cell(r,4+c,result["cie_marks"][idx][c]); ws.cell(r,10+2*c,result["cie_marks"][idx][c]/result["cie_max"][c]*100 if result["cie_max"][c] else 0); ws.cell(r,11+2*c,result["cie_levels"][idx][c])
    for c in range(4): ws.cell(15,9+c,result["cie"][c])
    ws["C57"] = len(students)

    ws = _sheet(wb, "6. SEE - Marks")
    starts=[4,9,14,19]
    for i,col in enumerate(starts): ws.cell(13,col,see_max[i])
    for r,student in enumerate(students,start=15):
        ws.cell(r,2,student.name); ws.cell(r,3,student.reg); ws.cell(r,41,student.grade)
        for c,col in enumerate(starts): ws.cell(r,col,student.see[c]); ws.cell(r,34+c,student.see[c])
        ws.cell(r,40,sum(student.see))

    ws = _sheet(wb, "7. SEE- CO Attainment")
    for i in range(4): ws.cell(21,4+i,see_max[i])
    for r,student in enumerate(students,start=22):
        idx=r-22; ws.cell(r,2,student.name); ws.cell(r,3,student.reg); ws.cell(r,22,student.grade)
        for c in range(4):
            ws.cell(r,4+c,student.see[c]); ws.cell(r,10+2*c,student.see[c]/see_max[c]*100 if see_max[c] else 0); ws.cell(r,11+2*c,result["see_levels"][idx][c])
    for c in range(4): ws.cell(15,9+c,result["see"][c])
    ws["C57"] = len(students)

    ws = _sheet(wb, "8. Course Feedback")
    for c in range(len(co_labels)): ws.cell(24,5+c,course.cos[c])
    for row in range(26,60):
        for col in range(2,15): ws.cell(row,col,None)
    for r,response in enumerate(survey,start=26):
        ws.cell(r,1,r-25)
        ws.cell(r,2,response["name"]); ws.cell(r,3,response["reg"])
        for c,rating in enumerate(response["ratings"]):
            ws.cell(r,5+c,rating); ws.cell(r,11+c,3 if rating==5 else 2 if rating>=3 else 1)
    for c in range(4):
        levels=[3 if x["ratings"][c]==5 else 2 if x["ratings"][c]>=3 else 1 for x in survey]
        ws.cell(12,11+c,levels.count(3)); ws.cell(13,11+c,levels.count(2)); ws.cell(14,11+c,levels.count(1)); ws.cell(16,11+c,len(survey)); ws.cell(17,11+c,len(students)); ws.cell(18,11+c,sum(levels)); ws.cell(19,11+c,result["indirect"][c] if survey else "NA")
    ws["A20"] = f"Course outcome survey response rate: {len(survey)} of {len(students)} students ({len(survey)/len(students)*100:.0f}%)." if students else "No enrolled students found."
    ws["A20"].fill=PatternFill("solid",fgColor="D9EAF7"); ws["A20"].font=Font(color="1F4E78",italic=True)

    ws = _sheet(wb, "9.Direct & Overall CO attinment")
    ws["B14"],ws["B15"] = direct_weights; ws["C19"],ws["C20"] = overall_weights
    for c in range(4):
        target=co_targets["current"][c]; col=4+c; ws.cell(14,col,result["cie"][c]); ws.cell(15,col,result["see"][c]); ws.cell(16,col,result["direct"][c]); ws.cell(19,col,result["direct"][c]); ws.cell(20,col,result["indirect"][c] if survey else "NA"); ws.cell(21,col,result["overall"][c]); ws.cell(22,col,target); ws.cell(23,col,"Y" if result["overall"][c]>=target else "N")

    ws = _sheet(wb, "10. ACTION PLAN & RCA-CO")
    co_rca=[]
    for i,target in enumerate(co_targets["current"]):
        rca,action=_co_analysis(i,course,result,target);co_rca.append(rca);ws.cell(8+i,8,action)
    ws["A20"]="Outcome analysis: "+" ".join(co_rca)
    ws["A15"] = course.faculty

    ws = _sheet(wb, "11. PO ATTAINMENT")
    for c in range(4): ws.cell(9+c,1,result["direct"][c]); ws.cell(22+c,1,result["overall"][c])
    attainment_headers = _label_cells(ws, outcome_labels)
    for p,label in enumerate(outcome_labels):
        target=po_targets["current"][p]
        header=next((x for x in attainment_headers.get(label,[]) if x.row==8),None)
        if not header: continue
        for value,value_row in ((result["direct_po"][p],15),(result["overall_po"][p],28)):
            mapped=value>0 and target is not None
            ws.cell(value_row,header.column,value if mapped else None)
            ws.cell(value_row+1,header.column,target if mapped else None)
            ws.cell(value_row+2,header.column,("Y" if value>=target else "N") if mapped else "NA")

    ws = _sheet(wb, "12. ACTION PLAN & RCA-PO")
    po_rca=[]
    action_labels=_label_cells(ws,outcome_labels)
    for i,(label,value,target) in enumerate(zip(outcome_labels,result["overall_po"],po_targets["current"])):
        contributors=[f"CO{c+1}" for c in range(len(course.cos)) if mapping[c][i]>0]
        rca,action=_po_analysis(label,value,target,contributors);po_rca.append((label,rca))
        row=next((cell.row for cell in action_labels.get(label,[]) if 9<=cell.row<=24),None)
        if row: ws.cell(row,8,action)
    ws["A26"]=course.faculty
    ws["A32"]="Program Outcome (PO) – Root Cause Analysis\n\n"+" ".join(rca for label,rca in po_rca if label.startswith("PO"))
    ws["A34"]="Program Specific Outcome (PSO) – Root Cause Analysis\n\n"+" ".join(rca for label,rca in po_rca if label.startswith("PSO"))
    wb.save(output)


def run_job(files: dict[str, Path | None], fields: dict[str, str], output_dir: Path) -> tuple[Path, dict[str, Any]]:
    template_info=inspect_template(files["template"])
    co_labels=json.loads(fields.get("co_labels") or "null") or template_info["cos"][:4]
    outcome_labels=json.loads(fields.get("outcome_labels") or "null") or template_info["outcomes"]
    if not co_labels: raise ValueError("At least one active CO is required.")
    if any(x not in template_info["cos"] for x in co_labels) or any(x not in template_info["outcomes"] for x in outcome_labels): raise ValueError("Outcome labels do not match the uploaded template. Please select the template again.")
    course = Course(code=fields.get("course_code", ""), name=fields.get("course_name", ""), faculty=fields.get("faculty", ""), school=fields.get("school", ""), program=fields.get("program", ""), semester=fields.get("semester", ""), year=fields.get("year", ""), odd_even=fields.get("odd_even", ""), section=fields.get("section", ""), course_type=fields.get("course_type", ""), cos=[""]*len(co_labels), bloom=[""]*len(co_labels))
    course = parse_course_plan(files.get("course_plan"), course)
    for i in range(len(co_labels)):
        if fields.get(f"co{i+1}"): course.cos[i]=fields[f"co{i+1}"]
        if fields.get(f"bloom{i+1}"): course.bloom[i]=fields[f"bloom{i+1}"]
    if files.get("marks"):
        students, ia_max, mid_max, see_max = parse_marks(files["marks"])
        warnings=[]
    else:
        raw,ia_max,mid_max,see_max,warnings=normalize_assessments(files["qp_analysis"],files.get("ia_marks") or [],files["midsem_marks"],files["final_marks"],co_labels,json.loads(fields.get("manual_splits") or "[]"),fields.get("mapping_source",""))
        students=[Student(x["reg"],x["name"],x["ia"],x["mid"],x["see"]) for x in raw]
    warnings += apply_grades(files.get("grades"), students, course.code)
    if files.get("grades"):
        pending=[s.reg for s in students if not s.grade]
        if pending:
            warnings.append("Students without a confirmed final grade were excluded from attainment: "+", ".join(pending)+". Their cells remain pending, not zero.")
            students=[s for s in students if s.grade]
        if not students: raise ValueError("No students have a confirmed final grade; attainment cannot be calculated yet.")
    registrations=[s.reg for s in students]
    if len(registrations)!=len(set(registrations)): raise ValueError("Duplicate registration numbers were found in the marks file.")
    students.sort(key=lambda s: (0,int(s.reg)) if s.reg.isdigit() else (1,s.reg))
    survey = parse_survey(files.get("survey"),len(co_labels))
    if survey and len(survey) < len(students): warnings.append(f"Survey response rate is {len(survey)} of {len(students)} students.")
    direct=(number(fields.get("cie_weight"),60),number(fields.get("see_weight"),40)); overall=(number(fields.get("direct_weight"),80),number(fields.get("indirect_weight"),20))
    if not math.isclose(sum(direct),100,abs_tol=0.01): raise ValueError("CIE and SEE weights must total 100%.")
    if not math.isclose(sum(overall),100,abs_tol=0.01): raise ValueError("Direct and indirect weights must total 100%.")
    mapping=json.loads(fields.get("mapping",json.dumps(DEFAULT_MAPPING)))
    if len(mapping)!=len(co_labels) or any(len(row)!=len(outcome_labels) for row in mapping) or any(number(v,-1)<0 or number(v,-1)>3 for row in mapping for v in row): raise ValueError(f"CO–PO/PSO mapping must be a {len(co_labels)} × {len(outcome_labels)} matrix with values from 0 to 3.")
    def targets(field: str, count: int, default: float, allow_blank: bool=False) -> dict[str,list[float|None]]:
        raw=json.loads(fields.get(field) or "{}"); result={}
        for key in ("previous","attained","current"):
            values=raw.get(key,[default]*count)
            if len(values)!=count or any(v is not None and not 0<=number(v,-1)<=3 for v in values): raise ValueError(f"All entered {field.replace('_',' ')} values must be between 0 and 3.")
            if not allow_blank and any(v is None for v in values): raise ValueError(f"Complete all {field.replace('_',' ')} values.")
            result[key]=[None if v is None else number(v) for v in values]
        return result
    co_targets=targets("co_targets",len(co_labels),2.4); po_targets=targets("po_targets",len(outcome_labels),2.0,True)
    for p in range(len(outcome_labels)):
        if any(mapping[c][p]>0 for c in range(len(co_labels))) and po_targets["current"][p] is None: raise ValueError(f"Enter the current target for mapped outcome {outcome_labels[p]}.")
    known=set(registrations)
    unmatched=[x["reg"] for x in survey if x["reg"] not in known]
    if unmatched: warnings.append("Survey registration numbers not found in marks: "+", ".join(unmatched))
    for student in students:
        for label,marks,maxima in [("IA",student.ia,ia_max),("Mid-semester",student.mid,mid_max),("SEE",student.see,see_max)]:
            if any(mark<0 or mark>maxima[i] for i,mark in enumerate(marks)): raise ValueError(f"{label} marks for {student.reg} fall outside the permitted CO maximum marks.")
    result=compute(students,ia_max,mid_max,see_max,survey,direct,overall,mapping,0)
    output_dir.mkdir(parents=True,exist_ok=True); label=course.code.replace(' ','_') or 'Course'; name=f"{uuid.uuid4().hex[:10]}_CO_PO_Attainment_{label}.xlsx"; output=output_dir/name
    fill_template(files["template"],output,course,students,(ia_max,mid_max,see_max),survey,result,mapping,co_targets,po_targets,direct,overall,co_labels,outcome_labels)
    summary={"file":name,"students":len(students),"survey_responses":len(survey),"warnings":warnings,"cie":[round(x,3) for x in result["cie"]],"see":[round(x,3) for x in result["see"]],"direct":[round(x,3) for x in result["direct"]],"indirect":[round(x,3) for x in result["indirect"]] if survey else ["NA"]*len(co_labels),"overall":[round(x,3) for x in result["overall"]],"po":[round(x,3) if x>0 else None for x in result["overall_po"]],"targets":co_targets["current"],"co_status":["Y" if x>=co_targets["current"][i] else "N" for i,x in enumerate(result["overall"])]}
    return output,summary
