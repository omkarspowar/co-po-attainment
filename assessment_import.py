from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


def txt(v: Any) -> str:
    return "" if v is None else str(v).strip()


def regno(v: Any) -> str:
    s=txt(v)
    return s[:-2] if s.endswith(".0") else s


def num(v: Any) -> float:
    try: return float(v)
    except (TypeError,ValueError): return 0.0


def qkey(v: Any) -> str:
    key=re.sub(r"[^A-Z0-9]","",txt(v).upper()).removeprefix("QUESTION")
    # Exam exports commonly prefix optional subparts with "[OR]".
    key=key.removeprefix("OR")
    return key[1:] if key.startswith("Q") and len(key)>1 and key[1].isdigit() else key


def question_header(v: Any) -> str | None:
    """Return a canonical Q1/Q1A-style header from an exam-export cell."""
    raw=txt(v).upper()
    if not re.search(r"\bQ\s*\d",raw) and not re.search(r"\b\d+\s*\)\s*[A-Z]\b",raw):
        return None
    key=qkey(raw)
    match=re.fullmatch(r"(\d+)([A-Z]?)",key)
    return f"Q{match.group(1)}{match.group(2)}" if match else None


@dataclass
class Question:
    key: str
    maximum: float
    co: str
    bloom: str=""


@dataclass
class Assessment:
    name: str
    kind: str
    maximum: float
    questions: list[Question]


def parse_qp_analysis(path: Path) -> list[Assessment]:
    if path.suffix.lower() != ".docx":
        raise ValueError("QP Analysis auto-reading currently requires a .docx file.")
    doc=Document(path)
    summary=[]
    for table in doc.tables:
        rows=[[txt(c.text) for c in row.cells] for row in table.rows]
        if rows and any("name of the assessment" in x.lower() for x in rows[0]):
            for row in rows[1:]:
                if len(row)>=3 and row[1] and "total" not in row[1].lower(): summary.append((row[1],num(row[2])))
            break
    maps=[]
    for table in doc.tables:
        rows=[[txt(c.text) for c in row.cells] for row in table.rows]
        if not rows: continue
        header_row=next((i for i,row in enumerate(rows[:3]) if any("mark" in x.lower() for x in row) and any(re.fullmatch(r"\s*co(?:\*|/clo|\s*/clo)?\s*",x,re.I) for x in row)),0)
        head=[x.lower() for x in rows[header_row]]
        qi=next((i for i,x in enumerate(head) if "q." in x or "q.no" in x or "q no" in x),None)
        mi=next((i for i,x in enumerate(head) if "mark" in x),None)
        ci=next((i for i,x in enumerate(head) if x in {"co","co*","co/clo","co /clo"} or "co/clo" in x),None)
        bi=next((i for i,x in enumerate(head) if "bloom" in x or x=="bl"),None)
        if qi is None or mi is None or ci is None: continue
        questions=[]
        for row in rows[header_row+1:]:
            if max(qi,mi,ci)>=len(row): continue
            co_match=re.search(r"\bCO\s*(\d+)\b",row[ci],re.I)
            if not co_match or not qkey(row[qi]): continue
            questions.append(Question(qkey(row[qi]),num(row[mi]),f"CO{co_match.group(1)}",row[bi] if bi is not None and bi<len(row) else ""))
        if questions: maps.append(questions)
    if not maps: raise ValueError("No question-to-CO mapping tables were found in the QP Analysis document.")
    assessments=[]
    for i,questions in enumerate(maps):
        name,maximum=summary[i] if i<len(summary) else (f"Assessment {i+1}",sum(q.maximum for q in questions))
        low=name.lower(); kind="mid" if "mid" in low else "final" if "end" in low or "final" in low or "see" in low else "ia"
        assessments.append(Assessment(name,kind,maximum or sum(q.maximum for q in questions),questions))
    return assessments


def is_reg_header(value: Any) -> bool:
    normalized=re.sub(r"[^a-z]","",txt(value).lower())
    return "registration" in normalized or "enrollmentid" in normalized or normalized in {"regno","regnumber","registernumber","usn","rollno","rollnum","studentid","orgdefinedid"}


def _parse_mark_rows(rows: list[list[Any]], source: str) -> tuple[list[str],list[dict[str,Any]]]:
    header_i=-1; inferred_reg_i=None
    for i,row in enumerate(rows):
        if any(is_reg_header(x) for x in row): header_i=i;break
    if header_i<0:
        # Infer the registration column from a student row when the header is absent or unconventional.
        for i,row in enumerate(rows):
            candidate=next((c for c,value in enumerate(row) if re.fullmatch(r"\d{6,}",regno(value))),None)
            if candidate is not None: header_i=max(0,i-1);inferred_reg_i=candidate;break
    if header_i<0: raise ValueError(f"Could not locate a registration-number column in {source}.")
    raw_header=rows[header_i]
    headers=[txt(x) or f"COL{i+1}" for i,x in enumerate(raw_header)]
    # Examination exports often place Q1..Qn, including Q1) A / [OR] Q1) B,
    # across several secondary header rows.
    for candidate_row in rows[header_i+1:header_i+10]:
        for i,value in enumerate(candidate_row):
            canonical=question_header(value)
            if canonical: headers[i]=canonical
    reg_i=next((i for i,x in enumerate(headers) if is_reg_header(x)),inferred_reg_i if inferred_reg_i is not None else 0)
    name_candidates=[i for i,x in enumerate(headers) if "name" in x.lower()]
    if name_candidates:
        def name_score(column: int) -> int:
            return sum(bool(txt(row[column])) and not re.fullmatch(r"\d{6,}",regno(row[column])) for row in rows[header_i+1:] if column<len(row))
        name_i=max(name_candidates,key=name_score)
    else: name_i=max(0,reg_i-1) if reg_i>0 else min(1,len(headers)-1)
    records=[]
    for row in rows[header_i+1:]:
        row=list(row)+[None]*(len(headers)-len(row)); reg=regno(row[reg_i])
        if not re.fullmatch(r"\d{6,}",reg): continue
        value_indices=[i for i in range(len(headers)) if i not in {reg_i,name_i}]
        records.append({"reg":reg,"name":txt(row[name_i]),
                        "values":{headers[i]:num(row[i]) for i in value_indices},
                        "missing":{headers[i] for i in value_indices if row[i] is None or txt(row[i]) in {"","-"}}})
    if not records: raise ValueError(f"No student records were found in {source}.")
    return headers,records


def read_mark_table(path: Path, preferred_score_count: int | None=None) -> tuple[list[str],list[dict[str,Any]]]:
    if path.suffix.lower()==".csv":
        with path.open(encoding="utf-8-sig",newline="") as f: return _parse_mark_rows(list(csv.reader(f)),path.name)
    wb=load_workbook(path,data_only=True,read_only=False)
    candidates=[]
    for ws in wb.worksheets:
        try:
            parsed=_parse_mark_rows([list(r) for r in ws.iter_rows(values_only=True)],f"{path.name}/{ws.title}")
        except ValueError:
            continue
        cols=_score_columns(parsed[1]);missing=sum(c in r["missing"] for r in parsed[1] for c in cols)
        exact=preferred_score_count is not None and len(cols)==preferred_score_count
        candidates.append((exact,-missing,len(parsed[1]),ws.title,parsed))
    if not candidates: raise ValueError(f"No student mark table was found in {path.name}.")
    if preferred_score_count is not None and any(x[0] for x in candidates): candidates=[x for x in candidates if x[0]]
    return max(candidates,key=lambda x:(x[0],x[1],x[2]))[-1]


def embedded_co_totals(path: Path, co_labels: list[str]) -> tuple[dict[str,list[float]],list[float]] | None:
    """Read an exam export's own per-CO totals and the mapping-derived maxima."""
    if path.suffix.lower()==".csv": return None
    wanted={x.upper():i for i,x in enumerate(co_labels)};wb=load_workbook(path,data_only=True,read_only=False)
    for ws in wb.worksheets:
        rows=[list(r) for r in ws.iter_rows(values_only=True)]
        header_i=next((i for i,row in enumerate(rows) if any(is_reg_header(x) for x in row)),None)
        if header_i is None: continue
        header=rows[header_i];reg_i=next(i for i,x in enumerate(header) if is_reg_header(x))
        name_i=next((i for i,x in enumerate(header) if "name" in txt(x).lower()),max(0,reg_i-1))
        total_i=next((i for i,x in enumerate(header) if "total" in txt(x).lower()),len(header))
        co_total_row=None;co_columns={}
        for row in rows[header_i:header_i+8]:
            found={wanted[txt(v).upper()]:i for i,v in enumerate(row) if txt(v).upper() in wanted and i>total_i}
            if len(found)>=2: co_total_row=row;co_columns=found;break
        marks_i=next((i for i,row in enumerate(rows[header_i:header_i+15],header_i) if any("question marks" in txt(v).lower() for v in row)),None)
        if co_total_row is None or marks_i is None: continue
        mapping_i=next((i for i,row in enumerate(rows[marks_i+1:marks_i+6],marks_i+1) if sum(txt(v).upper() in wanted for v in row[:total_i])>=2),None)
        if mapping_i is None: continue
        maxima=[0.0]*len(co_labels)
        for c in range(min(len(rows[marks_i]),len(rows[mapping_i]),total_i)):
            co=txt(rows[mapping_i][c]).upper()
            if co not in wanted and mapping_i+1<len(rows): co=txt(rows[mapping_i+1][c]).upper()
            # Some exports put B/C subpart maxima in the preceding "Question No"
            # row because of merged cells, while A remains in "Question Marks".
            maximum=num(rows[marks_i][c]) or (num(rows[marks_i-1][c]) if marks_i>0 else 0.0)
            if co in wanted: maxima[wanted[co]]+=maximum
        totals={}
        for row in rows[header_i+1:]:
            padded=list(row)+[None]*(max([reg_i,name_i]+list(co_columns.values()))+1-len(row));reg=regno(padded[reg_i])
            if re.fullmatch(r"\d{6,}",reg): totals[reg]=[num(padded[co_columns[i]]) if i in co_columns else 0.0 for i in range(len(co_labels))]
        if totals: return totals,maxima
    return None


def mapping_conflict(path: Path, assessment: Assessment, co_labels: list[str]) -> dict[str,Any] | None:
    embedded=embedded_co_totals(path,co_labels)
    if not embedded: return None
    _,embedded_max=embedded;qp_max=[0.0]*len(co_labels);index={co:i for i,co in enumerate(co_labels)}
    for q in assessment.questions:
        if q.co in index: qp_max[index[q.co]]+=q.maximum
    if all(abs(a-b)<0.001 for a,b in zip(embedded_max,qp_max)): return None
    return {"assessment":assessment.name,"file":path.name,"qp_maxima":qp_max,"embedded_maxima":embedded_max}


def match_question(header: str, questions: list[Question]) -> list[Question]:
    key=qkey(header)
    direct=[q for q in questions if q.key==key]
    if direct: return direct
    return [q for q in questions if re.match(rf"^{re.escape(key)}[A-Z]$",q.key)]


def preview_conflicts(mark_path: Path, assessment: Assessment) -> list[dict[str,Any]]:
    _,records=read_mark_table(mark_path); conflicts=[]
    headers=list(records[0]["values"])
    for header in headers:
        matched=match_question(header,assessment.questions)
        cos=sorted({q.co for q in matched})
        if len(cos)>1:
            conflicts.append({"assessment":assessment.name,"column":header,"cos":cos,"subparts":[{"question":q.key,"co":q.co,"maximum":q.maximum} for q in matched],"students":[{"reg":r["reg"],"name":r["name"],"mark":r["values"][header]} for r in records]})
    return conflicts


def _score_columns(records: list[dict[str,Any]]) -> list[str]:
    columns=list(records[0]["values"])
    rejected=("total","difference","percent","adjusted final","final marks","final grade","see grade","sl.","s.no","serial","section","adoc","name","orgdefined","student id","registration","enrollment","roll no","rollnum","reg no")
    columns=[c for c in columns if not any(x in c.lower() for x in rejected)]
    question_columns=[c for c in columns if re.fullmatch(r"Q\d+[A-Z]?",c,re.I)]
    if question_columns: return question_columns
    if len(columns)>1:
        last=columns[-1]; prior=columns[:-1]
        if sum(abs(r["values"][last]-sum(r["values"][c] for c in prior))<0.001 for r in records)>=max(1,len(records)-1): columns=prior
    return columns


def normalize_assessments(qp_path: Path, ia_paths: list[Path], mid_path: Path, final_path: Path, co_labels: list[str], manual_splits: list[dict[str,Any]], mapping_source: str="") -> tuple[list[dict[str,Any]],list[float],list[float],list[float],list[str]]:
    assessments=parse_qp_analysis(qp_path); ia_maps=[a for a in assessments if a.kind=="ia"]; mid_map=next((a for a in assessments if a.kind=="mid"),None); final_map=next((a for a in assessments if a.kind=="final"),None)
    if not mid_map or not final_map: raise ValueError("QP Analysis must contain Midsem and final-examination question mappings.")
    index={co:i for i,co in enumerate(co_labels)}; students={}; warnings=[]
    def student(reg,name):
        row=students.setdefault(reg,{"reg":reg,"name":name,"ia":[0.0]*len(co_labels),"mid":[0.0]*len(co_labels),"see":[0.0]*len(co_labels),"_sources":set()})
        if name and not row["name"]: row["name"]=name
        return row
    ia_max=[0.0]*len(co_labels);mid_max=[0.0]*len(co_labels);see_max=[0.0]*len(co_labels)
    # Combined IA file: one score column per IA; separate files: one file per IA.
    ia_cursor=0
    for path in ia_paths:
        _,records=read_mark_table(path,len(ia_maps) if len(ia_paths)==1 else 1); cols=_score_columns(records)
        missing=sum(col in record["missing"] for record in records for col in cols)
        if missing: warnings.append(f"{path.name}: {missing} blank/hyphen IA mark cells require confirmation.")
        maps=ia_maps[ia_cursor:ia_cursor+len(cols)] if len(ia_paths)==1 else ia_maps[ia_cursor:ia_cursor+1]
        if len(ia_paths)>1: cols=cols[:1]
        if len(maps)!=len(cols): raise ValueError(f"Could not match IA columns in {path.name} to the QP Analysis assessments.")
        for col,assessment in zip(cols,maps):
            cos={q.co for q in assessment.questions};
            if len(cos)!=1: raise ValueError(f"{assessment.name} contains multiple COs; upload question-wise IA marks.")
            co=next(iter(cos)); ci=index.get(co)
            if ci is None: continue
            ia_max[ci]+=assessment.maximum
            for record in records:
                row=student(record["reg"],record["name"]);row["_sources"].add("ia");row["ia"][ci]+=record["values"][col]
        ia_cursor+=len(maps)
    if ia_cursor<len(ia_maps): warnings.append(f"Only {ia_cursor} of {len(ia_maps)} internal assessments were matched.")
    split_index={(x["assessment"],qkey(x["column"]),x["reg"]):x for x in manual_splits}
    def apply(path,assessment,bucket,maxima):
        embedded=embedded_co_totals(path,co_labels)
        conflict=mapping_conflict(path,assessment,co_labels)
        if conflict and mapping_source not in {"embedded","qp"}:
            raise ValueError(f"{assessment.name} has different QP and marks-sheet CO mappings. Select which mapping to use in Assessment review.")
        if embedded and (mapping_source=="embedded" or not conflict):
            totals,embedded_max=embedded
            for i,value in enumerate(embedded_max): maxima[i]+=value
            for reg,values in totals.items():
                row=student(reg,"");row["_sources"].add(bucket);row[bucket]=values[:]
            if conflict: warnings.append(f"{assessment.name}: faculty selected the embedded marks-sheet CO totals instead of the conflicting QP mapping.")
            return
        _,records=read_mark_table(path);cols=_score_columns(records)
        missing=sum(col in record["missing"] for record in records for col in cols)
        if missing: warnings.append(f"{path.name}: {missing} blank/hyphen question marks require confirmation.")
        for col in cols:
            questions=match_question(col,assessment.questions)
            if not questions: continue
            by_co={}
            for q in questions: by_co[q.co]=by_co.get(q.co,0)+q.maximum
            for co,mx in by_co.items():
                if co in index: maxima[index[co]]+=mx
            for record in records:
                row=student(record["reg"],record["name"]);row["_sources"].add(bucket);target=row[bucket]
                if len(by_co)==1: target[index[next(iter(by_co))]]+=record["values"][col]
                else:
                    split=split_index.get((assessment.name,qkey(col),record["reg"]))
                    if not split: raise ValueError(f"Manual CO split is required for {assessment.name} {col}, registration {record['reg']}.")
                    total=sum(num(p.get("mark")) for p in split.get("parts",[]))
                    if abs(total-record["values"][col])>0.001: raise ValueError(f"Manual split for {record['reg']} {col} does not equal the original mark.")
                    for part in split["parts"]:
                        if part.get("co") in index: target[index[part["co"]]]+=num(part.get("mark"))
    apply(mid_path,mid_map,"mid",mid_max);apply(final_path,final_map,"see",see_max)
    ordered=sorted(students.values(),key=lambda x:(0,int(x["reg"])) if x["reg"].isdigit() else (1,x["reg"]))
    for source,label in (("ia","IA"),("mid","midterm"),("see","end-term")):
        absent=[row["reg"] for row in ordered if source not in row["_sources"]]
        if absent: warnings.append(f"{len(absent)} student(s) are missing from the {label} file: {', '.join(absent)}. Confirm absence before calculation.")
    for row in ordered: row.pop("_sources",None)
    return ordered,ia_max,mid_max,see_max,warnings
