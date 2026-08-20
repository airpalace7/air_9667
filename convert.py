#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
9667_테스트.xlsx -> data/*.json 변환 스크립트

사용법:
    python3 convert.py 9667_테스트.xlsx

엑셀을 수정한 뒤 이 스크립트를 다시 실행하면 data/ 폴더의 JSON 4개가
새로 생성됩니다. 그 JSON 파일들을 GitHub 저장소에 업로드(덮어쓰기)하면
웹페이지에 반영됩니다.
"""
import sys
import json
import re
import datetime
import openpyxl
import warnings
warnings.filterwarnings("ignore")

OUT_DIR = "data"


def hours_of(v):
    """timedelta -> 소수 시간(float). 아니면 None."""
    if isinstance(v, datetime.timedelta):
        return v.total_seconds() / 3600.0
    return None


def date_iso(v):
    """datetime -> 'YYYY-MM-DD'. 아니면 None."""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    return None


TSN_BASE_RE = {
    "ac": re.compile(r"^=SUM\(G\d+-B\$3\)$"),
    "eng1": re.compile(r"^=SUM\(G\d+-\$?F\$?2\)$"),
    "eng2": re.compile(r"^=SUM\(G\d+-\$?H\$?2\)$"),
}


CELL_REF_RE = re.compile(r"^=([A-Z]+)(\d+)$")


def ref_info(ws, formula_val):
    """'=D41' 같은 단순 셀 참조 수식이면 참조 대상 행/항목명을 돌려준다. 아니면 None."""
    if not isinstance(formula_val, str):
        return None
    m = CELL_REF_RE.match(formula_val)
    if not m:
        return None
    ref_row = int(m.group(2))
    ref_item = fmt_val(ws.cell(ref_row, 1).value)[0]
    return {"row": ref_row, "item": ref_item}


def is_dash_or_none(v):
    return v is None or (isinstance(v, str) and v.strip() == "-")


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def classify_row(ws_formula, r):
    """엑셀 수식을 보고 이 행이 표준 패턴(JS 재계산 가능)인지 판단."""
    f = ws_formula.cell(r, 6).value
    g = ws_formula.cell(r, 7).value
    h = ws_formula.cell(r, 8).value
    i = ws_formula.cell(r, 9).value

    ok_g = is_dash_or_none(g) or (isinstance(g, str) and re.match(r"^=SUM\(E\d+\+C\d+\)$", g))
    ok_f = is_dash_or_none(f) or (isinstance(f, str) and re.match(r"^=SUM\(D\d+\+B\d+\)$", f))
    ok_h = is_dash_or_none(h) or (isinstance(h, str) and re.match(r"^=F\d+-\$?L\$?2$", h))

    base = None
    for name, pat in TSN_BASE_RE.items():
        if isinstance(i, str) and pat.match(i):
            base = name
            break
    ok_i = is_dash_or_none(i) or (base is not None)

    recalculable = ok_f and ok_g and ok_h and ok_i
    return recalculable, base


def fmt_val(v):
    """셀 값을 표시용 문자열로 변환. 반환값: (문자열, 원값이 음수/초과인지 여부)"""
    if v is None:
        return "", False
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d"), False
    if isinstance(v, datetime.timedelta):
        total = v.total_seconds()
        neg = total < 0
        total = abs(total)
        total_hours = int(total // 3600)
        m = int((total % 3600) // 60)
        sign = "-" if neg else ""
        s = f"{sign}{total_hours}:{m:02d}"
        return s, neg
    if isinstance(v, (int, float)):
        neg = v < 0
        if isinstance(v, float) and not v.is_integer():
            s = f"{v:.2f}".rstrip("0").rstrip(".")
        else:
            s = str(int(v))
        return s, neg
    return str(v), False


def convert_insp(ws, ws_formula):
    info = {
        "등록기호": fmt_val(ws["B2"].value)[0],
        "TSN": fmt_val(ws["B3"].value)[0],
        "L/D": fmt_val(ws["D3"].value)[0],
        "1ENG_TSN": fmt_val(ws["F2"].value)[0],
        "2ENG_TSN": fmt_val(ws["H2"].value)[0],
        "1ENG_GG_CYC": fmt_val(ws["F3"].value)[0],
        "2ENG_GG_CYC": fmt_val(ws["H3"].value)[0],
        "1ENG_PT_CYC": fmt_val(ws["F4"].value)[0],
        "2ENG_PT_CYC": fmt_val(ws["H4"].value)[0],
        "DATE": fmt_val(ws["L2"].value)[0],
    }
    raw = {
        "ac_tsn_hours": hours_of(ws["B3"].value),
        "eng1_tsn_hours": hours_of(ws["F2"].value),
        "eng2_tsn_hours": hours_of(ws["H2"].value),
        "today": date_iso(ws["L2"].value),
    }
    sections = []
    cur = None
    pending_log_row = None  # "CARGO HOOK - REMOVED DATE" 같은 로그형 항목 뒤에 오는 빈 행들을 담을 곳
    for r in range(7, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if a and isinstance(a, str) and a.strip().startswith("▶"):
            cur = {"title": a.strip(), "rows": []}
            sections.append(cur)
            pending_log_row = None
            continue
        if cur is None:
            continue
        if a is None and b is None:
            # item 칸이 비어있는 행. 바로 위가 "날짜 로그" 헤더 행이었다면(예: CARGO HOOK -
            # REMOVED DATE), 이 행의 D/E열을 제거일/재장착일 한 쌍으로 로그에 추가한다.
            if pending_log_row is not None:
                removed = date_iso(ws.cell(r, 4).value)
                reinstalled = date_iso(ws.cell(r, 5).value)
                pending_log_row["log_entries"].append({
                    "removed_date_iso": removed,
                    "reinstalled_date_iso": reinstalled,
                })
            continue
        item = fmt_val(a)[0]
        if not item:
            continue
        interval_day_raw = ws.cell(r, 2).value
        interval_time_raw = ws.cell(r, 3).value
        performed_date_raw = ws.cell(r, 4).value
        performed_time_raw = ws.cell(r, 5).value
        interval_day = fmt_val(interval_day_raw)[0]
        interval_time = fmt_val(interval_time_raw)[0]
        performed_date = fmt_val(performed_date_raw)[0]
        performed_time = fmt_val(performed_time_raw)[0]
        next_date = fmt_val(ws.cell(r, 6).value)[0]
        next_time = fmt_val(ws.cell(r, 7).value)[0]
        remain_day_raw = ws.cell(r, 8).value
        remain_day, neg_day = fmt_val(remain_day_raw)
        remain_time = fmt_val(ws.cell(r, 9).value)[0]
        remark = fmt_val(ws.cell(r, 10).value)[0]
        overdue = isinstance(remain_day_raw, (int, float)) and remain_day_raw < 0

        recalculable, tsn_base = classify_row(ws_formula, r)
        # D열(PERFORMED DATE)/E열(PERFORMED A/C TIME) 자체가 다른 행을 참조하는 수식이면
        # (예: =D41) 그 항목은 독립적으로 입력하는 값이 아니라 다른 행의 값을 그대로 따라가는
        # 것이므로 웹에서 직접 수정할 수 없게 한다. 그 외에는(원본이 raw 값이면) 편집 가능.
        performed_date_editable = not is_formula(ws_formula.cell(r, 4).value)
        performed_time_editable = not is_formula(ws_formula.cell(r, 5).value)
        performed_date_ref = ref_info(ws, ws_formula.cell(r, 4).value)
        performed_time_ref = ref_info(ws, ws_formula.cell(r, 5).value)

        # "CARGO HOOK - REMOVED DATE"처럼 D/E열이 실제 값이 아니라 "제거일"/"재장착일" 같은
        # 열 라벨로 쓰이는 헤더 행: 그 자체는 입력칸이 아니므로 편집 불가로 두고, 뒤따르는
        # 빈 행들을 로그 항목으로 모은다.
        is_date_log = isinstance(performed_date_raw, str) and isinstance(performed_time_raw, str) \
            and not is_formula(performed_date_raw) and not is_formula(performed_time_raw) \
            and performed_date_raw.strip() != "" and performed_time_raw.strip() != ""

        row_obj = {
            "item": item,
            "interval_day": interval_day,
            "interval_time": interval_time,
            "performed_date": performed_date,
            "performed_time": performed_time,
            "next_date": next_date,
            "next_time": next_time,
            "remain_day": remain_day,
            "remain_time": remain_time,
            "remark": remark,
            "overdue": overdue,
            "recalc": recalculable,
            "tsn_base": tsn_base,
            "performed_date_editable": False if is_date_log else performed_date_editable,
            "performed_time_editable": False if is_date_log else performed_time_editable,
            "performed_date_ref": None if is_date_log else performed_date_ref,
            "performed_time_ref": None if is_date_log else performed_time_ref,
            "interval_day_num": interval_day_raw if isinstance(interval_day_raw, (int, float)) else None,
            "interval_time_hours": hours_of(interval_time_raw),
            "performed_date_iso": None if is_date_log else date_iso(performed_date_raw),
            "performed_time_hours": None if is_date_log else hours_of(performed_time_raw),
        }
        if is_date_log:
            row_obj["date_log"] = True
            row_obj["log_entries"] = []
            pending_log_row = row_obj
        else:
            pending_log_row = None
        cur["rows"].append(row_obj)
    return {"info": info, "raw": raw, "sections": sections}


def convert_component(ws, header_rows_end, data_start):
    info_lines = []
    for r in range(1, header_rows_end):
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        if b:
            info_lines.append({"label": str(b), "value": fmt_val(c)[0]})
    rows = []
    for r in range(data_start, ws.max_row + 1):
        no = ws.cell(r, 1).value
        name = ws.cell(r, 2).value
        if no is None and not name:
            continue
        remaining_time_raw = ws.cell(r, 17).value
        remaining_time, neg = fmt_val(remaining_time_raw)
        overdue = isinstance(remaining_time_raw, (int, float)) and remaining_time_raw < 0
        if isinstance(remaining_time_raw, datetime.timedelta) and remaining_time_raw.total_seconds() < 0:
            overdue = True
        next_date_raw = ws.cell(r, 16).value
        if isinstance(next_date_raw, datetime.datetime):
            today_cell = ws.parent["AC TRP, TBO"]["C1"].value if "AC TRP, TBO" in ws.parent.sheetnames else None
            today_dt = today_cell if isinstance(today_cell, datetime.datetime) else datetime.datetime.now()
            if next_date_raw < today_dt:
                overdue = True
        rows.append({
            "no": fmt_val(no)[0],
            "name": fmt_val(name)[0],
            "pn": fmt_val(ws.cell(r, 3).value)[0],
            "sn": fmt_val(ws.cell(r, 4).value)[0],
            "type": fmt_val(ws.cell(r, 5).value)[0],
            "exchange_cycle": fmt_val(ws.cell(r, 6).value)[0],
            "installation_time": fmt_val(ws.cell(r, 8).value)[0],
            "installation_date": fmt_val(ws.cell(r, 9).value)[0],
            "location": fmt_val(ws.cell(r, 10).value)[0],
            "tsn": fmt_val(ws.cell(r, 11).value)[0],
            "usage_time": fmt_val(ws.cell(r, 13).value)[0],
            "next_exchange_time": fmt_val(ws.cell(r, 15).value)[0],
            "next_exchange_date": fmt_val(ws.cell(r, 16).value)[0],
            "remaining_time": remaining_time,
            "note": fmt_val(ws.cell(r, 19).value)[0],
            "overdue": overdue,
        })
    return {"info": info_lines, "rows": rows}


def main():
    if len(sys.argv) < 2:
        print("사용법: python3 convert.py <엑셀파일.xlsx>")
        sys.exit(1)
    path = sys.argv[1]
    wb = openpyxl.load_workbook(path, data_only=True)
    wb_formula = openpyxl.load_workbook(path, data_only=False)

    import os
    os.makedirs(OUT_DIR, exist_ok=True)

    data_insp = convert_insp(wb["HL9667 INSP"], wb_formula["HL9667 INSP"])
    with open(f"{OUT_DIR}/insp.json", "w", encoding="utf-8") as f:
        json.dump(data_insp, f, ensure_ascii=False, indent=2)

    data_trp = convert_component(wb["AC TRP, TBO"], header_rows_end=9, data_start=10)
    with open(f"{OUT_DIR}/trp.json", "w", encoding="utf-8") as f:
        json.dump(data_trp, f, ensure_ascii=False, indent=2)

    data_eng1 = convert_component(wb["#1 ENG TRP, TBO"], header_rows_end=10, data_start=12)
    with open(f"{OUT_DIR}/eng1.json", "w", encoding="utf-8") as f:
        json.dump(data_eng1, f, ensure_ascii=False, indent=2)

    data_eng2 = convert_component(wb["#2 ENG TRP, TBO"], header_rows_end=10, data_start=12)
    with open(f"{OUT_DIR}/eng2.json", "w", encoding="utf-8") as f:
        json.dump(data_eng2, f, ensure_ascii=False, indent=2)

    print("완료: data/insp.json, data/trp.json, data/eng1.json, data/eng2.json 생성됨")


if __name__ == "__main__":
    main()
