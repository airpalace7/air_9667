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
import datetime
import openpyxl
import warnings
warnings.filterwarnings("ignore")

OUT_DIR = "data"


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
        days = int(total // 86400)
        rem = total - days * 86400
        h = int(rem // 3600)
        m = int((rem % 3600) // 60)
        sign = "-" if neg else ""
        if days > 0:
            s = f"{sign}{days}일 {h:02d}:{m:02d}"
        else:
            s = f"{sign}{h:02d}:{m:02d}"
        return s, neg
    if isinstance(v, (int, float)):
        neg = v < 0
        if isinstance(v, float) and not v.is_integer():
            s = f"{v:.2f}".rstrip("0").rstrip(".")
        else:
            s = str(int(v))
        return s, neg
    return str(v), False


def convert_insp(ws):
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
    sections = []
    cur = None
    for r in range(7, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if a and isinstance(a, str) and a.strip().startswith("▶"):
            cur = {"title": a.strip(), "rows": []}
            sections.append(cur)
            continue
        if cur is None:
            continue
        if a is None and b is None:
            continue
        item = fmt_val(a)[0]
        if not item:
            continue
        interval_day = fmt_val(ws.cell(r, 2).value)[0]
        interval_time = fmt_val(ws.cell(r, 3).value)[0]
        performed_date = fmt_val(ws.cell(r, 4).value)[0]
        performed_time = fmt_val(ws.cell(r, 5).value)[0]
        next_date = fmt_val(ws.cell(r, 6).value)[0]
        next_time = fmt_val(ws.cell(r, 7).value)[0]
        remain_day_raw = ws.cell(r, 8).value
        remain_day, neg_day = fmt_val(remain_day_raw)
        remain_time = fmt_val(ws.cell(r, 9).value)[0]
        remark = fmt_val(ws.cell(r, 10).value)[0]
        overdue = isinstance(remain_day_raw, (int, float)) and remain_day_raw < 0
        cur["rows"].append({
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
        })
    return {"info": info, "sections": sections}


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

    import os
    os.makedirs(OUT_DIR, exist_ok=True)

    data_insp = convert_insp(wb["HL9667 INSP"])
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
