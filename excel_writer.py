"""수집 결과를 주간 양식(template.xlsx)에 채워 bytes로 반환"""
from __future__ import annotations

import io
from copy import copy
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TEMPLATE = Path(__file__).with_name("template.xlsx")
FIRST_ROW = 5
CHECK_FILL = PatternFill("solid", fgColor="FFF2CC")

# 양식 컬럼(A~AD) ↔ Rec 속성, 확인필요 플래그 이름
COLS = [
    ("공시일", None), ("납입일", "납입일"), ("발행사", None), ("종류", "공모/사모 구분 확인"),
    ("총액", "전자등록총액"), (lambda r: "미분류", None), (lambda r: "미분류", None),
    (lambda r: None, None),  # H 주관사: 직접 입력
    ("투자자", "투자자"), ("법인유형", None), ("코드", None),
    ("coupon", "Coupon"), ("ytm", "YTM"), ("ytp", "YTP"),
    ("put_start", "Put 시작"), ("put_cycle", "Put 주기"),
    ("ytc", "YTC"), ("call_start", "Call 시작"), ("call_end", "Call 종료"), ("call_cycle", "Call 주기"),
    ("call_pct", "Call %"), ("만기일", "만기일"), ("만기", None), ("전환개시", "전환개시"),
    ("할증율", "할증율"), ("refix_down", "전환가액"), ("refix_down_par", "하향Refixing(액면가o) 확인"),
    ("refix_up", "상향Refixing(자동감지)"), ("refix_cycle", "Refixing 주기"), ("신고서", None),
]


def _flagged(r, name):
    return bool(name) and any(f.startswith(name) for f in r.flags)


def write_excel(recs, start: date, end: date) -> bytes:
    wb = load_workbook(TEMPLATE)
    ws = wb.worksheets[0]
    ws["A3"] = f"   {start:%Y.%m.%d}~{end:%Y.%m.%d}"
    styles = {c: copy(ws.cell(FIRST_ROW, c)._style) for c in range(1, len(COLS) + 1)}
    height = ws.row_dimensions[FIRST_ROW].height

    for i, r in enumerate(recs):
        row = FIRST_ROW + i
        ws.row_dimensions[row].height = height
        for c, (src, flag) in enumerate(COLS, start=1):
            v = src(r) if callable(src) else getattr(r, src)
            cell = ws.cell(row, c)
            cell._style = copy(styles[c])
            cell.value = v
            if _flagged(r, flag) or (v is None and not callable(src)):
                cell.fill = CHECK_FILL

    # 검토용 시트
    ws2 = wb.create_sheet("검토용")
    head = ["공시일", "발행사", "종류", "회차", "공시명", "확인필요", "API 주관사", "DART 원문",
            "Put 원문발췌", "Call 원문발췌", "Refixing 원문발췌"]
    ws2.append(head)
    for c in ws2[1]:
        c.font = Font(bold=True)
    for r in recs:
        ws2.append([r.공시일, r.발행사, r.종류, r.회차, r.report_nm, ", ".join(r.flags),
                    r.주관사_API, r.link,
                    r.snippets.get("put", "")[:1000], r.snippets.get("call", "")[:1000],
                    r.snippets.get("refix", "")[:1000]])
        cell = ws2.cell(ws2.max_row, 8)
        cell.hyperlink = r.link
        cell.font = Font(color="0563C1", underline="single")
        ws2.cell(ws2.max_row, 1).number_format = "yyyy-mm-dd"
    widths = [12, 18, 10, 8, 40, 40, 14, 20, 60, 60, 60]
    for i, w in enumerate(widths, start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    for row in ws2.iter_rows(min_row=2):
        for c in row[8:]:
            c.alignment = Alignment(wrap_text=True, vertical="top")

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()
