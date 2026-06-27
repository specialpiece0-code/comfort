# -*- coding: utf-8 -*-
"""20개 부처 입찰/지원·공모 데이터를 단일 워크북으로 통합 (열 구분 방식).
- 구분은 행이 아니라 앞쪽 '부처' / '처·청' 열로 표시 -> 정렬·필터 용이.
- 맨 앞: 총합 시트 (입찰 전체 + 아래쪽 지원·공모 전체)
- 이후: 부처별 시트 20개 (입찰 + 아래쪽 지원·공모)
"""
import glob, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC_DIR = "data"
OUT = "정부_입찰_지원공모_통합_2025.xlsx"

# 앞에 부처 / 처·청 구분 열을 추가
BID_BASE = ["공고일","업무구분","공고번호","공고명","계약방법","낙찰방법","추정가격",
            "배정예산","공고기관","수요기관","입찰개시","입찰마감","개찰일시","공고상세URL"]
SUP_BASE = ["등록일","공고명","소관기관","수행기관","지원분야","세부분야","지원대상",
            "신청기간","신청방법","상세URL"]
BID_COLS = ["부처","처·청"] + BID_BASE          # 16
SUP_COLS = ["부처"] + SUP_BASE                   # 11
NBID, NSUP = len(BID_COLS), len(SUP_COLS)
# 숫자 서식 적용할 열(0-based): 추정가격, 배정예산 -> 입찰에서 2칸 밀림
BID_NUM_IDX = {8, 9}

# ---- 스타일 ----
F_TITLE   = Font(bold=True, size=14, color="FFFFFF")
FILL_TITLE= PatternFill("solid", fgColor="1F4E78")
F_SECT    = Font(bold=True, size=12, color="FFFFFF")
FILL_SECT = PatternFill("solid", fgColor="2E75B6")
F_HDR     = Font(bold=True, color="FFFFFF")
FILL_HDR  = PatternFill("solid", fgColor="44546A")
F_BU      = Font(bold=True, color="1F3864")
FILL_BU   = PatternFill("solid", fgColor="D9E1F2")   # 부처 열
FILL_CHEO = PatternFill("solid", fgColor="EDF2FA")   # 처·청 열
ALIGN_L   = Alignment(horizontal="left", vertical="center")
THIN      = Side(style="thin", color="BFBFBF")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BID_WIDTHS = [20,16,19,9,16,40,14,22,14,14,24,24,19,19,19,30]
SUP_WIDTHS = [20,19,38,16,18,12,14,14,22,28,30]

def dept_order(files):
    named = {os.path.basename(f): f for f in files}
    seq = [f for n,f in sorted(named.items()) if not n.startswith("00_")]
    etc = named.get("00_기타_독립기관.xlsx")
    return seq + ([etc] if etc else [])

def read_sheet(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
            if any(c is not None and str(c).strip() != "" for c in r)]
    wb.close()
    return rows

def load_dept(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets = wb.sheetnames
    wb.close()
    name = os.path.splitext(os.path.basename(path))[0]
    units = [(s, read_sheet(path, s)) for s in sheets if s not in ("요약","지원·공모")]
    support = read_sheet(path, "지원·공모") if "지원·공모" in sheets else []
    return {"name": name, "units": units, "support": support}

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def merge_banner(ws, r, ncols, text, font, fill, h=22):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(row=r, column=1, value=text)
    c.font = font; c.fill = fill; c.alignment = ALIGN_L
    ws.row_dimensions[r].height = h
    return r+1

def write_header(ws, r, cols):
    for ci, name in enumerate(cols, 1):
        c = ws.cell(row=r, column=ci, value=name)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = ALIGN_L; c.border = BORDER
    ws.row_dimensions[r].height = 18
    return r+1

def write_bid_row(ws, r, bu, unit, values):
    a = ws.cell(row=r, column=1, value=bu);   a.font=F_BU; a.fill=FILL_BU;   a.border=BORDER
    b = ws.cell(row=r, column=2, value=unit); b.fill=FILL_CHEO; b.border=BORDER
    for k in range(len(BID_BASE)):
        v = values[k] if k < len(values) else None
        cell = ws.cell(row=r, column=3+k, value=v); cell.border = BORDER
        if (2+k) in BID_NUM_IDX and isinstance(v,(int,float)):
            cell.number_format = "#,##0"
    return r+1

def write_sup_row(ws, r, bu, values):
    a = ws.cell(row=r, column=1, value=bu); a.font=F_BU; a.fill=FILL_BU; a.border=BORDER
    for k in range(len(SUP_BASE)):
        v = values[k] if k < len(values) else None
        cell = ws.cell(row=r, column=2+k, value=v); cell.border = BORDER
    return r+1

def build_total_sheet(wb, depts):
    ws = wb.create_sheet(title="총합")
    set_widths(ws, BID_WIDTHS)
    bid_total = sum(len(rows) for d in depts for _,rows in d["units"])
    sup_total = sum(len(d["support"]) for d in depts)
    r = 1
    r = merge_banner(ws, r, NBID,
        f"2025년 정부 부처 입찰·지원공모 통합 총합  (입찰 {bid_total:,}건 / 지원·공모 {sup_total:,}건)",
        F_TITLE, FILL_TITLE)
    r += 1
    r = merge_banner(ws, r, NBID, "【입찰】", F_SECT, FILL_SECT)
    r = write_header(ws, r, BID_COLS)
    ws.freeze_panes = ws.cell(row=r, column=3)      # 헤더행 + 부처/처·청 열 고정
    for d in depts:
        for unit, rows in d["units"]:
            for vals in rows:
                r = write_bid_row(ws, r, d["name"], unit, vals)
    r += 2
    r = merge_banner(ws, r, NBID, f"【지원·공모】  (총 {sup_total:,}건)", F_SECT, FILL_SECT)
    r = write_header(ws, r, SUP_COLS)
    for d in depts:
        for vals in d["support"]:
            r = write_sup_row(ws, r, d["name"], vals)
    return ws

def build_dept_sheet(wb, dept):
    ws = wb.create_sheet(title=dept["name"][:31])
    set_widths(ws, BID_WIDTHS)
    bid_total = sum(len(rows) for _,rows in dept["units"])
    r = 1
    r = merge_banner(ws, r, NBID, f"{dept['name']}  (입찰 {bid_total:,}건)", F_TITLE, FILL_TITLE)
    r += 1
    r = merge_banner(ws, r, NBID, "【입찰】", F_SECT, FILL_SECT)
    r = write_header(ws, r, BID_COLS)
    ws.freeze_panes = ws.cell(row=r, column=3)
    for unit, rows in dept["units"]:
        for vals in rows:
            r = write_bid_row(ws, r, dept["name"], unit, vals)
    if dept["support"]:
        r += 2
        r = merge_banner(ws, r, NBID, f"【지원·공모】 ({len(dept['support']):,}건)", F_SECT, FILL_SECT)
        r = write_header(ws, r, SUP_COLS)
        for vals in dept["support"]:
            r = write_sup_row(ws, r, dept["name"], vals)
    return ws

def main():
    files = [f for f in dept_order(glob.glob(os.path.join(SRC_DIR,"*.xlsx")))
             if os.path.basename(f) != "00_전체요약.xlsx"]
    print(f"부처 파일 {len(files)}개 로드...")
    depts = [load_dept(f) for f in files]
    for d in depts:
        bt = sum(len(r) for _,r in d["units"])
        print(f"  {d['name']}: 입찰 {bt:,} / 지원·공모 {len(d['support']):,} / 단위 {len(d['units'])}")
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    print("총합 시트 작성..."); build_total_sheet(wb, depts)
    print("부처별 시트 작성...")
    for d in depts: build_dept_sheet(wb, d)
    wb.move_sheet("총합", -(len(wb.sheetnames)-1))
    wb.save(OUT)
    print(f"저장: {OUT}  (시트 {len(wb.sheetnames)}개)")

if __name__ == "__main__":
    main()
