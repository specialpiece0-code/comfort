# -*- coding: utf-8 -*-
"""20개 부처 입찰/지원·공모 데이터를 단일 워크북으로 통합.
- 맨 앞: 총합 시트 (입찰: 부→처/청 2단계 구분, 아래쪽: 지원·공모 부 구분)
- 이후: 부처별 시트 20개 (입찰: 처/청 구분, 아래쪽: 지원·공모)
"""
import glob, os, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC_DIR = "data"
OUT = "정부_입찰_지원공모_통합_2025.xlsx"

BID_COLS = ["공고일","업무구분","공고번호","공고명","계약방법","낙찰방법","추정가격",
            "배정예산","공고기관","수요기관","입찰개시","입찰마감","개찰일시","공고상세URL"]
SUP_COLS = ["등록일","공고명","소관기관","수행기관","지원분야","세부분야","지원대상",
            "신청기간","신청방법","상세URL"]
NBID, NSUP = len(BID_COLS), len(SUP_COLS)

# ---- 스타일 ----
F_TITLE   = Font(bold=True, size=14, color="FFFFFF")
FILL_TITLE= PatternFill("solid", fgColor="1F4E78")
F_SECT    = Font(bold=True, size=12, color="FFFFFF")
FILL_SECT = PatternFill("solid", fgColor="2E75B6")
F_BU      = Font(bold=True, size=12, color="FFFFFF")
FILL_BU   = PatternFill("solid", fgColor="305496")
F_CHEO    = Font(bold=True, size=10, color="1F3864")
FILL_CHEO = PatternFill("solid", fgColor="D9E1F2")
F_HDR     = Font(bold=True, color="FFFFFF")
FILL_HDR  = PatternFill("solid", fgColor="44546A")
ALIGN_L   = Alignment(horizontal="left", vertical="center")
THIN      = Side(style="thin", color="BFBFBF")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 부처 파일 순서: 01~19 먼저, 00_기타_독립기관 마지막 (전체요약 순번과 동일)
def dept_order(files):
    named = {os.path.basename(f): f for f in files}
    seq = [f for n,f in sorted(named.items()) if not n.startswith("00_")]
    etc = named.get("00_기타_독립기관.xlsx")
    return seq + ([etc] if etc else [])

def read_sheet(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    # 완전히 빈 행 제거
    rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
    wb.close()
    return rows

def load_dept(path):
    """부처 파일 -> {name, units:[(unit_name, rows)], support_rows:[...] }"""
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets = wb.sheetnames
    wb.close()
    name = os.path.splitext(os.path.basename(path))[0]
    units = []
    for s in sheets:
        if s in ("요약", "지원·공모"):
            continue
        units.append((s, read_sheet(path, s)))
    support = read_sheet(path, "지원·공모") if "지원·공모" in sheets else []
    return {"name": name, "units": units, "support": support}

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

BID_WIDTHS = [19,9,16,44,14,22,14,14,26,26,19,19,19,30]
SUP_WIDTHS = [19,40,16,18,12,14,14,22,28,30]

def write_data_row(ws, r, values, ncols):
    for ci in range(ncols):
        v = values[ci] if ci < len(values) else None
        cell = ws.cell(row=r, column=ci+1, value=v)
        cell.border = BORDER
        if ci in (6,7) and isinstance(v, (int, float)):  # 추정가격/배정예산
            cell.number_format = "#,##0"
    return r+1

def merge_banner(ws, r, ncols, text, font, fill):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    c = ws.cell(row=r, column=1, value=text)
    c.font = font; c.fill = fill; c.alignment = ALIGN_L
    ws.row_dimensions[r].height = 22 if fill in (FILL_BU, FILL_SECT, FILL_TITLE) else 18
    return r+1

def write_header(ws, r, cols, ncols):
    for ci, name in enumerate(cols, 1):
        c = ws.cell(row=r, column=ci, value=name)
        c.font = F_HDR; c.fill = FILL_HDR; c.alignment = ALIGN_L; c.border = BORDER
    ws.row_dimensions[r].height = 18
    return r+1

def build_dept_sheet(wb, dept):
    title = dept["name"]
    sheet_name = title[:31]
    ws = wb.create_sheet(title=sheet_name)
    set_widths(ws, BID_WIDTHS)
    bid_total = sum(len(rows) for _, rows in dept["units"])
    r = 1
    r = merge_banner(ws, r, NBID, f"{title}  (입찰 {bid_total:,}건)", F_TITLE, FILL_TITLE)
    r += 1
    r = merge_banner(ws, r, NBID, "【입찰】", F_SECT, FILL_SECT)
    r = write_header(ws, r, BID_COLS, NBID)
    ws.freeze_panes = ws.cell(row=r, column=1)
    for unit, rows in dept["units"]:
        r = merge_banner(ws, r, NBID, f"─ {unit} ({len(rows):,})", F_CHEO, FILL_CHEO)
        for vals in rows:
            r = write_data_row(ws, r, vals, NBID)
    # 지원·공모 구역
    if dept["support"]:
        r += 2
        r = merge_banner(ws, r, NBID, f"【지원·공모】 ({len(dept['support']):,}건)", F_SECT, FILL_SECT)
        r = write_header(ws, r, SUP_COLS, NSUP)
        for vals in dept["support"]:
            r = write_data_row(ws, r, vals, NSUP)
    return ws

def build_total_sheet(wb, depts):
    ws = wb.create_sheet(title="총합")
    set_widths(ws, BID_WIDTHS)
    bid_total = sum(len(rows) for d in depts for _, rows in d["units"])
    sup_total = sum(len(d["support"]) for d in depts)
    r = 1
    r = merge_banner(ws, r, NBID,
        f"2025년 정부 부처 입찰·지원공모 통합 총합  (입찰 {bid_total:,}건 / 지원·공모 {sup_total:,}건)",
        F_TITLE, FILL_TITLE)
    r += 1
    # ---- 입찰 ----
    r = merge_banner(ws, r, NBID, "【입찰】", F_SECT, FILL_SECT)
    r = write_header(ws, r, BID_COLS, NBID)
    ws.freeze_panes = ws.cell(row=r, column=1)
    for d in depts:
        d_total = sum(len(rows) for _, rows in d["units"])
        r = merge_banner(ws, r, NBID, f"■■■ {d['name']}  (입찰 {d_total:,}건)  ■■■", F_BU, FILL_BU)
        for unit, rows in d["units"]:
            r = merge_banner(ws, r, NBID, f"  ─ {unit} ({len(rows):,})", F_CHEO, FILL_CHEO)
            for vals in rows:
                r = write_data_row(ws, r, vals, NBID)
    # ---- 지원·공모 ----
    r += 2
    r = merge_banner(ws, r, NBID, f"【지원·공모】  (총 {sup_total:,}건)", F_SECT, FILL_SECT)
    r = write_header(ws, r, SUP_COLS, NSUP)
    for d in depts:
        if not d["support"]:
            continue
        r = merge_banner(ws, r, NSUP, f"■■■ {d['name']}  (지원·공모 {len(d['support']):,}건)  ■■■", F_BU, FILL_BU)
        for vals in d["support"]:
            r = write_data_row(ws, r, vals, NSUP)
    return ws

def main():
    files = dept_order(glob.glob(os.path.join(SRC_DIR, "*.xlsx")))
    files = [f for f in files if os.path.basename(f) != "00_전체요약.xlsx"]
    print(f"부처 파일 {len(files)}개 로드 중...")
    depts = []
    for f in files:
        d = load_dept(f)
        bt = sum(len(r) for _, r in d["units"])
        print(f"  {d['name']}: 입찰 {bt:,} / 지원·공모 {len(d['support']):,} / 단위 {len(d['units'])}")
        depts.append(d)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    print("총합 시트 작성...")
    build_total_sheet(wb, depts)
    print("부처별 시트 작성...")
    for d in depts:
        build_dept_sheet(wb, d)

    # 시트 순서: 총합 먼저
    wb.move_sheet("총합", -(len(wb.sheetnames)-1))
    print(f"저장: {OUT}  (시트 {len(wb.sheetnames)}개)")
    wb.save(OUT)
    print("완료:", wb.sheetnames)

if __name__ == "__main__":
    main()
