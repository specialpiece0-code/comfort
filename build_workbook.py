# -*- coding: utf-8 -*-
"""20개 부처 입찰/지원·공모 데이터를 단일 워크북으로 통합 (열 구분 + 정제).
- 구분은 행이 아니라 앞쪽 '부처' / '처·청' / '긴급' 열로 표시 -> 정렬·필터 용이.
- 정제: ①같은 공고번호는 최신 차수 1건만 ②제목에 재공고/정정/유찰/취소/연기 포함 행 삭제
        ③긴급공고는 '긴급' 열(Y)로 표기해 보존
- 맨 앞: 총합 시트 (입찰 전체 + 아래쪽 지원·공모 전체)
- 이후: 부처별 시트 20개
"""
import glob, os, re, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC_DIR = "data"
OUT = "정부_입찰_지원공모_통합_2025.xlsx"

KILL = ["재공고", "정정", "유찰", "취소", "연기"]   # 공고명에 포함 시 삭제
URGENT = "긴급"

# 앞에 부처 / 처·청 / 긴급 구분 열을 추가
BID_BASE = ["공고일","업무구분","공고번호","공고명","계약방법","낙찰방법","추정가격",
            "배정예산","공고기관","수요기관","입찰개시","입찰마감","개찰일시","공고상세URL"]
SUP_BASE = ["등록일","공고명","소관기관","수행기관","지원분야","세부분야","지원대상",
            "신청기간","신청방법","상세URL"]
BID_COLS = ["부처","처·청","긴급"] + BID_BASE      # 17
SUP_COLS = ["부처"] + SUP_BASE                     # 11
NBID, NSUP = len(BID_COLS), len(SUP_COLS)
# BID_BASE 내 숫자 서식 인덱스: 추정가격(6), 배정예산(7)
BID_BASE_NUM = {6, 7}
# BID_BASE 내 위치
I_NO, I_NAME, I_URL = 2, 3, 13

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
F_URG     = Font(bold=True, color="C00000")
FILL_URG  = PatternFill("solid", fgColor="FCE4E4")   # 긴급 열
ALIGN_C   = Alignment(horizontal="center", vertical="center")
ALIGN_L   = Alignment(horizontal="left", vertical="center")
THIN      = Side(style="thin", color="BFBFBF")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

BID_WIDTHS = [20,16,6,19,9,16,40,14,22,14,14,24,24,19,19,19,30]   # 17
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

def parse_cha(url):
    m = re.search(r'bidPbancOrd=(\d+)', str(url or ""))
    return int(m.group(1)) if m else -1

def is_kill(name):
    s = str(name or "")
    return any(k in s for k in KILL)

def clean_bids(depts):
    """차수 최신만 + 키워드 삭제 적용. 각 행에 긴급 플래그 부여.
    반환: (정제된 depts 구조, 통계 dict). units 의 rows -> (urgent, values)."""
    # 1) 모든 입찰 행 수집
    recs = []
    for di, d in enumerate(depts):
        for ui, (unit, rows) in enumerate(d["units"]):
            for ri, vals in enumerate(rows):
                no = vals[I_NO] if len(vals) > I_NO else None
                recs.append({
                    "di": di, "ui": ui, "ri": ri, "no": no,
                    "name": vals[I_NAME] if len(vals) > I_NAME else "",
                    "cha": parse_cha(vals[I_URL] if len(vals) > I_URL else ""),
                    "urgent": URGENT in str(vals[I_NAME] if len(vals) > I_NAME else ""),
                })
    total = len(recs)
    # 2) 공고번호별 최신 차수만
    best = {}
    keep_ids = set()
    for idx, rc in enumerate(recs):
        rc["id"] = idx
        no = rc["no"]
        if no is None:
            keep_ids.add(idx)             # 공고번호 없는 행은 그대로 유지
            continue
        cur = best.get(no)
        if cur is None or rc["cha"] > cur["cha"]:
            best[no] = rc
    for rc in best.values():
        keep_ids.add(rc["id"])
    after_dedup = len(keep_ids)
    # 3) 키워드 삭제
    kill_hits = {k: 0 for k in KILL}
    final_ids = set()
    for idx in keep_ids:
        nm = str(recs[idx]["name"] or "")
        hit = [k for k in KILL if k in nm]
        if hit:
            for k in hit: kill_hits[k] += 1
        else:
            final_ids.add(idx)
    after_kill = len(final_ids)
    # 4) 재구성 (원래 순서 보존) + 긴급 카운트
    loc = {(rc["di"], rc["ui"], rc["ri"]): rc for rc in recs}
    new_depts, urgent_cnt = [], 0
    for di, d in enumerate(depts):
        new_units = []
        for ui, (unit, rows) in enumerate(d["units"]):
            kept = []
            for ri, vals in enumerate(rows):
                rc = loc[(di, ui, ri)]
                if rc["id"] in final_ids:
                    kept.append((rc["urgent"], vals))
                    if rc["urgent"]: urgent_cnt += 1
            new_units.append((unit, kept))
        new_depts.append({"name": d["name"], "units": new_units, "support": d["support"]})
    stats = {"total": total, "after_dedup": after_dedup, "after_kill": after_kill,
             "removed_dup": total - after_dedup, "removed_kill": after_dedup - after_kill,
             "kill_hits": kill_hits, "urgent": urgent_cnt}
    return new_depts, stats

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

def write_bid_row(ws, r, bu, unit, urgent, values):
    a = ws.cell(row=r, column=1, value=bu);   a.font=F_BU; a.fill=FILL_BU;   a.border=BORDER
    b = ws.cell(row=r, column=2, value=unit); b.fill=FILL_CHEO; b.border=BORDER
    u = ws.cell(row=r, column=3, value="Y" if urgent else None)
    u.border=BORDER; u.alignment=ALIGN_C
    if urgent: u.font=F_URG; u.fill=FILL_URG
    for k in range(len(BID_BASE)):
        v = values[k] if k < len(values) else None
        cell = ws.cell(row=r, column=4+k, value=v); cell.border = BORDER
        if k in BID_BASE_NUM and isinstance(v,(int,float)):
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
    ws.freeze_panes = ws.cell(row=r, column=4)      # 헤더행 + 부처/처·청/긴급 열 고정
    for d in depts:
        for unit, rows in d["units"]:
            for urgent, vals in rows:
                r = write_bid_row(ws, r, d["name"], unit, urgent, vals)
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
    ws.freeze_panes = ws.cell(row=r, column=4)
    for unit, rows in dept["units"]:
        for urgent, vals in rows:
            r = write_bid_row(ws, r, dept["name"], unit, urgent, vals)
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
    depts, st = clean_bids(depts)
    print("\n[정제 결과]")
    print(f"  입찰 원본            : {st['total']:,}")
    print(f"  -최신차수만(중복제거): -{st['removed_dup']:,}  -> {st['after_dedup']:,}")
    print(f"  -키워드 삭제         : -{st['removed_kill']:,}  -> {st['after_kill']:,}")
    print(f"     세부: " + " / ".join(f"{k} {v:,}" for k,v in st['kill_hits'].items()))
    print(f"  긴급(Y) 표기         : {st['urgent']:,}건")
    print(f"  최종 입찰 행         : {st['after_kill']:,}\n")
    for d in depts:
        bt = sum(len(r) for _,r in d["units"])
        print(f"  {d['name']}: 입찰 {bt:,} / 지원·공모 {len(d['support']):,}")
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    print("총합 시트 작성..."); build_total_sheet(wb, depts)
    print("부처별 시트 작성...")
    for d in depts: build_dept_sheet(wb, d)
    wb.move_sheet("총합", -(len(wb.sheetnames)-1))
    wb.save(OUT)
    print(f"저장: {OUT}  (시트 {len(wb.sheetnames)}개)")

if __name__ == "__main__":
    main()
