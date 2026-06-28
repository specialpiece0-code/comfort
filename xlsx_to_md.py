# -*- coding: utf-8 -*-
"""xlsx 파일을 NotebookLM 인식용 Markdown으로 변환.
- 파일 1개 -> MD 1개 (모든 시트를 GitHub 마크다운 표로)
- 사용: python xlsx_to_md.py <out_dir> <file1.xlsx> [file2.xlsx ...]
"""
import sys, os, re, openpyxl

def cell(v):
    if v is None: return ""
    s = str(v)
    s = s.replace("\r", " ").replace("\n", "<br>")
    s = s.replace("|", "\\|")
    return s.strip()

def sheet_to_md(ws):
    rows = list(ws.iter_rows(values_only=True))
    # 끝쪽 빈 열 제거를 위해 최대 유효 열 계산
    ncols = 0
    for r in rows:
        for j in range(len(r) - 1, -1, -1):
            if r[j] is not None and str(r[j]).strip() != "":
                ncols = max(ncols, j + 1); break
    if ncols == 0: return "_(빈 시트)_\n"
    # 완전 빈 행 제거
    data = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
    if not data: return "_(빈 시트)_\n"
    out = []
    header = [cell(data[0][j] if j < len(data[0]) else "") for j in range(ncols)]
    header = [h if h else f"열{i+1}" for i, h in enumerate(header)]
    out.append("| " + " | ".join(header) + " |")
    out.append("| " + " | ".join(["---"] * ncols) + " |")
    for r in data[1:]:
        out.append("| " + " | ".join(cell(r[j] if j < len(r) else "") for j in range(ncols)) + " |")
    return "\n".join(out) + "\n"

def convert(path, out_dir):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    # 부처명: 요약 시트 B1, 없으면 파일명
    name = None
    if "요약" in wb.sheetnames:
        r1 = next(wb["요약"].iter_rows(min_row=1, max_row=1, values_only=True))
        if len(r1) > 1 and r1[1]: name = str(r1[1]).strip()
    if not name:
        name = os.path.splitext(os.path.basename(path))[0]
    parts = [f"# {name}\n"]
    for s in wb.sheetnames:
        parts.append(f"\n## {s}\n")
        parts.append(sheet_to_md(wb[s]))
    wb.close()
    safe = re.sub(r'[\\/:*?"<>|]', "_", name)
    outp = os.path.join(out_dir, f"{safe}.md")
    with open(outp, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return outp, name

if __name__ == "__main__":
    out_dir = sys.argv[1]; os.makedirs(out_dir, exist_ok=True)
    for p in sys.argv[2:]:
        outp, name = convert(p, out_dir)
        sz = os.path.getsize(outp)
        print(f"{name}: {outp}  ({sz/1024:.0f} KB)")
