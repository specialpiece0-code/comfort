# -*- coding: utf-8 -*-
"""
[로컬 실행용] 나라장터(g2b)·기업마당(bizinfo) 공고 URL에 접속해 사업 내용을
30자 이내로 요약하여 summaries.csv 로 저장한다.

특징
- 통합 워크북에서 모든 공고 URL을 자동 수집
- Playwright(실제 크롬)로 페이지 렌더링 후 본문 추출
- 요약: 기본은 규칙기반(공고명/사업개요 정리). ANTHROPIC_API_KEY 가 있으면 Claude 사용
- 중단·재실행 가능(resumable): 이미 끝난 URL은 summaries.csv 를 읽어 건너뜀
- 동시 N탭 처리 + 요청 간 지연 + 재시도

사용법 (Windows cmd, repo 폴더에서)
  python -m pip install playwright openpyxl
  python -m playwright install chromium
  python scrape_summaries.py                       # 기본 실행
  python scrape_summaries.py --concurrency 6 --delay 0.4
  set ANTHROPIC_API_KEY=sk-ant-...                 # (선택) LLM 요약 켜기
  python -m pip install anthropic                   # (선택) LLM 사용 시

결과: summaries.csv  (열: url, 공고번호, source, 사업요약)
  -> 이 파일을 repo에 커밋/푸시하면 제가 워크북에 '사업요약' 열을 합쳐드립니다.
"""
import argparse, asyncio, csv, os, re, sys
from pathlib import Path

import openpyxl

WORKBOOK_DEFAULT = "정부_입찰_지원공모_통합_2025.xlsx"
OUT_CSV = "summaries.csv"
FIELDS = ["url", "공고번호", "source", "사업요약"]

# 본문에서 사업 설명이 있을 만한 구간 라벨
HINT_LABELS = ["사업개요", "용역개요", "과업내용", "사업목적", "구매규격",
               "공고내용", "납품요구", "제안요청", "사업명", "용역명", "물품명"]
# 요약에서 걷어낼 잡음
NOISE = re.compile(r"(긴급|재공고|정정공고|\[.*?\]|\(.*?\)|제\d+호|\d{4}[-./]\d{1,2}[-./]\d{1,2})")


def collect_urls(workbook):
    """워크북의 모든 시트에서 http(s) URL을 (url, 공고번호) 로 수집(중복 제거)."""
    wb = openpyxl.load_workbook(workbook, data_only=True, read_only=True)
    seen = {}
    bno = re.compile(r"bidPbancNo=([A-Z0-9]+)")
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and v.startswith("http") and v not in seen:
                    m = bno.search(v)
                    seen[v] = m.group(1) if m else ""
    wb.close()
    return seen


def load_done(path):
    """요약이 실제로 채워진 URL만 '완료'로 간주(빈 요약은 재실행 시 재시도)."""
    done = set()
    if Path(path).exists():
        with open(path, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                if r.get("url") and (r.get("사업요약") or "").strip():
                    done.add(r["url"])
    return done


def trim30(text):
    s = NOISE.sub("", text or "")
    s = re.sub(r"\s+", " ", s).strip(" -·,/")
    return s[:30]


def rule_summary(page_text, gongomyeong):
    """규칙기반 요약: 본문 라벨 구간 우선, 없으면 공고명 정리."""
    if page_text:
        for kw in HINT_LABELS:
            i = page_text.find(kw)
            if i != -1:
                seg = page_text[i + len(kw): i + len(kw) + 80]
                seg = trim30(seg)
                if len(seg) >= 6:
                    return seg
    return trim30(gongomyeong)


# ---- (선택) LLM 요약 ----
_llm = None
def get_llm():
    global _llm
    if _llm is not None:
        return _llm
    if not os.environ.get("ANTHROPIC_API_KEY"):
        _llm = False; return _llm
    try:
        import anthropic
        _llm = anthropic.Anthropic()
    except Exception as e:
        print(f"[LLM 비활성] anthropic 사용 불가: {e}")
        _llm = False
    return _llm

def llm_summary(page_text, gongomyeong):
    client = get_llm()
    if not client:
        return None
    body = (page_text or gongomyeong or "")[:2000]
    try:
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=60,
            messages=[{"role": "user", "content":
                f"다음 공공 입찰/지원 공고의 핵심 사업 내용을 한국어 30자 이내로 "
                f"명사형 요약만 출력해줘(따옴표·설명 금지).\n공고명: {gongomyeong}\n본문:\n{body}"}],
        )
        return trim30(msg.content[0].text)
    except Exception:
        return None


async def fetch_one(page, url, delay, retries=2):
    """페이지 본문 텍스트와 공고명 추정값을 반환."""
    last = ""
    for attempt in range(retries + 1):
        try:
            await page.goto(url, wait_until="networkidle", timeout=30000)
            await page.wait_for_timeout(int(delay * 1000))
            text = await page.inner_text("body")
            title = ""
            try:
                title = (await page.title()) or ""
            except Exception:
                pass
            return text, title
        except Exception as e:
            last = str(e)[:80]
            await page.wait_for_timeout(800 * (attempt + 1))
    return None, last


async def worker(name, browser, queue, results, lock, args, stats):
    ctx = await browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36")
    page = await ctx.new_page()
    while True:
        try:
            url, gno = queue.get_nowait()
        except asyncio.QueueEmpty:
            break
        text, title = await fetch_one(page, url, args.delay)
        src = "g2b" if "g2b.go.kr" in url else ("bizinfo" if "bizinfo" in url else "etc")
        gongomyeong = title.split(" - ")[0].strip() if title else ""
        if text is None:
            summary = ""                       # 실패 → 빈칸(나중에 재실행으로 보완)
        else:
            summary = llm_summary(text, gongomyeong) or rule_summary(text, gongomyeong)
        async with lock:
            results.append({"url": url, "공고번호": gno, "source": src, "사업요약": summary})
            stats["done"] += 1
            if not summary:
                stats["empty"] += 1
            if len(results) >= 50:
                flush(results, args.out)
            if stats["done"] % 50 == 0:
                cum = stats["base"] + stats["done"]
                print(f"  진행 {cum:,}/{stats['total']:,}  (이번 세션 {stats['done']:,}, "
                      f"빈 요약 {stats['empty']:,})")
    await ctx.close()


def flush(results, out):
    new = not Path(out).exists()
    # append 모드: 이미 파일이 있으면 헤더 생략
    mode = "a" if not new else "w"
    with open(out, mode, encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        if new:
            w.writeheader()
        for r in results:
            w.writerow(r)
    results.clear()


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--workbook", default=WORKBOOK_DEFAULT)
    ap.add_argument("--out", default=OUT_CSV)
    ap.add_argument("--concurrency", type=int, default=4)
    ap.add_argument("--delay", type=float, default=0.5, help="페이지 로드 후 대기(초)")
    ap.add_argument("--limit", type=int, default=0, help="테스트용: 처음 N건만")
    args = ap.parse_args()

    from playwright.async_api import async_playwright

    urls = collect_urls(args.workbook)
    done = load_done(args.out)
    todo = [(u, g) for u, g in urls.items() if u not in done]
    if args.limit:
        todo = todo[:args.limit]
    print(f"전체 {len(urls):,} / 완료 {len(done):,} / 이번 작업 {len(todo):,}건")
    if not todo:
        print("할 일이 없습니다. (모두 완료)"); return

    queue = asyncio.Queue()
    for item in todo:
        queue.put_nowait(item)
    results, lock = [], asyncio.Lock()
    stats = {"done": 0, "empty": 0, "base": len(done), "total": len(urls)}

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        workers = [asyncio.create_task(worker(f"W{i+1}", browser, queue, results, lock, args, stats))
                   for i in range(args.concurrency)]
        await asyncio.gather(*workers)
        if results:
            flush(results, args.out)
        await browser.close()
    print(f"빈 요약(실패/추출불가): {stats['empty']:,}건 — 재실행하면 자동 재시도됩니다.")
    print(f"완료. 결과 저장: {args.out}")
    print("이 파일을 repo에 커밋/푸시하면 워크북에 '사업요약' 열을 합쳐드립니다.")


if __name__ == "__main__":
    # Windows에서 Python 3.8+ 는 기본 Proactor 루프라 별도 정책 설정 불필요
    if sys.platform.startswith("win") and sys.version_info < (3, 8):
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    asyncio.run(main())
