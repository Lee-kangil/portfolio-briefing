# -*- coding: utf-8 -*-
"""
포트폴리오 / 워치리스트 그룹별 풀 브리핑 (네이버 금융 + 야후 파이낸스 시세 + Google 뉴스)
=====================================================================
보유 포트폴리오 + 동행학교 1·2·3군 그룹을 각각 별도 HTML 리포트로 생성합니다.
국내 종목 시세는 네이버 금융 실시간 API, 미국 종목 시세는 야후 파이낸스 API에서 받아옵니다.
(pykrx 미사용 → KRX 로그인/차단 문제 없음)
[로컬 실행 - Windows]
  pip install requests
  python portfolio_full_briefing.py
  -> Downloads 폴더에 그룹별 portfolio_<그룹>_YYYYMMDD.html / .csv 생성
[클라우드 - GitHub Actions]
  BRIEFING_OUT=docs 지정 시 그 폴더에 g1.html.. + index.html(목차) 생성
* 인터넷 되는 환경에서 실행. (Cowork 샌드박스는 외부망 차단)
종목은 (이름, 코드, 시장) 으로 지정. 시장이 KOSPI/KOSDAQ 면 네이버, US/NASDAQ/NYSE/AMEX 면 야후.
네이버가 돌려주는 종목명과 비교해 코드 오류 시 ⚠ 표시 (미국 종목은 교차검증 생략).
"""
import os
import re
import csv
import sys
import json
import time
import html
import subprocess
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from email.utils import parsedate_to_datetime
import xml.etree.ElementTree as ET
def _ensure(pkg, import_name=None):
    try:
        __import__(import_name or pkg)
    except ImportError:
        print(f"[설치] {pkg} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
_ensure("requests")
import requests
KST = timezone(timedelta(hours=9))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 briefing-bot"
C_UP, C_DOWN, C_FLAT = "#e53e3e", "#2b6cb0", "#718096"
NEWS_PER_STOCK = 2
# 미국 시장(야후 파이낸스 라우팅). 국내는 KOSPI/KOSDAQ.
US_MARKETS = {"US", "NASDAQ", "NYSE", "AMEX"}
def _is_us(market):
    return str(market).upper() in US_MARKETS
# ---------------------------------------------------------------------------
# 그룹 정의 — (이름, 코드/티커, 시장). 종목 추가/삭제는 줄 단위로.
#   국내: 시장 = KOSPI / KOSDAQ, 코드 = 6자리(특수 티커 포함)
#   미국: 시장 = US / NASDAQ / NYSE / AMEX, 코드 = 영문 티커
# ---------------------------------------------------------------------------
GROUPS = {
    "보유 포트폴리오": [
        ("삼성전자", "005930", "KOSPI"), ("SK하이닉스", "000660", "KOSPI"),
        ("DRAM (Roundhill Memory ETF)", "DRAM", "NASDAQ"),
    ],
    "1군 · 글로벌 브랜드": [
        ("삼성물산", "028260", "KOSPI"), ("LG", "003550", "KOSPI"),
        ("LG전자우", "066575", "KOSPI"), ("LG디스플레이", "034220", "KOSPI"),
        ("LG화학우", "051915", "KOSPI"), ("HD현대중공업", "329180", "KOSPI"),
        ("두산우", "000155", "KOSPI"), ("두산에너빌리티", "034020", "KOSPI"),
        ("OCI홀딩스", "010060", "KOSPI"), ("한화솔루션", "009830", "KOSPI"),
        ("SKC", "011790", "KOSPI"), ("롯데에너지머티리얼즈", "020150", "KOSPI"),
    ],
    "2군 · 대기업 집단 계열사": [
        ("한국전력", "015760", "KOSPI"), ("한전기술", "052690", "KOSPI"),
        ("대한항공", "003490", "KOSPI"), ("한화3우B", "00088K", "KOSPI"),
        ("SK이노베이션", "096770", "KOSPI"), ("KCC", "002380", "KOSPI"),
        ("DL", "000210", "KOSPI"), ("DL이앤씨", "375500", "KOSPI"),
        ("GS건설", "006360", "KOSPI"), ("삼성에피스홀딩스", "0126Z0", "KOSPI"),
        ("롯데케미칼", "011170", "KOSPI"), ("HD한국조선해양", "009540", "KOSPI"),
        ("SK가스", "018670", "KOSPI"), ("한국가스공사", "036460", "KOSPI"),
        ("현대제철", "004020", "KOSPI"), ("OCI", "456040", "KOSPI"),
    ],
    "3군 · 아직 시장이 발견하지 못한 기업": [
        ("한온시스템", "018880", "KOSPI"), ("HMM", "011200", "KOSPI"),
        ("LX하우시스", "108670", "KOSPI"), ("LX세미콘", "108320", "KOSDAQ"),
        ("LX홀딩스", "383800", "KOSPI"), ("벽산", "007210", "KOSPI"),
        ("동방", "004140", "KOSPI"), ("아이에스동서", "010780", "KOSPI"),
        ("LX인터내셔널", "001120", "KOSPI"), ("SK오션플랜트", "100090", "KOSDAQ"),
        ("유니드", "014830", "KOSPI"), ("태웅", "044490", "KOSDAQ"),
        ("태광", "023160", "KOSDAQ"), ("성광벤드", "014620", "KOSDAQ"),
        ("하이록코리아", "013030", "KOSDAQ"), ("HDC", "012630", "KOSPI"),
        ("케이에스피", "073010", "KOSDAQ"), ("파미셀", "005690", "KOSPI"),
        ("보령", "003850", "KOSPI"), ("동화기업", "025900", "KOSDAQ"),
    ],
}
# ---------------------------------------------------------------------------
# 시세 공통
# ---------------------------------------------------------------------------
def _find_quote(obj):
    """JSON 안에서 'nv'(현재가)와 'cv'(전일대비)를 가진 dict 를 재귀 탐색."""
    if isinstance(obj, dict):
        if "nv" in obj and "cv" in obj:
            return obj
        for v in obj.values():
            r = _find_quote(v)
            if r:
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = _find_quote(v)
            if r:
                return r
    return None
def _num(x):
    try:
        return float(str(x).replace(",", ""))
    except Exception:
        return None
_H = {"User-Agent": UA, "Referer": "https://finance.naver.com/",
      "Accept": "application/json"}
def _mk(close, diff, pct, name, cur="KRW"):
    """시세 dict 생성. KRW 는 정수 원, USD 는 소수 2자리 유지."""
    if cur == "USD":
        close = round(close, 2)
        diff = round(diff, 2)
    else:
        close = int(round(close))
        diff = int(round(diff))
    return {"close": close, "diff": diff,
            "pct": round(pct, 2), "name": (name or "").strip(), "cur": cur,
            "date": datetime.now(KST).strftime("%Y-%m-%d")}
# ---------------------------------------------------------------------------
# 시세: 네이버 금융 실시간 API (국내)
# ---------------------------------------------------------------------------
def _from_mstock(code):
    # 모바일 네이버 증권 API (가장 안정적)
    r = requests.get(f"https://m.stock.naver.com/api/stock/{code}/basic",
                     timeout=8, headers=_H)
    r.raise_for_status()
    j = r.json()
    close = _num(j.get("closePrice"))
    if close is None:
        return None
    cdir = str((j.get("compareToPreviousPrice") or {}).get("code", ""))
    sign = 1 if cdir in ("1", "2") else (-1 if cdir in ("4", "5") else 0)
    diff = abs(_num(j.get("compareToPreviousClosePrice")) or 0)
    pct = abs(_num(j.get("fluctuationsRatio")) or 0)
    return _mk(close, sign * diff, sign * pct, j.get("stockName"))
def _from_polling(code):
    r = requests.get(
        f"https://polling.finance.naver.com/api/realtime/domestic/stock/{code}",
        timeout=8, headers=_H)
    r.raise_for_status()
    d = _find_quote(r.json())
    if not d:
        return None
    close = _num(d.get("nv"))
    if close is None:
        return None
    rf = str(d.get("rf", "")).strip()
    sign = 1 if rf in ("1", "2") else (-1 if rf in ("4", "5") else 0)
    return _mk(close, sign * abs(_num(d.get("cv")) or 0),
               sign * abs(_num(d.get("cr")) or 0), d.get("nm"))
# ---------------------------------------------------------------------------
# 시세: 야후 파이낸스 chart API (미국)
# ---------------------------------------------------------------------------
def _from_yahoo(ticker):
    """야후 파이낸스 v8 chart API. query1 실패 시 query2 폴백. USD 시세 반환."""
    last = None
    for host in ("query1", "query2"):
        try:
            url = (f"https://{host}.finance.yahoo.com/v8/finance/chart/"
                   f"{ticker}?range=5d&interval=1d")
            r = requests.get(url, timeout=10,
                             headers={"User-Agent": UA, "Accept": "application/json"})
            r.raise_for_status()
            j = r.json()
            res = (((j.get("chart") or {}).get("result") or [None])[0]) or {}
            meta = res.get("meta") or {}
            close = _num(meta.get("regularMarketPrice"))
            prev = _num(meta.get("chartPreviousClose") or meta.get("previousClose"))
            if close is None or prev is None:
                last = "빈응답"
                continue
            diff = close - prev
            pct = (diff / prev * 100) if prev else 0.0
            name = meta.get("shortName") or meta.get("symbol") or ticker
            return _mk(close, diff, pct, name, cur="USD")
        except Exception as e:
            last = f"{type(e).__name__}:{str(e)[:50]}"
    raise RuntimeError(last or "yahoo 실패")
def fetch_price(code, market="KOSPI"):
    """시장에 따라 야후(미국) 또는 네이버(국내) 시세를 시도. (시세 dict, 에러) 반환."""
    if _is_us(market):
        try:
            res = _from_yahoo(code)
            if res:
                return res, None
            return None, "yahoo:빈응답"
        except Exception as e:
            return None, f"yahoo:{type(e).__name__}:{str(e)[:60]}"
    errs = []
    for fn in (_from_mstock, _from_polling):
        try:
            res = fn(code)
            if res:
                return res, None
            errs.append(f"{fn.__name__}:빈응답")
        except Exception as e:
            errs.append(f"{fn.__name__}:{type(e).__name__}:{str(e)[:60]}")
    return None, " | ".join(errs)
# ---------------------------------------------------------------------------
# 뉴스: Google News RSS (국내=한국어, 미국=영어)
# ---------------------------------------------------------------------------
def fetch_news(name, market="KOSPI", n=NEWS_PER_STOCK):
    clean = re.sub(r"\(.*?\)", "", name).strip()  # 괄호 설명 제거
    if _is_us(market):
        q = quote(f"{clean} ETF stock")
        url = f"https://news.google.com/rss/search?q={q}&hl=en-US&gl=US&ceid=US:en"
    else:
        q = quote(f"{clean} 주가")
        url = f"https://news.google.com/rss/search?q={q}&hl=ko&gl=KR&ceid=KR:ko"
    try:
        r = requests.get(url, timeout=10, headers={"User-Agent": UA})
        r.raise_for_status()
        root = ET.fromstring(r.content)
    except Exception as e:
        return [{"title": f"(뉴스 수집 실패: {e})", "link": "", "source": "", "date": ""}]
    items = []
    for it in root.findall(".//item")[:n]:
        title = (it.findtext("title") or "").strip()
        link = (it.findtext("link") or "").strip()
        src_el = it.find("{http://news.google.com}source")
        if src_el is None:
            src_el = it.find("source")
        source = src_el.text.strip() if (src_el is not None and src_el.text) else ""
        if " - " in title:
            head, tail = title.rsplit(" - ", 1)
            title = head
            if not source:
                source = tail
        pub = it.findtext("pubDate") or ""
        try:
            dt = parsedate_to_datetime(pub).astimezone(KST).strftime("%m.%d %H:%M")
        except Exception:
            dt = ""
        items.append({"title": title, "link": link, "source": source, "date": dt})
    return items or [{"title": "(최근 뉴스 없음)", "link": "", "source": "", "date": ""}]
def _norm(s):
    return re.sub(r"\s+", "", s or "")
def _slug(i):
    return f"g{i+1}"
def main():
    now = datetime.now(KST)
    out_dir = os.environ.get("BRIEFING_OUT") or os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(out_dir, exist_ok=True)
    cloud = bool(os.environ.get("BRIEFING_OUT"))
    stamp = now.strftime("%Y%m%d")
    index_links = []
    for gi, (gtitle, members) in enumerate(GROUPS.items()):
        print("=" * 70)
        print(f" [{gtitle}]  종목 {len(members)}개")
        print("=" * 70)
        data, warns = [], []
        for name, code, mkt in members:
            price, err = fetch_price(code, mkt)
            news = fetch_news(name, mkt)
            time.sleep(0.3)
            warn = ""
            if price:
                # 코드↔이름 교차검증 (국내만; 미국은 종목명 표기 차이로 생략)
                if (not _is_us(mkt)) and price["name"] and _norm(price["name"]) != _norm(name):
                    warn = f"코드확인필요(네이버명:{price['name']})"
                    warns.append(f"{name}->{price['name']}")
                ar = "▲" if price["pct"] > 0 else ("▼" if price["pct"] < 0 else "-")
                print(f"  {name:<20} {_fmt_close(price):>12} {ar}{price['pct']:+.2f}%"
                      + (f"  ⚠ {warn}" if warn else ""))
            else:
                print(f"  {name:<20} (시세 실패: {err})")
            data.append({"name": name, "code": code, "market": mkt,
                         "price": price, "news": news, "warn": warn})
        valid = [d for d in data if d["price"]]
        up = sum(1 for d in valid if d["price"]["pct"] > 0)
        down = sum(1 for d in valid if d["price"]["pct"] < 0)
        flat = len(valid) - up - down
        top = max(valid, key=lambda d: abs(d["price"]["pct"])) if valid else None
        if warns:
            print("  [확인필요]", ", ".join(warns))
        safe = re.sub(r"[^0-9A-Za-z가-힣]+", "_", gtitle).strip("_")
        with open(os.path.join(out_dir, f"portfolio_{safe}_{stamp}.csv"),
                  "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["종목명", "코드", "시장", "종가", "등락률(%)", "통화", "기준일", "비고"])
            for d in data:
                p = d["price"]
                w.writerow([d["name"], d["code"], d["market"],
                            p["close"] if p else "", p["pct"] if p else "",
                            p["cur"] if p else "", p["date"] if p else "", d["warn"]])
        dated = os.path.join(out_dir, f"portfolio_{safe}_{stamp}.html")
        _write_html(data, (up, down, flat, top), now, gtitle, dated)
        if cloud:
            fixed = f"{_slug(gi)}.html"
            _write_html(data, (up, down, flat, top), now, gtitle,
                        os.path.join(out_dir, fixed))
            index_links.append((gtitle, fixed, up, down, flat, top))
    if cloud:
        _write_index(index_links, now, os.path.join(out_dir, "index.html"))
    print("\n완료. 저장 위치:", out_dir)
# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------
def _pct_style(pct):
    if pct is None:
        return "color:#718096", "—"
    if pct > 0:
        return f"color:{C_UP};font-weight:700", f"▲ {pct:+.2f}%"
    if pct < 0:
        return f"color:{C_DOWN};font-weight:700", f"▼ {pct:+.2f}%"
    return f"color:{C_FLAT};font-weight:700", f"{pct:+.2f}%"
def _fmt_close(p):
    """통화에 맞춘 현재가 문자열. USD=$X.XX, KRW=X,XXX원."""
    if not p:
        return "—"
    if p.get("cur") == "USD":
        return f"${p['close']:,.2f}"
    return f"{p['close']:,}원"
def _fmt_diff(p):
    """통화에 맞춘 전일대비 문자열 (뒤 공백 포함)."""
    if not p or p.get("diff") is None:
        return ""
    if p.get("cur") == "USD":
        return f"{p['diff']:+,.2f} USD "
    return f"{p['diff']:+,}원 "
CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;background:#f4f6f9;color:#1a202c;padding:0 0 50px;line-height:1.55}
.wrap{max-width:980px;margin:0 auto;padding:0 16px}
header{background:linear-gradient(135deg,#1f2a44,#2c3e50);color:#fff;padding:24px 0}
header h1{font-size:22px;font-weight:800}
header .meta{font-size:13px;color:#c5cfe0;margin-top:5px}
.bar{background:#fff;border-radius:8px;padding:13px 16px;margin:18px 0;font-size:14px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
h2{font-size:18px;margin:26px 0 12px;padding-left:10px;border-left:5px solid #2c3e50}
table{width:100%;border-collapse:collapse;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.06);font-size:13.5px}
th{background:#2c3e50;color:#fff;padding:10px 8px}
td{padding:9px 8px;border-bottom:1px solid #e2e8f0;text-align:center}
td.nm{text-align:left;font-weight:700} td.cd{color:#5a6675;font-size:12px}
td.num{text-align:right;font-variant-numeric:tabular-nums}
tr:last-child td{border-bottom:none}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:14px;margin-top:6px}
.card{background:#fff;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.07);padding:14px 16px;border-top:4px solid #2c3e50}
.ctop{display:flex;justify-content:space-between;align-items:baseline;gap:8px}
.ctop h3{font-size:15px;font-weight:800} .ctop .cd{font-size:11px;color:#5a6675}
.cprice{font-size:20px;font-weight:800;margin-top:4px}
.cchg{font-size:13px;margin-bottom:8px}
.warn{color:#b45309;font-size:11.5px;margin-bottom:6px}
.clabel{display:inline-block;font-size:11px;font-weight:700;color:#fff;background:#2c3e50;border-radius:4px;padding:1px 7px;margin:6px 0 4px}
.news{font-size:12px;margin:5px 0;padding-left:11px;position:relative}
.news::before{content:'·';position:absolute;left:1px;color:#2c3e50;font-weight:800}
.news a{color:#1d4ed8;text-decoration:none} .news a:hover{text-decoration:underline}
.nmeta{display:block;color:#94a3b8;font-size:11px}
.idx a{display:block;background:#fff;border-radius:10px;padding:14px 16px;margin:10px 0;text-decoration:none;color:#1a202c;box-shadow:0 1px 3px rgba(0,0,0,.06)}
.idx a:hover{background:#eef2f7}
footer{text-align:center;font-size:11.5px;color:#94a3b8;margin-top:34px}
"""
# ---------------------------------------------------------------------------
# GitHub Action 수동 실행 버튼 (목차 페이지 상단에 삽입됨)
#  - 토큰은 공개 소스에 안 들어가고, 사용자 브라우저(localStorage)에만 저장됨
#  - 저장소/브랜치가 바뀌면 아래 REPO / BRANCH 만 수정
# ---------------------------------------------------------------------------
RUN_BUTTON_HTML = """
<div id="gh-run-box" style="margin:18px 0;padding:14px 16px;border:1px solid #e2e8f0;border-radius:10px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.06);">
  <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
    <button id="gh-run-btn" style="padding:9px 16px;border:0;border-radius:8px;background:#2563eb;color:#fff;font-size:14px;font-weight:700;cursor:pointer;">▶ 지금 브리핑 새로 실행</button>
    <span id="gh-run-status" style="font-size:13px;color:#5a6675;"></span>
    <button id="gh-token-btn" title="토큰 설정" style="margin-left:auto;padding:6px 10px;border:1px solid #cbd5e0;border-radius:8px;background:#fff;font-size:12px;color:#5a6675;cursor:pointer;">⚙ 토큰</button>
  </div>
  <div id="gh-token-row" style="display:none;margin-top:10px;">
    <input id="gh-token-input" type="password" placeholder="GitHub PAT 붙여넣기 (한 번만)" autocomplete="off" style="width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #cbd5e0;border-radius:8px;font-size:13px;">
    <div style="display:flex;gap:8px;margin-top:8px;">
      <button id="gh-token-save" style="padding:7px 12px;border:0;border-radius:8px;background:#16a34a;color:#fff;font-size:13px;cursor:pointer;">저장</button>
      <button id="gh-token-clear" style="padding:7px 12px;border:1px solid #cbd5e0;border-radius:8px;background:#fff;font-size:13px;color:#5a6675;cursor:pointer;">삭제</button>
      <span style="font-size:11px;color:#94a3b8;align-self:center;">이 브라우저에만 저장됨</span>
    </div>
  </div>
</div>
<script>
(function () {
  var REPO = "Lee-kangil/portfolio-briefing";
  var WORKFLOW = "briefing.yml";
  var BRANCH = "main";
  var KEY = "gh_pat_portfolio";
  var $ = function (id) { return document.getElementById(id); };
  var statusEl = $("gh-run-status");
  function setStatus(t, c) { statusEl.textContent = t; statusEl.style.color = c || "#5a6675"; }
  $("gh-token-btn").onclick = function () {
    var row = $("gh-token-row");
    row.style.display = (row.style.display === "none") ? "block" : "none";
    if (localStorage.getItem(KEY)) $("gh-token-input").placeholder = "토큰 저장됨 — 새로 바꾸려면 붙여넣기";
  };
  $("gh-token-save").onclick = function () {
    var v = $("gh-token-input").value.trim();
    if (!v) { setStatus("토큰이 비어 있습니다.", "#dc2626"); return; }
    localStorage.setItem(KEY, v);
    $("gh-token-input").value = "";
    $("gh-token-row").style.display = "none";
    setStatus("토큰 저장 완료 ✓", "#16a34a");
  };
  $("gh-token-clear").onclick = function () {
    localStorage.removeItem(KEY);
    setStatus("토큰 삭제됨", "#5a6675");
  };
  $("gh-run-btn").onclick = async function () {
    var token = localStorage.getItem(KEY);
    if (!token) { setStatus("먼저 ⚙토큰을 등록하세요.", "#dc2626"); $("gh-token-row").style.display = "block"; return; }
    this.disabled = true;
    setStatus("실행 요청 중…", "#2563eb");
    try {
      var res = await fetch(
        "https://api.github.com/repos/" + REPO + "/actions/workflows/" + WORKFLOW + "/dispatches",
        { method: "POST",
          headers: { "Accept": "application/vnd.github+json", "Authorization": "Bearer " + token, "X-GitHub-Api-Version": "2022-11-28" },
          body: JSON.stringify({ ref: BRANCH }) }
      );
      if (res.status === 204) {
        setStatus("실행 시작됨 ✓ 1~2분 뒤 새로고침하세요.", "#16a34a");
      } else if (res.status === 401 || res.status === 403) {
        setStatus("토큰 권한 오류(" + res.status + "). Actions 쓰기 권한 확인.", "#dc2626");
      } else if (res.status === 404) {
        setStatus("저장소/워크플로/브랜치 이름 확인 필요(404).", "#dc2626");
      } else {
        var msg = "";
        try { msg = (await res.json()).message || ""; } catch (e) {}
        setStatus("실패(" + res.status + ") " + msg, "#dc2626");
      }
    } catch (e) {
      setStatus("네트워크 오류: " + e.message, "#dc2626");
    } finally {
      this.disabled = false;
    }
  };
})();
</script>
"""
def _write_html(data, summary, now, group_title, path):
    up, down, flat, top = summary
    rows = []
    for d in data:
        p = d["price"]
        style, txt = _pct_style(p["pct"] if p else None)
        close = _fmt_close(p)
        rows.append(
            "<tr><td class='nm'>" + html.escape(d["name"]) + "</td>"
            "<td class='cd'>" + html.escape(str(d["code"])) + "</td><td>" + d["market"] + "</td>"
            "<td class='num'>" + close + "</td>"
            "<td class='num' style='" + style + "'>" + txt + "</td>"
            "<td class='cd'>" + (p["date"] if p else "—") + "</td></tr>")
    table = "".join(rows)
    cards = []
    for d in data:
        p = d["price"]
        style, txt = _pct_style(p["pct"] if p else None)
        price_txt = _fmt_close(p) if p else "시세 없음"
        diff_txt = _fmt_diff(p)
        warn_html = ("<div class='warn'>⚠ " + html.escape(d["warn"]) + "</div>") if d.get("warn") else ""
        news_html = ""
        for nw in d["news"]:
            t = html.escape(nw["title"])
            meta = " · ".join(x for x in (nw["source"], nw["date"]) if x)
            if nw["link"]:
                news_html += ("<div class='news'><a href='" + html.escape(nw["link"])
                              + "' target='_blank'>" + t + "</a><span class='nmeta'>"
                              + html.escape(meta) + "</span></div>")
            else:
                news_html += "<div class='news'>" + t + "</div>"
        cards.append(
            "<div class='card'><div class='ctop'><h3>" + html.escape(d["name"])
            + "</h3><span class='cd'>" + html.escape(str(d["code"])) + " · " + d["market"]
            + "</span></div><div class='cprice'>" + price_txt
            + "</div><div class='cchg' style='" + style + "'>" + diff_txt + txt
            + "</div>" + warn_html + "<div class='clabel'>최신 뉴스</div>" + news_html + "</div>")
    cards_html = "".join(cards)
    top_html = ""
    if top:
        ts, tt = _pct_style(top["price"]["pct"])
        top_html = ("오늘의 주목 <b>" + html.escape(top["name"]) + "</b> "
                    "<span style='" + ts + "'>" + tt + "</span>")
    gen = now.strftime("%Y-%m-%d %H:%M KST")
    gt = html.escape(group_title)
    doc = (
        "<!DOCTYPE html><html lang='ko'><head><meta charset='UTF-8'>"
        "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
        "<title>" + gt + " · " + now.strftime("%Y.%m.%d") + "</title><style>" + CSS
        + "</style></head><body><header><div class='wrap'>"
        "<h1>📊 " + gt + "</h1><div class='meta'>생성 " + gen
        + " · 시세: 네이버 금융/야후 파이낸스 · 뉴스: Google News</div>"
        "</div></header><div class='wrap'>"
        "<div class='bar'>상승 <b style='color:" + C_UP + "'>" + str(up)
        + "</b> · 하락 <b style='color:" + C_DOWN + "'>" + str(down)
        + "</b> · 보합 " + str(flat) + " &nbsp;|&nbsp; " + top_html + "</div>"
        "<h2>요약</h2><table><thead><tr>"
        "<th>종목명</th><th>코드</th><th>시장</th><th>현재가</th><th>등락률</th><th>기준</th>"
        "</tr></thead><tbody>" + table + "</tbody></table>"
        "<h2>종목별 상세 (시세 + 뉴스)</h2><div class='cards'>" + cards_html + "</div>"
        "<footer>시세: 네이버 금융/야후 파이낸스 · 뉴스: Google News RSS · 정보 제공 목적이며 투자 권유가 아닙니다.</footer>"
        "</div></body></html>"
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
def _write_index(links, now, path):
    gen = now.strftime("%Y-%m-%d %H:%M KST")
    items = ""
    for (title, href, up, down, flat, top) in links:
        ts, tt = _pct_style(top["price"]["pct"]) if top else ("", "—")
        items += ("<a href='" + href + "'><b>" + html.escape(title) + "</b><br>"
                  "<span style='font-size:13px;color:#5a6675'>상승 " + str(up)
                  + " · 하락 " + str(down) + " · 보합 " + str(flat)
                  + " | 주목 " + (html.escape(top['name']) if top else '-') + " "
                  "<span style='" + ts + "'>" + tt + "</span></span></a>")
    doc = ("<!DOCTYPE html><html lang='ko'><head><meta charset='UTF-8'>"
           "<meta name='viewport' content='width=device-width, initial-scale=1.0'>"
           "<title>포트폴리오 브리핑 목차 · " + now.strftime("%Y.%m.%d") + "</title><style>"
           + CSS + "</style></head><body><header><div class='wrap'>"
           "<h1>📊 포트폴리오 브리핑 — 그룹별</h1><div class='meta'>생성 " + gen
           + "</div></div></header><div class='wrap idx'>" + RUN_BUTTON_HTML + items
           + "<footer>시세: 네이버 금융/야후 파이낸스 · 뉴스: Google News · 투자 권유가 아닙니다.</footer>"
           "</div></body></html>")
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)
if __name__ == "__main__":
    main()
