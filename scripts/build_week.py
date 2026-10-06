"""주간 소재 리포트 데이터 만들기 (구글 시트 → public/data/*.json + public/thumbs/*.jpg)

    python scripts/build_week.py              # 전체 주차 다시 계산
    python scripts/build_week.py --pending    # 태그 안 된 소재 목록 + 원본 이미지 내려받기 (분석용)

시트 '★ 페스룸_메타_클로드'를 서비스 계정 키로 읽습니다. 키 경로는 환경변수 PETHROOM_GSA_KEY,
없으면 기존 동기화 스크립트(pethroom-sync)와 같은 위치를 씁니다.
"""
import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys
import urllib.request
from collections import defaultdict
from pathlib import Path

import gspread

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "public" / "data"
THUMBS = ROOT / "public" / "thumbs"
KEY_PATH = os.environ.get("PETHROOM_GSA_KEY") or str(
    Path.home() / "Downloads" / "pethroom google hub json 키" / "valid-broker-510106-b2-d54daeab26b6.json")
SHEET_ID = "1oYsssIMz-rT8IpNLscMHel_gbUt3I82QJYjHBa-uygs"

MIN_SPEND = 20      # 이 이상 쓴 소재만 '제대로 집행'으로 보고 공통점 집계에 넣음
TAG_SPEND = 5       # 이 이상 쓴 소재는 태그 대상
PRODUCTS = {"EFC": "Facial Comb", "PSS": "Pink Shampoo"}

# 행사 탭: 캠페인 이름에 PBDD가 들어간 광고만 '행사 소재'
EVENT = {
    "key": "pbdd", "name": "PBDD", "title": "Prime Big Deal Days",
    "campaign": re.compile(r"PBDD", re.I),
    "from": "2026-10-03", "to": "2026-10-07",
    "phases": [
        {"key": "pre", "label": "사전", "sub": "잠재고객 모으기", "from": "2026-10-03", "to": "2026-10-05"},
        {"key": "main", "label": "본행사", "sub": "실제 할인 기간", "from": "2026-10-06", "to": "2026-10-07"},
    ],
}

TAG_COLS = ["광고 이름", "제품", "포맷", "이미지 문구", "문구 각도", "모델·견종", "비포애프터", "권위 배지", "가격·할인", "CTA", "메모", "분석일"]
REPORT_COLS = ["주 시작", "한 줄 요약", "잘된 공통점", "안된 공통점", "다음 주 제작 가이드", "랜딩 제안", "작성", "상태"]


def num(v):
    if v is None:
        return None
    s = re.sub(r"[$,%\s]", "", str(v))
    if not s or s.startswith("#") or s in "-–":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def day(v):
    m = re.match(r"(\d{4}-\d{2}-\d{2})", str(v or ""))
    return m.group(1) if m else None


def product_of(name, campaign=""):
    if re.search(r"(^|_)(EFC|IFC)(_|$)", name, re.I) or re.search(r"facial comb", campaign, re.I):
        return "EFC"
    if re.search(r"(^|_)(PSS|SMP)(_|$)", name, re.I) or re.search(r"pink shampoo", campaign, re.I):
        return "PSS"
    return "기타"


def slug(name):
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_").lower() or "ad"


def monday(d):
    d = dt.date.fromisoformat(d)
    return d - dt.timedelta(days=d.weekday())


def open_sheet():
    return gspread.service_account(filename=KEY_PATH).open_by_key(SHEET_ID)


def tab(sh, title, header=None):
    try:
        return sh.worksheet(title).get_all_values()
    except gspread.WorksheetNotFound:
        if header is None:
            return []
        ws = sh.add_worksheet(title, rows=500, cols=len(header))
        ws.update([header], "A1")
        ws.freeze(rows=1)
        return [header]


def load(sh):
    board = sh.worksheet("소재보드").get_all_values()
    master = sh.worksheet("소재마스터").get_all_values()
    meta = sh.worksheet("RAW_meta").get_all_values()
    amazon = sh.worksheet("RAW_amazon").get_all_values()
    tags = tab(sh, "소재태그", TAG_COLS)
    reports = tab(sh, "주간리포트", REPORT_COLS)
    return board, master, meta, amazon, tags, reports


def ads_from(board, master, meta, amazon, tags):
    ads = {}

    def ad(name):
        if name not in ads:
            ads[name] = {"name": name, "daily": defaultdict(lambda: [0.0] * 5), "amz": defaultdict(lambda: [0.0] * 5),
                         "campaign": "", "image": "", "live": None, "verdict": "", "why": "", "status": "",
                         "total_spend": None, "total_roas": None, "tags": {}}
        return ads[name]

    for r in master[1:]:
        if not r or not r[0].strip() or r[0].startswith("[중복"):
            continue
        a = ad(r[0].strip())
        a["live"] = day(r[5] if len(r) > 5 else "")
        a["campaign"] = r[11] if len(r) > 11 else ""
        a["image"] = r[13] if len(r) > 13 else ""

    for r in board[7:]:
        if not r or not r[0].strip() or not (len(r) > 2 and r[2].strip()):
            continue
        a = ad(r[0].strip())
        g = lambda i: r[i] if len(r) > i else ""
        a.update(status=g(2), total_spend=num(g(3)), total_roas=num(g(22)) if num(g(13)) else None,
                 verdict=g(31), why=g(32))

    # RAW_meta: 날짜 캠페인 광고세트 광고ID 광고이름 지출 노출 도달 링크클릭 LPV 리드
    for r in meta[1:]:
        d = day(r[0] if r else "")
        if not d or len(r) < 5 or not r[4].strip():
            continue
        a = ad(r[4].strip())
        a["campaign"] = a["campaign"] or r[1]
        if EVENT["campaign"].search(r[1]):
            a["event_campaign"] = r[1]
        v = a["daily"][d]
        for i, c in enumerate((5, 6, 8, 9, 10)):  # 지출 노출 클릭 LPV 리드
            v[i] += num(r[c] if len(r) > c else None) or 0

    # RAW_amazon: 날짜 캠페인 광고그룹 퍼블리셔 클릭 DPV ATC 구매 수량 매출 광고이름
    for r in amazon[1:]:
        d = day(r[0] if r else "")
        name = r[10].strip() if len(r) > 10 else ""
        if not d or not name:
            continue
        v = ad(name)["amz"][d]
        for i, c in enumerate((4, 5, 6, 7, 9)):  # 클릭 DPV ATC 구매 매출
            v[i] += num(r[c] if len(r) > c else None) or 0

    head = tags[0] if tags else TAG_COLS
    for r in tags[1:]:
        if r and r[0].strip() and r[0].strip() in ads:
            ads[r[0].strip()]["tags"] = {h: (r[i].strip() if i < len(r) else "") for i, h in enumerate(head[1:], 1)}

    for a in ads.values():
        a["product"] = product_of(a["name"], a["campaign"])
        m = re.match(r"(\d{4})(\d{2})(\d{2})", a["name"])
        spent = sorted(d for d, v in a["daily"].items() if v[1] > 0)
        a["launch"] = a["live"] or (f"{m[1]}-{m[2]}-{m[3]}" if m else None) or (spent[0] if spent else None)
    return ads


def sums(series, d0, d1, n):
    out = [0.0] * n
    for d, v in series.items():
        if d0 <= d <= d1:
            for i in range(n):
                out[i] += v[i]
    return out


def kpi(spend, impr, clicks, lpv, leads, az, dpv, atc, orders, sales, amz_ok):
    div = lambda a, b: (a / b) if b else None
    return {
        "spend": round(spend, 2), "impr": int(impr), "clicks": int(clicks), "leads": int(leads),
        "cpl": div(spend, leads), "ctr": div(clicks, impr), "lpvRate": div(leads, lpv),
        "amzClicks": int(az), "atc": int(atc), "orders": int(orders), "sales": round(sales, 2),
        "atcRate": div(atc, az) if amz_ok else None, "roas": div(sales, spend) if amz_ok else None,
    }


def ensure_thumb(a):
    """썸네일(가로 240px)을 public/thumbs에 저장. 메타 이미지 주소는 며칠 뒤 만료되므로 처음 볼 때 받아 둠."""
    out = THUMBS / f"{slug(a['name'])}.jpg"
    if out.exists():
        return f"thumbs/{out.name}"
    url = a.get("image") or ""
    if not url.startswith("http"):
        return ""
    THUMBS.mkdir(parents=True, exist_ok=True)
    tmp = out.with_suffix(".src")
    try:
        urllib.request.urlretrieve(url, tmp)
        ps = ROOT / "scripts" / "resize.ps1"
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps),
                        "-In", str(tmp), "-Out", str(out), "-Width", "240"], check=True, capture_output=True)
    except Exception as e:  # 만료된 주소 등 — 썸네일 없이 진행
        print(f"  썸네일 실패 {a['name']}: {e}")
        return ""
    finally:
        tmp.unlink(missing_ok=True)
    return f"thumbs/{out.name}"


def card(a, w, amz_ok):
    t = a["tags"]
    return {
        "name": a["name"], "product": a["product"], "thumb": ensure_thumb(a),
        "verdict": a["verdict"], "why": a["why"], "status": a["status"],
        "format": t.get("포맷", ""), "copy": t.get("이미지 문구", ""), "angle": t.get("문구 각도", ""),
        "model": t.get("모델·견종", ""), "memo": t.get("메모", ""),
        "week": w, "totalSpend": a["total_spend"], "totalRoas": a["total_roas"],
    }


def split_lines(s):
    return [x.strip(" -•\t") for x in (s or "").splitlines() if x.strip(" -•\t")]


def parse_reports(reports):
    """주간리포트 탭 → {키: 행}. 키는 주 시작 날짜, 또는 행사 이름(예: PBDD)."""
    head = reports[0] if reports else REPORT_COLS
    out = {}
    for r in reports[1:]:
        if r and r[0].strip():
            out[day(r[0]) or r[0].strip()] = {h: (r[i] if i < len(r) else "") for i, h in enumerate(head)}
    return out


def report_of(r):
    return {
        "summary": r.get("한 줄 요약", ""), "good": split_lines(r.get("잘된 공통점")),
        "bad": split_lines(r.get("안된 공통점")), "guide": split_lines(r.get("다음 주 제작 가이드")),
        "landing": split_lines(r.get("랜딩 제안")), "author": r.get("작성", ""), "status": r.get("상태", ""),
    }


def build_event(ads, reports, meta_to, amz_to):
    """행사 탭 데이터: 단계별(사전/본행사) 행사 소재 성과 + 같은 기간 상시 소재와 비교."""
    E = EVENT
    ev = [a for a in ads.values() if a.get("event_campaign")]
    base = [a for a in ads.values() if not a.get("event_campaign")]

    def total(group, d0, d1, amz_ok):
        t = [0.0] * 10
        for a in group:
            for i, x in enumerate(sums(a["daily"], d0, d1, 5) + sums(a["amz"], d0, d1, 5)):
                t[i] += x
        return kpi(*t, amz_ok)

    def ctype(c):
        return "메인행사" if "메인" in c else "장바구니" if "장바구니" in c else "잠재고객" if "잠재" in c else "트래픽"

    phases = []
    for p in E["phases"]:
        s, e = p["from"], p["to"]
        started = meta_to >= s
        amz_cov = "full" if amz_to >= e else ("partial" if amz_to >= s else "none")
        amz_ok = amz_cov != "none"
        items = []
        for a in ev:
            m = sums(a["daily"], s, e, 5)
            z = sums(a["amz"], s, e, 5)
            if m[0] < 1:
                continue
            k = kpi(*m, *z, amz_ok)
            c = card(a, k, amz_ok)
            c["ctype"] = ctype(a["event_campaign"])
            items.append(c)
        if p["key"] == "pre":  # 사전: 리드 모으기가 목표 → 리드 많은 순
            items.sort(key=lambda c: (-c["week"]["leads"], c["week"]["cpl"] or 99))
        else:
            items.sort(key=lambda c: (-c["week"]["sales"], -c["week"]["atc"], -c["week"]["clicks"]))
        phases.append({
            **p, "started": started, "dataTo": min(meta_to, e) if started else "", "amzCoverage": amz_cov,
            "event": total(ev, s, e, amz_ok), "base": total(base, s, e, amz_ok),
            "creatives": items[:8], "more": max(0, len(items) - 8), "count": len(items),
        })

    days = []
    d = dt.date.fromisoformat(E["from"])
    while d.isoformat() <= E["to"]:
        x = d.isoformat()
        days.append({"date": x, "hasMeta": x <= meta_to, "hasAmz": x <= amz_to,
                     "event": total(ev, x, x, x <= amz_to), "base": total(base, x, x, x <= amz_to)})
        d += dt.timedelta(days=1)

    out = {"key": E["key"], "name": E["name"], "title": E["title"], "from": E["from"], "to": E["to"],
           "metaTo": meta_to, "amzTo": amz_to, "phases": phases, "days": days,
           "report": report_of(parse_reports(reports).get(E["name"], {}))}
    (DATA / f"{E['key']}.json").write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"{E['name']}: 행사 소재 {len(ev)}개 · " + " / ".join(f"{p['label']} {p['count']}개 ${p['event']['spend']:,.0f}" for p in phases))
    return {"key": E["key"], "name": E["name"], "from": E["from"], "to": E["to"]}


def build(ads, reports, meta_to, amz_to, event=None):
    first = min((d for a in ads.values() for d in a["daily"]), default=meta_to)
    weeks = []
    w0 = monday(first)
    while w0 <= monday(meta_to):
        weeks.append(w0)
        w0 += dt.timedelta(days=7)

    rep = parse_reports(reports)

    index = []
    prev_k = None
    for w in weeks:
        s, e = w.isoformat(), (w + dt.timedelta(days=6)).isoformat()
        partial = e > meta_to
        amz_cov = "full" if amz_to >= e else ("partial" if amz_to >= s else "none")
        amz_ok = amz_cov != "none"

        rows, tot = [], [0.0] * 10
        for a in ads.values():
            m = sums(a["daily"], s, e, 5)
            z = sums(a["amz"], s, e, 5)
            if m[0] <= 0 and z[4] <= 0:
                continue
            k = kpi(m[0], m[1], m[2], m[3], m[4], z[0], z[1], z[2], z[3], z[4], amz_ok)
            rows.append((a, k))
            for i, v in enumerate(m + z):
                tot[i] += v
        K = kpi(*tot, amz_ok)

        uploaded = [a for a in ads.values() if a["launch"] and s <= a["launch"] <= e]
        K["uploads"] = len(uploaded)
        K["launched"] = sum(1 for a in uploaded if sum(v[0] for v in a["daily"].values()) > 0)
        K["active"] = len(rows)

        by_p = defaultdict(lambda: [0.0] * 10)
        for a, k in rows:
            p = by_p[a["product"]]
            for i, f in enumerate(("spend", "impr", "clicks", None, "leads", "amzClicks", None, "atc", "orders", "sales")):
                if f:
                    p[i] += k[f] or 0
        products = [{"key": p, "label": PRODUCTS.get(p, p), **kpi(*v, amz_ok)} for p, v in by_p.items()]
        products.sort(key=lambda x: -x["spend"])

        real = [(a, k) for a, k in rows if k["spend"] >= MIN_SPEND]
        if amz_ok:
            winners = sorted((x for x in real if (x[1]["roas"] or 0) >= 0.5 and x[1]["orders"] >= 2),
                             key=lambda x: -(x[1]["sales"]))[:5]
            losers = sorted((x for x in real if (x[1]["roas"] or 0) < 0.25), key=lambda x: -x[1]["spend"])[:5]
        else:  # 아마존 데이터 전: CPL로만 1차 판단
            winners = sorted((x for x in real if x[1]["cpl"] and x[1]["cpl"] <= 0.8), key=lambda x: x[1]["cpl"])[:5]
            losers = sorted((x for x in real if not x[1]["cpl"] or x[1]["cpl"] > 1.0), key=lambda x: -x[1]["spend"])[:5]

        # 공통점 집계: 시작~이번 주 끝까지 누적, 제대로 집행된 소재만
        cum = []
        for a in ads.values():
            m = sums(a["daily"], "0000", e, 5)
            z = sums(a["amz"], "0000", e, 5)
            if m[0] >= MIN_SPEND:
                cum.append((a, m, z))
        tag_stats = {}
        for dim in ("문구 각도", "포맷", "모델·견종"):
            g = defaultdict(lambda: [0, [0.0] * 10])
            for a, m, z in cum:
                v = a["tags"].get(dim) or "미분류"
                key = (a["product"], v)
                g[key][0] += 1
                for i, x in enumerate(m + z):
                    g[key][1][i] += x
            tag_stats[dim] = sorted(
                ({"product": p, "value": v, "n": n, **kpi(*t, amz_to >= s)} for (p, v), (n, t) in g.items()),
                key=lambda x: (x["product"], -(x["roas"] or 0), -x["spend"]))
        starved = sum(1 for a in ads.values()
                      if a["launch"] and a["launch"] <= e and 0 < sum(v[0] for v in a["daily"].values()) < MIN_SPEND)

        r = rep.get(s, {})
        out = {
            "report": report_of(r),
            "start": s, "end": e, "partial": partial, "metaTo": meta_to, "amzTo": amz_to, "amzCoverage": amz_cov,
            "kpi": K, "prev": prev_k, "products": products,
            "winners": [card(a, k, amz_ok) for a, k in winners],
            "losers": [card(a, k, amz_ok) for a, k in losers],
            "uploaded": [{"name": a["name"], "product": a["product"], "thumb": ensure_thumb(a),
                          "spend": round(sum(v[0] for d, v in a["daily"].items() if s <= d <= e), 2)}
                         for a in sorted(uploaded, key=lambda a: a["name"])],
            "tagStats": tag_stats, "starved": starved, "minSpend": MIN_SPEND,
        }
        (DATA / f"week-{s}.json").write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        index.append({"start": s, "end": e, "partial": partial, "hasReport": bool(out["report"]["summary"])})
        prev_k = K
        print(f"{s}~{e}  업로드 {K['uploads']} · 지출 ${K['spend']:,.0f} · 위너 {len(winners)} · 리포트 {'O' if r else '-'}")

    kst = dt.datetime.now(dt.timezone.utc) + dt.timedelta(hours=9)
    (DATA / "index.json").write_text(json.dumps({
        "weeks": index[::-1], "event": event, "metaTo": meta_to, "amzTo": amz_to, "builtAt": kst.strftime("%Y-%m-%d %H:%M"),
    }, ensure_ascii=False, indent=1), encoding="utf-8")


def pending(ads):
    """태그가 없는 소재(누적 지출 TAG_SPEND 이상) 원본 이미지를 받아 목록으로 출력."""
    out = Path(os.environ.get("TEMP", "/tmp")) / "pethroom_tag"
    out.mkdir(exist_ok=True)
    todo = []
    for a in sorted(ads.values(), key=lambda a: a["name"]):
        spend = sum(v[0] for v in a["daily"].values())
        if a["tags"] or spend < TAG_SPEND:
            continue
        f = out / f"{slug(a['name'])}.jpg"
        if not f.exists() and a["image"].startswith("http"):
            try:
                urllib.request.urlretrieve(a["image"], f)
            except Exception as e:
                print(f"  이미지 실패 {a['name']}: {e}")
        todo.append({"name": a["name"], "product": a["product"], "spend": round(spend, 2),
                     "file": str(f) if f.exists() else "", "video": "/t15." in a["image"] or "_VD_" in a["name"]})
    print(json.dumps(todo, ensure_ascii=False, indent=1))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pending", action="store_true")
    args = ap.parse_args()
    sh = open_sheet()
    board, master, meta, amazon, tags, reports = load(sh)
    ads = ads_from(board, master, meta, amazon, tags)
    if args.pending:
        return pending(ads)
    meta_to = day(board[2][1]) if len(board) > 2 and len(board[2]) > 1 else None
    meta_to = meta_to or max(d for a in ads.values() for d in a["daily"])
    amz_to = day(board[4][1]) if len(board) > 4 and len(board[4]) > 1 else ""
    amz_to = amz_to or max((d for a in ads.values() for d in a["amz"]), default="")
    DATA.mkdir(parents=True, exist_ok=True)
    event = build_event(ads, reports, meta_to, amz_to)
    build(ads, reports, meta_to, amz_to, event)


if __name__ == "__main__":
    main()
