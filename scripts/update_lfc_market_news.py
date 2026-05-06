from __future__ import annotations

import email.utils
import json
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path


KST = timezone(timedelta(hours=9))
ROOT = Path(__file__).resolve().parents[1]
LENDER_JSON = ROOT / "docs" / "data" / "lfc-lenders.json"
OUT = ROOT / "docs" / "data" / "lfc-market-news.json"

KEYWORDS = ["부동산 PF", "대출", "건전성", "신용등급", "금융당국", "책임준공", "부동산금융"]
RELEVANCE_TERMS = ["부동산", "PF", "대출", "건전성", "신용등급", "금융당국", "책임준공", "부동산금융"]
GLOBAL_FALLBACK_QUERIES = [
    "부동산 PF 대출 건전성",
    "부동산 PF 금융당국",
    "부동산금융 신용등급",
]


@dataclass(frozen=True)
class LenderGroup:
    lender: str
    relation: str
    search_names: tuple[str, ...]


FALLBACK_LENDERS = [
    LenderGroup("메리츠증권", "816 Tr.A-1 SPC 관련 모니터링", ("메리츠증권", "메리츠화재")),
    LenderGroup("NH투자증권", "816 Tr.A-2 SPC 및 대리금융 관련 모니터링", ("NH투자증권", "NH금융")),
    LenderGroup("신한투자증권", "816공간제일차 SPC 관련 모니터링", ("신한투자증권", "신한금융")),
    LenderGroup("대신증권", "이터널하이브 SPC 관련 모니터링", ("대신증권", "대신저축은행")),
    LenderGroup("KB국민은행", "본PF 후보 주관기관 모니터링", ("KB국민은행", "KB금융")),
]


def normalize_title(title: str) -> str:
    return " ".join((title or "").replace("\xa0", " ").split())


def load_lenders() -> list[LenderGroup]:
    if not LENDER_JSON.exists():
        return FALLBACK_LENDERS
    try:
        payload = json.loads(LENDER_JSON.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"warn: cannot read {LENDER_JSON}: {exc}", file=sys.stderr)
        return FALLBACK_LENDERS
    groups = []
    for row in payload.get("lenders", []):
        lender = str(row.get("lender") or "").strip()
        if not lender:
            continue
        relation = str(row.get("relation") or "").strip()
        names = tuple(str(v).strip() for v in row.get("searchNames", []) if str(v).strip())
        groups.append(LenderGroup(lender, relation, names or (lender,)))
    return groups or FALLBACK_LENDERS


def google_news_url(query: str) -> str:
    quoted = urllib.parse.quote_plus(query)
    return f"https://news.google.com/rss/search?q={quoted}%20when:3d&hl=ko&gl=KR&ceid=KR:ko"


def fetch_rss(query: str) -> list[dict]:
    req = urllib.request.Request(
        google_news_url(query),
        headers={"User-Agent": "Mozilla/5.0 IOTA-LFC-NewsBot/1.1"},
    )
    with urllib.request.urlopen(req, timeout=25) as resp:
        data = resp.read()
    root = ET.fromstring(data)
    out = []
    for item in root.findall(".//item"):
        title = normalize_title(item.findtext("title") or "")
        link = item.findtext("link") or ""
        pub = item.findtext("pubDate") or ""
        source = item.findtext("source") or ""
        try:
            dt = email.utils.parsedate_to_datetime(pub).astimezone(KST)
        except Exception:
            continue
        out.append({"date": dt.date().isoformat(), "title": title, "publisher": source, "url": link})
    return out


def related(article: dict, group: LenderGroup) -> bool:
    title = article["title"]
    compact = title.replace(" ", "")
    if not any(name and name.replace(" ", "") in compact for name in group.search_names):
        return False
    return any(term.replace(" ", "") in compact for term in RELEVANCE_TERMS)


def market_related(article: dict) -> bool:
    compact = article["title"].replace(" ", "")
    return any(term.replace(" ", "") in compact for term in RELEVANCE_TERMS)


def sector_queries(group: LenderGroup) -> list[str]:
    text = " ".join((group.lender, group.relation, *group.search_names))
    queries = []
    if "저축은행" in text:
        queries += ["저축은행 부동산 PF 건전성", "저축은행 부동산 PF 신용등급"]
    if "증권" in text or "투자증권" in text:
        queries += ["증권사 부동산 PF 대출", "증권사 부동산 PF 건전성"]
    if "자산운용" in text or "운용" in text or "사모부동산투자신탁" in text:
        queries += ["자산운용 부동산 PF", "부동산펀드 PF 대출"]
    if "화재" in text or "보험" in text:
        queries += ["보험사 부동산 PF 대출", "보험사 대체투자 건전성"]
    if "카드" in text:
        queries += ["카드사 부동산 PF 대출", "여전사 부동산 PF"]
    if "은행" in text and "저축은행" not in text:
        queries += ["은행 부동산 PF 대출", "은행 부동산PF 건전성"]
    if "소노" in text:
        queries += ["호텔 개발 부동산 PF", "부동산 PF 대출"]
    queries += GLOBAL_FALLBACK_QUERIES
    deduped = []
    for query in queries:
        if query not in deduped:
            deduped.append(query)
    return deduped


def append_articles(
    articles: list[dict],
    seen: set[str],
    candidates: list[dict],
    cutoff: str,
    accept,
    match_mode: str,
    limit: int = 3,
) -> None:
    for article in candidates:
        if article["date"] < cutoff:
            continue
        if not accept(article):
            continue
        key = article["title"]
        if key in seen:
            continue
        seen.add(key)
        copied = dict(article)
        copied["title"] = f"[{copied['date']}] {copied['title']}"
        copied["matchMode"] = match_mode
        articles.append(copied)
        if len(articles) >= limit:
            break


def collect_group_articles(group: LenderGroup, cutoff: str, limit: int = 3) -> list[dict]:
    seen: set[str] = set()
    articles: list[dict] = []

    for name in group.search_names:
        for keyword in KEYWORDS:
            query = f"{name} {keyword}"
            try:
                candidates = fetch_rss(query)
            except Exception as exc:
                print(f"warn: {query}: {exc}", file=sys.stderr)
                continue
            append_articles(articles, seen, candidates, cutoff, lambda a, g=group: related(a, g), "대주명+PF키워드", limit)
            if len(articles) >= limit:
                return articles[:limit]
            time.sleep(0.2)

    for name in group.search_names:
        query = name
        try:
            candidates = fetch_rss(query)
        except Exception as exc:
            print(f"warn: {query}: {exc}", file=sys.stderr)
            continue
        append_articles(articles, seen, candidates, cutoff, lambda a, g=group: any(n.replace(" ", "") in a["title"].replace(" ", "") for n in g.search_names), "대주명", limit)
        if len(articles) >= limit:
            return articles[:limit]
        time.sleep(0.2)

    for query in sector_queries(group):
        try:
            candidates = fetch_rss(query)
        except Exception as exc:
            print(f"warn: {query}: {exc}", file=sys.stderr)
            continue
        append_articles(articles, seen, candidates, cutoff, market_related, "업권/시장", limit)
        if len(articles) >= max(1, limit):
            return articles[:limit]
        time.sleep(0.2)

    return articles[:limit]


def collect() -> dict:
    now = datetime.now(KST)
    cutoff = (now - timedelta(days=3)).date().isoformat()
    items = []
    for group in load_lenders():
        articles = collect_group_articles(group, cutoff, 3)
        items.append({"lender": group.lender, "relation": group.relation, "articles": articles[:3]})
    return {"generatedAt": now.isoformat(timespec="seconds"), "windowDays": 3, "items": items}


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    payload = collect()
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
