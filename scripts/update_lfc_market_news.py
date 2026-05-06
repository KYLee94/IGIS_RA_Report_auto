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
OUT = ROOT / "docs" / "data" / "lfc-market-news.json"

KEYWORDS = ["부동산 PF", "대출", "건전성", "신용등급", "금융당국", "책임준공", "부동산금융"]
RELEVANCE_TERMS = ["부동산", "PF", "건전성", "신용등급", "금융당국", "책임준공", "부동산금융"]


@dataclass(frozen=True)
class LenderGroup:
    lender: str
    relation: str
    search_names: tuple[str, ...]


LENDERS = [
    LenderGroup("우리은행", "816 Tr.B / 은행권 PF 모니터링", ("우리은행", "우리금융")),
    LenderGroup("메리츠화재", "816 Tr.A-1 / 보험권 대주", ("메리츠화재", "메리츠금융")),
    LenderGroup("한국증권금융", "816 Tr.B / 증권금융 대주", ("한국증권금융",)),
    LenderGroup("저축은행권", "대신저축은행·흥국저축은행", ("대신저축은행", "흥국저축은행", "저축은행")),
    LenderGroup("NH투자증권", "대리금융기관·SPC 관련", ("NH투자증권", "NH금융")),
    LenderGroup("신한투자증권", "427/816 투자자·SPC 관련", ("신한투자증권", "신한금융")),
    LenderGroup("KB국민은행", "통합 PF 후보 주관기관", ("KB국민은행", "KB금융")),
]


def normalize_title(title: str) -> str:
    return " ".join((title or "").replace("\xa0", " ").split())


def google_news_url(query: str) -> str:
    quoted = urllib.parse.quote_plus(query)
    return f"https://news.google.com/rss/search?q={quoted}%20when:3d&hl=ko&gl=KR&ceid=KR:ko"


def fetch_rss(query: str) -> list[dict]:
    req = urllib.request.Request(
        google_news_url(query),
        headers={"User-Agent": "Mozilla/5.0 IOTA-LFC-NewsBot/1.0"},
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
    if not any(name in title for name in group.search_names):
        return False
    compact = title.replace(" ", "")
    return any(term.replace(" ", "") in compact for term in RELEVANCE_TERMS)


def collect() -> dict:
    now = datetime.now(KST)
    cutoff = (now - timedelta(days=3)).date()
    items = []
    for group in LENDERS:
        seen = set()
        articles = []
        for name in group.search_names:
            for keyword in KEYWORDS:
                query = f"{name} {keyword}"
                try:
                    candidates = fetch_rss(query)
                except Exception as exc:
                    print(f"warn: {query}: {exc}", file=sys.stderr)
                    continue
                for article in candidates:
                    if article["date"] < cutoff.isoformat():
                        continue
                    if not related(article, group):
                        continue
                    key = article["title"]
                    if key in seen:
                        continue
                    seen.add(key)
                    article["title"] = f"[{article['date']}] {article['title']}"
                    articles.append(article)
                    if len(articles) >= 3:
                        break
                if len(articles) >= 3:
                    break
                time.sleep(0.2)
            if len(articles) >= 3:
                break
        items.append({"lender": group.lender, "relation": group.relation, "articles": articles[:3]})
    return {"generatedAt": now.isoformat(timespec="seconds"), "windowDays": 3, "items": items}


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    payload = collect()
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
