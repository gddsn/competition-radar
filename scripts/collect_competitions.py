from __future__ import annotations

import logging
from typing import Dict, List, Sequence
from urllib.parse import urldefrag, urljoin, urlparse

from bs4 import BeautifulSoup

from extract_fields import enrich_row
from utils import (
    LATEST_CSV,
    SEED_CSV,
    SOURCES_YML,
    configure_logging,
    discover_candidates,
    fetch_url,
    html_to_text,
    load_yaml,
    merge_rows,
    read_csv,
    save_sqlite,
    write_csv,
)

NOTICE_LINK_KEYWORDS = [
    "通知",
    "公告",
    "报名",
    "赛程",
    "时间安排",
    "参赛指南",
    "参赛",
    "初赛",
    "复赛",
    "决赛",
    "提交",
    "公示",
]


def _as_list(value: object) -> List[str]:
    if isinstance(value, list):
        return [str(item) for item in value]
    if value is None:
        return []
    return [str(value)]


def _same_site(url: str, base_url: str) -> bool:
    parsed = urlparse(url)
    base = urlparse(base_url)
    return parsed.scheme in {"http", "https"} and parsed.netloc == base.netloc


def _notice_links(html: str, base_url: str, limit: int = 5) -> List[str]:
    soup = BeautifulSoup(html or "", "html.parser")
    links: List[str] = []
    seen: set[str] = set()
    for anchor in soup.find_all("a"):
        text = str(anchor.get_text(" ", strip=True) or "")
        href = str(anchor.get("href", "") or "")
        full_url = urldefrag(urljoin(base_url, href)).url
        haystack = f"{text} {full_url}"
        if not _same_site(full_url, base_url):
            continue
        if not any(keyword in haystack for keyword in NOTICE_LINK_KEYWORDS):
            continue
        if full_url in seen:
            continue
        seen.add(full_url)
        links.append(full_url)
        if len(links) >= limit:
            break
    return links


def enrich_seed_rows(seed_rows: Sequence[Dict[str, str]], timeout: int = 12, detail_limit: int = 5) -> List[Dict[str, str]]:
    enriched_rows: List[Dict[str, str]] = []
    for row in seed_rows:
        official_url = row.get("官方链接", "").strip()
        if not official_url.startswith(("http://", "https://")):
            enriched_rows.append(row)
            continue

        html, error = fetch_url(official_url, timeout=timeout)
        if error:
            logging.info("种子官网跳过：%s | %s", official_url, error)
            enriched_rows.append(row)
            continue

        texts = [html_to_text(html)[:3500]]
        for detail_url in _notice_links(html, official_url, limit=detail_limit):
            detail_html, detail_error = fetch_url(detail_url, timeout=timeout)
            if detail_error:
                logging.info("种子通知页跳过：%s | %s", detail_url, detail_error)
                continue
            texts.append(html_to_text(detail_html)[:3500])
        enriched_rows.append(enrich_row(row, f"{row.get('比赛名称', '')} " + " ".join(texts)))
    return enriched_rows


def collect_from_sources(config: Dict[str, object]) -> List[Dict[str, str]]:
    defaults = config.get("defaults", {}) or {}
    global_keywords = _as_list(config.get("global_keywords"))
    timeout = int(defaults.get("timeout_seconds", 12))
    max_items = int(defaults.get("max_items_per_source", 8))
    fetch_detail_pages = bool(defaults.get("fetch_detail_pages", True))

    rows: List[Dict[str, str]] = []
    sources: Sequence[Dict[str, object]] = config.get("sources", []) or []
    for source in sources:
        if not source.get("enabled", True):
            continue
        name = str(source.get("name", "未命名来源"))
        url = str(source.get("url", "")).strip()
        category = str(source.get("category", "未分类"))
        keywords = list(dict.fromkeys(global_keywords + _as_list(source.get("keywords"))))
        if not url:
            logging.warning("跳过没有 URL 的来源：%s", name)
            continue

        logging.info("采集来源：%s", name)
        html, error = fetch_url(url, timeout=timeout)
        if error:
            logging.warning("来源访问失败：%s | %s", name, error)
            continue

        candidates = discover_candidates(
            html=html,
            base_url=url,
            source_name=name,
            category=category,
            keywords=keywords,
            max_items=max_items,
        )

        page_text = html_to_text(html)
        for candidate in candidates:
            detail_text = page_text[:2500]
            detail_url = candidate.get("官方链接", "")
            if fetch_detail_pages and detail_url.startswith(("http://", "https://")) and detail_url != url:
                detail_html, detail_error = fetch_url(detail_url, timeout=timeout)
                if detail_error:
                    logging.info("详情页跳过：%s | %s", detail_url, detail_error)
                elif detail_html:
                    detail_text = html_to_text(detail_html)
            rows.append(enrich_row(candidate, f"{candidate.get('比赛名称', '')} {detail_text}"))
    return rows


def main() -> None:
    configure_logging()
    config = load_yaml(SOURCES_YML)
    defaults = config.get("defaults", {}) or {}
    seed_rows = enrich_seed_rows(
        read_csv(SEED_CSV),
        timeout=int(defaults.get("timeout_seconds", 12)),
        detail_limit=int(defaults.get("max_seed_detail_pages", 5)),
    )
    collected_rows = collect_from_sources(config)
    rows = merge_rows(seed_rows, collected_rows)
    write_csv(rows, LATEST_CSV)
    save_sqlite(rows)
    logging.info("完成采集：种子 %s 条，网页新增 %s 条，合并后 %s 条", len(seed_rows), len(collected_rows), len(rows))


if __name__ == "__main__":
    main()
