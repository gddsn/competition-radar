from __future__ import annotations

import logging
from typing import Dict, List, Sequence

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


def _as_list(value: object) -> List[str]:
    if isinstance(value, list):
        return [str(item) for item in value]
    if value is None:
        return []
    return [str(value)]


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
    seed_rows = read_csv(SEED_CSV)
    collected_rows = collect_from_sources(config)
    rows = merge_rows(seed_rows, collected_rows)
    write_csv(rows, LATEST_CSV)
    save_sqlite(rows)
    logging.info("完成采集：种子 %s 条，网页新增 %s 条，合并后 %s 条", len(seed_rows), len(collected_rows), len(rows))


if __name__ == "__main__":
    main()
