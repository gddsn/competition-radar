from __future__ import annotations

import re
from typing import Dict

from utils import LATEST_CSV, is_blank, normalize_space, read_csv, write_csv


DATE_TOKEN = (
    r"(?:20\d{2}\s*[年./-]\s*\d{1,2}\s*(?:[月./-]\s*\d{1,2}\s*日?)?"
    r"|20\d{2}\s*年"
    r"|\d{1,2}\s*月\s*\d{1,2}\s*日?)"
)
DATE_RANGE = rf"{DATE_TOKEN}(?:\s*(?:至|到|—|-|~|－)\s*{DATE_TOKEN})?"


def _first_labeled_date(text: str, *labels: str) -> str:
    for label in labels:
        match = re.search(label, text)
        if not match:
            continue
        window = text[match.start() : match.start() + 140]
        date_match = re.search(DATE_RANGE, window)
        if date_match:
            return normalize_space(date_match.group(0))
    return ""


def _first_labeled_text(text: str, *labels: str, max_chars: int = 90) -> str:
    for label in labels:
        pattern = rf"(?:{label})\s*[:：]?\s*([^。；;\n]{{2,{max_chars}}})"
        match = re.search(pattern, text)
        if match:
            return normalize_space(match.group(1))
    return ""


def _sentence_with(text: str, *keywords: str, max_chars: int = 100) -> str:
    for keyword in keywords:
        index = text.find(keyword)
        if index < 0:
            continue
        start = max(0, index - 20)
        end = min(len(text), index + max_chars)
        return normalize_space(text[start:end])
    return ""


def extract_from_text(text: str) -> Dict[str, str]:
    cleaned = normalize_space(text)
    return {
        "报名时间": _first_labeled_date(cleaned, "报名时间", "报名截止", "报名", "注册时间"),
        "比赛时间": _first_labeled_date(cleaned, "比赛时间", "竞赛时间", "赛事时间", "决赛", "初赛"),
        "结果公布时间": _first_labeled_date(cleaned, "结果公布", "获奖公示", "名单公布", "公示时间"),
        "奖金": _sentence_with(cleaned, "奖金", "奖励", "奖品", "万元", "人民币"),
        "证书": _sentence_with(cleaned, "证书", "获奖证明", "参赛证明"),
        "主办单位": _first_labeled_text(cleaned, "主办单位", "主办方", "主办"),
    }


def enrich_row(row: Dict[str, str], text: str) -> Dict[str, str]:
    extracted = extract_from_text(text)
    updated = dict(row)
    for field, value in extracted.items():
        if is_blank(updated.get(field)) and not is_blank(value):
            updated[field] = value
    return updated


def main() -> None:
    rows = read_csv(LATEST_CSV)
    enriched = [enrich_row(row, " ".join(row.values())) for row in rows]
    write_csv(enriched, LATEST_CSV)


if __name__ == "__main__":
    main()

