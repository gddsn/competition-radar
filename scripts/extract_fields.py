from __future__ import annotations

import re
from typing import Dict, Iterable, List

from utils import LATEST_CSV, TIME_FIELDS, is_blank, normalize_space, read_csv, time_value_rank, write_csv


DATE_TOKEN = (
    r"(?:20\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日?"
    r"|20\d{2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{1,2}"
    r"|20\d{2}\s*年\s*\d{1,2}\s*月"
    r"|20\d{2}\s*[./-]\s*\d{1,2}"
    r"|\d{1,2}\s*月\s*\d{1,2}\s*日?)"
)
DATE_RANGE = rf"{DATE_TOKEN}(?:\s*(?:至|到|—|-|~|－)\s*{DATE_TOKEN})?"
PARAGRAPH_KEYWORDS = [
    "报名时间",
    "报名截止",
    "比赛时间",
    "竞赛时间",
    "提交截止",
    "作品提交",
    "初赛",
    "复赛",
    "决赛",
    "结果公布",
    "获奖公示",
    "公示时间",
]

FIELD_LABELS = {
    "报名时间": ["报名时间", "报名截止", "报名", "注册时间"],
    "比赛时间": ["比赛时间", "竞赛时间", "赛事时间", "提交截止", "作品提交", "初赛", "复赛", "决赛"],
    "结果公布时间": ["结果公布", "获奖公示", "名单公布", "公示时间"],
}


def _split_segments(text: str) -> List[str]:
    compact = normalize_space(text)
    pieces = re.split(r"(?<=[。！？；;])\s*", compact)
    segments = [piece for piece in pieces if any(keyword in piece for keyword in PARAGRAPH_KEYWORDS)]
    if segments:
        return segments
    return [compact]


def _keyword_windows(text: str, labels: Iterable[str], window_size: int = 180) -> List[str]:
    windows: List[str] = []
    for label in labels:
        for match in re.finditer(re.escape(label), text):
            start = max(0, match.start() - 35)
            end = min(len(text), match.end() + window_size)
            windows.append(text[start:end])
    return windows


def _date_priority(value: str) -> int:
    normalized = normalize_space(value)
    if "2026" in normalized:
        return 0
    if "2025" in normalized:
        return 1
    if re.search(r"20\d{2}", normalized):
        return 2
    return 3


def _first_labeled_date(text: str, *labels: str) -> str:
    candidates: List[str] = []
    search_windows: List[str] = []
    for segment in _split_segments(text):
        search_windows.extend(_keyword_windows(segment, labels))
    for window in search_windows:
        for date_match in re.finditer(DATE_RANGE, window):
            candidates.append(normalize_space(date_match.group(0)))
    if candidates:
        return sorted(candidates, key=_date_priority)[0]
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
        "报名时间": _first_labeled_date(cleaned, *FIELD_LABELS["报名时间"]),
        "比赛时间": _first_labeled_date(cleaned, *FIELD_LABELS["比赛时间"]),
        "结果公布时间": _first_labeled_date(cleaned, *FIELD_LABELS["结果公布时间"]),
        "奖金": _sentence_with(cleaned, "奖金", "奖励", "奖品", "万元", "人民币"),
        "证书": _sentence_with(cleaned, "证书", "获奖证明", "参赛证明"),
        "主办单位": _first_labeled_text(cleaned, "主办单位", "主办方", "主办"),
    }


def enrich_row(row: Dict[str, str], text: str) -> Dict[str, str]:
    extracted = extract_from_text(text)
    updated = dict(row)
    for field, value in extracted.items():
        if field in TIME_FIELDS and time_value_rank(value) > time_value_rank(updated.get(field)):
            updated[field] = value
        elif is_blank(updated.get(field)) and not is_blank(value):
            updated[field] = value
    return updated


def main() -> None:
    rows = read_csv(LATEST_CSV)
    enriched = [enrich_row(row, " ".join(row.values())) for row in rows]
    write_csv(enriched, LATEST_CSV)


if __name__ == "__main__":
    main()
