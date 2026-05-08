from __future__ import annotations

import csv
import datetime as dt
import logging
import re
import sqlite3
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple
from urllib.parse import urldefrag, urljoin, urlparse

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover - Python < 3.9 fallback.
    ZoneInfo = None  # type: ignore[assignment]

import requests
import yaml
from bs4 import BeautifulSoup


ROOT_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT_DIR / "data"
REPORTS_DIR = ROOT_DIR / "reports"
SOURCES_YML = DATA_DIR / "sources.yml"
SEED_CSV = DATA_DIR / "competitions_seed.csv"
LATEST_CSV = DATA_DIR / "competitions_latest.csv"
DB_PATH = DATA_DIR / "competitions.db"
WEEKLY_MD = REPORTS_DIR / "weekly_report.md"
WEEKLY_XLSX = REPORTS_DIR / "weekly_report.xlsx"
SCHEDULE_MD = REPORTS_DIR / "schedule_report.md"
SCHEDULE_XLSX = REPORTS_DIR / "schedule_report.xlsx"

if ZoneInfo is not None:
    BEIJING_TZ = ZoneInfo("Asia/Shanghai")
else:
    BEIJING_TZ = dt.timezone(dt.timedelta(hours=8), name="Asia/Shanghai")

FIELDS = [
    "比赛名称",
    "类别",
    "含金量评级",
    "含金量分",
    "报名时间",
    "比赛时间",
    "结果公布时间",
    "奖金",
    "证书",
    "主办单位",
    "官方链接",
    "金融数学/金融就业适配",
    "金融路线价值评分",
    "商赛简历适配",
    "商赛价值评分",
    "数学建模适配",
    "数学建模能力评分",
    "数据分析/量化/风控/投研适配",
    "数据分析量化评分",
    "推荐参加优先级",
    "推荐理由",
    "信息来源",
    "来源页面",
    "更新时间",
]
TIME_FIELDS = {"报名时间", "比赛时间", "结果公布时间"}
UNCERTAIN_TIME_PATTERNS = [
    "通常每年",
    "每年上半年",
    "每年下半年",
    "每年春季",
    "每年秋季",
    "每年",
    "以年度通知为准",
    "关注学校通知",
    "赛后数月",
    "待核实",
    "待核验",
    "待定",
    "另行通知",
    "暂未公布",
]
EXPLICIT_DATE_RE = re.compile(
    r"20\d{2}\s*(?:年\s*\d{1,2}\s*月(?:\s*\d{1,2}\s*日?)?|[./-]\s*\d{1,2}(?:[./-]\s*\d{1,2})?)"
)

REQUEST_HEADERS = {
    "User-Agent": (
        "competition-radar/0.1 "
        "(public academic competition monitor; contact via repository)"
    )
}


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s %(message)s",
    )


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)


def today_str() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d")


def now_beijing_str() -> str:
    return dt.datetime.now(BEIJING_TZ).strftime("%Y-%m-%d %H:%M")


def today_beijing_str() -> str:
    return dt.datetime.now(BEIJING_TZ).strftime("%Y-%m-%d")


def normalize_space(value: object) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", text).strip()


def is_blank(value: object) -> bool:
    return normalize_space(value) in {"", "待补充", "待核验", "None", "nan"}


def time_value_rank(value: object) -> int:
    text = normalize_space(value)
    if is_blank(text):
        return 0
    if EXPLICIT_DATE_RE.search(text):
        return 3
    if any(pattern in text for pattern in UNCERTAIN_TIME_PATTERNS):
        return 2
    return 1


def ensure_row(row: Dict[str, object]) -> Dict[str, str]:
    cleaned = {field: normalize_space(row.get(field, "")) for field in FIELDS}
    if is_blank(cleaned["更新时间"]):
        cleaned["更新时间"] = today_str()
    return cleaned


def read_csv(path: Path) -> List[Dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        return [ensure_row(row) for row in reader]


def write_csv(rows: Sequence[Dict[str, object]], path: Path) -> None:
    ensure_dirs()
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=FIELDS, extrasaction="ignore", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            writer.writerow(ensure_row(row))


def load_yaml(path: Path) -> Dict[str, object]:
    with path.open("r", encoding="utf-8") as file:
        loaded = yaml.safe_load(file) or {}
    return loaded


def fetch_url(url: str, timeout: int = 12) -> Tuple[str, str]:
    try:
        response = requests.get(url, headers=REQUEST_HEADERS, timeout=timeout)
        response.raise_for_status()
        if not response.encoding:
            response.encoding = response.apparent_encoding or "utf-8"
        return response.text, ""
    except requests.RequestException as exc:
        return "", f"{type(exc).__name__}: {exc}"


def html_to_text(html: str) -> str:
    soup = BeautifulSoup(html or "", "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    return normalize_space(soup.get_text(" ", strip=True))


def html_title(html: str) -> str:
    soup = BeautifulSoup(html or "", "html.parser")
    if soup.title and soup.title.string:
        return normalize_space(soup.title.string)
    return ""


def keyword_hit(text: str, keywords: Iterable[str]) -> bool:
    lowered = text.lower()
    return any(normalize_space(keyword).lower() in lowered for keyword in keywords if keyword)


def discover_candidates(
    html: str,
    base_url: str,
    source_name: str,
    category: str,
    keywords: Sequence[str],
    max_items: int,
) -> List[Dict[str, str]]:
    soup = BeautifulSoup(html or "", "html.parser")
    candidates: List[Dict[str, str]] = []
    seen: set[str] = set()

    title = html_title(html)
    if title and keyword_hit(title, keywords):
        candidates.append(
            ensure_row(
                {
                    "比赛名称": title,
                    "类别": category,
                    "官方链接": base_url,
                    "信息来源": source_name,
                    "来源页面": base_url,
                }
            )
        )
        seen.add(row_identity(candidates[-1]))

    for anchor in soup.find_all("a"):
        text = normalize_space(anchor.get_text(" ", strip=True))
        href = normalize_space(anchor.get("href", ""))
        if len(text) < 4 or len(text) > 120:
            continue
        link = urldefrag(urljoin(base_url, href)).url if href else base_url
        if urlparse(link).scheme not in {"http", "https"}:
            continue
        if not keyword_hit(f"{text} {link}", keywords):
            continue
        row = ensure_row(
            {
                "比赛名称": text,
                "类别": category,
                "官方链接": link,
                "信息来源": source_name,
                "来源页面": base_url,
            }
        )
        identity = row_identity(row)
        if identity in seen:
            continue
        seen.add(identity)
        candidates.append(row)
        if len(candidates) >= max_items:
            break
    return candidates


def canonical_url(url: str) -> str:
    parsed = urlparse(normalize_space(url))
    if not parsed.netloc:
        return ""
    path = re.sub(r"/+$", "", parsed.path)
    return f"{parsed.netloc.lower()}{path.lower()}"


def row_identity(row: Dict[str, object]) -> str:
    name = normalize_space(row.get("比赛名称", "")).lower()
    link = canonical_url(str(row.get("官方链接", "")))
    return f"{name}|{link}" if link else name


def merge_rows(*groups: Sequence[Dict[str, object]]) -> List[Dict[str, str]]:
    merged: Dict[str, Dict[str, str]] = {}
    order: List[str] = []
    for group in groups:
        for raw_row in group:
            row = ensure_row(raw_row)
            if is_blank(row["比赛名称"]):
                continue
            identity = row_identity(row)
            if not identity:
                continue
            if identity not in merged:
                merged[identity] = row
                order.append(identity)
                continue
            current = merged[identity]
            for field in FIELDS:
                if field in TIME_FIELDS:
                    if time_value_rank(row.get(field)) > time_value_rank(current.get(field)):
                        current[field] = row[field]
                elif is_blank(current.get(field)) and not is_blank(row.get(field)):
                    current[field] = row[field]
    return [merged[identity] for identity in order]


def clamp(value: float, minimum: int = 0, maximum: int = 100) -> int:
    return int(max(minimum, min(maximum, round(value))))


def fit_label(score: int) -> str:
    if score >= 75:
        return "高"
    if score >= 45:
        return "中"
    return "低"


def save_sqlite(rows: Sequence[Dict[str, object]], path: Path = DB_PATH) -> None:
    ensure_dirs()
    connection = sqlite3.connect(path)
    try:
        columns = ", ".join(f'"{field}" TEXT' for field in FIELDS)
        connection.execute(f"CREATE TABLE IF NOT EXISTS competitions ({columns})")
        connection.execute("DELETE FROM competitions")
        placeholders = ", ".join("?" for _ in FIELDS)
        quoted_fields = ", ".join(f'"{field}"' for field in FIELDS)
        sql = f"INSERT INTO competitions ({quoted_fields}) VALUES ({placeholders})"
        connection.executemany(
            sql,
            [[ensure_row(row).get(field, "") for field in FIELDS] for row in rows],
        )
        connection.commit()
    finally:
        connection.close()
