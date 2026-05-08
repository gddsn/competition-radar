from __future__ import annotations

from typing import Dict, Iterable, List, Tuple

from utils import LATEST_CSV, clamp, fit_label, read_csv, save_sqlite, write_csv


CATEGORY_BASE = {
    "数学建模类": 68,
    "金融数学 / 金融科技 / 量化 / 风控相关比赛": 66,
    "商业案例分析类": 58,
    "创新创业类": 62,
    "数据分析 / 统计建模类": 64,
    "编程 / 算法 / AI 应用类": 60,
    "英语与财经表达类": 46,
}

PRESTIGE_KEYWORDS = {
    "教育部": 16,
    "全国大学生": 14,
    "中国高等教育学会": 14,
    "国家级": 12,
    "国际": 12,
    "中国工业与应用数学学会": 12,
    "中国计算机学会": 10,
    "共青团中央": 10,
    "CFA": 10,
    "高教社杯": 10,
    "挑战杯": 12,
    "中国国际大学生创新大赛": 14,
    "MCM": 12,
    "ICM": 12,
}

FINANCE_KEYWORDS = {
    "金融": 18,
    "金融科技": 20,
    "量化": 18,
    "风控": 18,
    "风险管理": 16,
    "投研": 16,
    "投资": 12,
    "证券": 12,
    "银行": 10,
    "保险": 10,
    "CFA": 18,
    "商业分析": 10,
    "财务": 10,
    "数据分析": 8,
}

BUSINESS_KEYWORDS = {
    "商业": 16,
    "案例": 14,
    "市场调查": 18,
    "创业": 14,
    "创新创业": 16,
    "管理": 10,
    "咨询": 12,
    "财务": 10,
    "营销": 10,
    "电子商务": 14,
    "Research Challenge": 12,
}

MODELING_KEYWORDS = {
    "数学建模": 26,
    "统计建模": 22,
    "建模": 16,
    "运筹": 14,
    "优化": 12,
    "预测": 12,
    "仿真": 10,
    "论文": 8,
    "MCM": 20,
    "ICM": 20,
}

DATA_KEYWORDS = {
    "数据": 16,
    "数据分析": 20,
    "统计": 18,
    "大数据": 20,
    "算法": 16,
    "人工智能": 18,
    "AI": 16,
    "机器学习": 18,
    "计算智能": 16,
    "Python": 12,
    "量化": 14,
    "风控": 14,
    "投研": 12,
}


def _combined_text(row: Dict[str, str]) -> str:
    return " ".join(str(value) for value in row.values())


def _keyword_score(text: str, weights: Dict[str, int]) -> int:
    lowered = text.lower()
    return sum(weight for keyword, weight in weights.items() if keyword.lower() in lowered)


def _rating(score: int) -> str:
    if score >= 86:
        return "S"
    if score >= 72:
        return "A"
    if score >= 58:
        return "B"
    return "C"


def _priority(overall: int) -> str:
    if overall >= 82:
        return "P0 强烈推荐"
    if overall >= 68:
        return "P1 优先关注"
    if overall >= 52:
        return "P2 可选参加"
    return "P3 低优先级"


def _top_reasons(items: Iterable[Tuple[str, int]]) -> str:
    labels = [label for label, score in sorted(items, key=lambda item: item[1], reverse=True) if score >= 60]
    if not labels:
        return "适合作为补充经历，建议结合时间成本选择。"
    return "、".join(labels[:3]) + "价值较高，建议优先跟进。"


def score_row(row: Dict[str, str]) -> Dict[str, str]:
    scored = dict(row)
    text = _combined_text(row)
    category = row.get("类别", "")
    prestige = clamp(CATEGORY_BASE.get(category, 45) + _keyword_score(text, PRESTIGE_KEYWORDS))
    finance = clamp(24 + _keyword_score(text, FINANCE_KEYWORDS))
    business = clamp(22 + _keyword_score(text, BUSINESS_KEYWORDS))
    modeling = clamp(20 + _keyword_score(text, MODELING_KEYWORDS))
    data_quant = clamp(24 + _keyword_score(text, DATA_KEYWORDS))

    if "数学建模" in category:
        modeling = clamp(modeling + 18)
        data_quant = clamp(data_quant + 8)
        finance = clamp(finance + 22)
    if "金融" in category or "量化" in category or "风控" in category:
        finance = clamp(finance + 18)
        data_quant = clamp(data_quant + 8)
    if "商业" in category:
        business = clamp(business + 18)
        finance = clamp(finance + 6)
    if "创新创业" in category:
        business = clamp(business + 12)
    if "数据分析" in category or "统计建模" in category:
        data_quant = clamp(data_quant + 16)
        modeling = clamp(modeling + 8)
        finance = clamp(finance + 10)
    if "编程" in category or "AI" in category:
        data_quant = clamp(data_quant + 12)

    overall = clamp(
        prestige * 0.30
        + finance * 0.25
        + data_quant * 0.20
        + modeling * 0.15
        + business * 0.10
    )

    scored["含金量分"] = str(prestige)
    scored["含金量评级"] = _rating(prestige)
    scored["金融路线价值评分"] = str(finance)
    scored["金融数学/金融就业适配"] = fit_label(finance)
    scored["商赛价值评分"] = str(business)
    scored["商赛简历适配"] = fit_label(business)
    scored["数学建模能力评分"] = str(modeling)
    scored["数学建模适配"] = fit_label(modeling)
    scored["数据分析量化评分"] = str(data_quant)
    scored["数据分析/量化/风控/投研适配"] = fit_label(data_quant)
    scored["推荐参加优先级"] = _priority(overall)
    scored["推荐理由"] = _top_reasons(
        [
            ("金融路线", finance),
            ("数据分析/量化", data_quant),
            ("数学建模", modeling),
            ("商赛简历", business),
            ("含金量", prestige),
        ]
    )
    return scored


def main() -> None:
    rows = read_csv(LATEST_CSV)
    scored_rows: List[Dict[str, str]] = [score_row(row) for row in rows]
    write_csv(scored_rows, LATEST_CSV)
    save_sqlite(scored_rows)
    print(f"Scored {len(scored_rows)} competitions.")


if __name__ == "__main__":
    main()
