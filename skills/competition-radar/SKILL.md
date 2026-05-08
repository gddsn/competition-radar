---
name: competition-radar
description: Maintain the competition-radar project for collecting, scoring, and reporting Chinese university competition opportunities, especially for finance math, modeling, business case, data analysis, quant, risk, investment research, and student resume goals.
---

# Competition Radar

Use this skill when updating the `competition-radar` project: competition sources, seed data, scoring rules, weekly reports, or automation.

## Workflow

1. Inspect `data/sources.yml`, `data/competitions_seed.csv`, and the scripts before editing.
2. Keep the MVP pipeline runnable with:
   ```bash
   python scripts/collect_competitions.py
   python scripts/score_competitions.py
   python scripts/generate_report.py
   ```
3. Preserve the CSV field names in `scripts/utils.py` unless the report and Excel generator are updated together.
4. Prefer explainable scoring changes in `scripts/score_competitions.py`; do not add opaque ranking logic.
5. Add new public sources only when they do not require login, cookies, paid APIs, or secrets.

## Data Rules

- Keep seed rows useful even when websites fail.
- Leave uncertain dates as `以年度通知为准` instead of inventing exact dates.
- Put official links in `官方链接` only when reasonably confident; otherwise leave the cell blank for later verification.
- Deduplicate by competition name and official link.

## Report Rules

- `weekly_report.md` should highlight P0/P1 competitions first.
- `weekly_report.xlsx` should keep both a focused recommendation sheet and a complete data sheet.
- Prioritize competitions that support finance math graduate applications, quant/risk/investment research internships, business analysis resumes, modeling ability, or data analysis skills.

