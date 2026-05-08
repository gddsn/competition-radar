# competition-radar

大学生竞赛情报系统 MVP。它每周自动整理适合大学生参加、并对金融数学考研、金融就业、商赛简历、数学建模、数据分析、量化、风控、投研有帮助的竞赛信息。

## 功能

- 配置化维护竞赛信息源：`data/sources.yml`
- 合并人工种子数据与公开网页采集结果：`data/competitions_seed.csv`
- 抽取报名时间、比赛时间、结果公布时间、奖金、证书、主办单位、官方链接
- 生成含金量评级、金融路线价值评分、商赛价值评分、数学建模能力评分、数据分析/量化/风控/投研评分
- 输出 `CSV`、`SQLite`、`Markdown` 周报、`Excel` 表格和竞赛日程报告
- 通过 GitHub Actions 每周日北京时间 00:00 自动刷新

## 快速开始

```bash
cd competition-radar
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python scripts/collect_competitions.py
python scripts/score_competitions.py
python scripts/generate_report.py
python scripts/generate_schedule_report.py
```

生成结果：

- `data/competitions_latest.csv`
- `data/competitions.db`
- `reports/weekly_report.md`
- `reports/weekly_report.xlsx`
- `reports/schedule_report.md`
- `reports/schedule_report.xlsx`

`weekly_report` 主要用于判断含金量、金融路线价值、商赛简历价值、数学建模能力和数据分析/量化适配度。`schedule_report` 主要用于查看比赛时间、报名/截止时间、结果公布/获奖公示时间。

## 数据字段

核心字段包括：比赛名称、类别、含金量评级、含金量分、报名时间、比赛时间、结果公布时间、奖金、证书、主办单位、官方链接、金融路线价值评分、商赛价值评分、数学建模能力评分、数据分析量化评分、推荐参加优先级、推荐理由、信息来源、来源页面、更新时间。

## 评分逻辑

MVP 使用可解释的关键词和类别规则：

- 含金量：参考主办单位、全国性/国际性、是否常见于高校竞赛榜单、是否有教育部或权威学会背景
- 金融路线价值：金融、量化、风控、投研、证券、银行、保险、金融科技等关键词
- 商赛价值：商业案例、咨询、市场调研、创业、管理、财务分析等关键词
- 数学建模价值：数学建模、统计建模、运筹优化、预测、仿真、论文写作等关键词
- 数据分析价值：数据分析、统计、算法、AI、机器学习、Python、量化研究等关键词

评分规则集中在 `scripts/score_competitions.py`，后续可以按你的简历目标继续调权重。

## 每周自动运行

`.github/workflows/weekly_competition_report.yml` 会在北京时间每周日 00:00 自动运行。GitHub Actions 的 cron 使用 UTC 时间，因此配置为：

```yaml
0 16 * * 6
```

也就是 UTC 周六 16:00，对应北京时间周日 00:00。工作流会安装依赖，依次执行采集、评分、周报生成和日程报告生成，并在数据变化时自动提交更新。

所有比赛时间都以官方通知为准。日程报告中标记为“待核实”或“周期性/待核实”的比赛，不要直接导入日历，应先回到官方链接确认。

## 后续扩展方向

- 接入公众号选题：从周报中自动筛选“本周最值得关注的 5 个比赛”
- 接入网站：把 `competitions_latest.csv` 或 `competitions.db` 作为后端数据源
- 接入题库项目：把数学建模、统计、Python、金融分析能力要求映射到训练题单
- 兼容 MATLAB：后续为数模赛题分析、仿真、优化模型保留 MATLAB 脚本目录
