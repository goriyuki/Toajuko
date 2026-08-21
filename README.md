# Data Science & Operations Research Portfolio

Hi, I am **Yijun Chen (陈奕均)**, an MSc Data and Decision Analytics student at the University of Southampton. This repository is a curated collection of selected projects from my master's study, undergraduate work, and industry internship.

I am interested in turning ambiguous operational problems into reproducible data workflows, interpretable models, and practical recommendations.

## Featured work

| Project | What I worked on | Methods and tools | Status |
| --- | --- | --- | --- |
| [PCB placement optimisation](case-studies/pcb-placement-optimisation/) | Led a team project that generated and scored machine-head assignments and placement plans | Python, optimisation heuristics, validation, experiment design | Team repository linked with contribution notes |
| [Healthcare appointment simulation](case-studies/healthcare-simulation/) | Built a discrete-event model and compared capacity-allocation policies | AnyLogic, Monte Carlo simulation, confidence intervals | Public case summary; assessed materials withheld |
| [UK economic indicator forecasting](projects/economic-forecasting/) | Prepared monthly indicators and compared time-series forecasts | pandas, ARIMA/SARIMA, statsmodels, backtesting | Reusable, data-agnostic pipeline |
| [Formula similarity matching](projects/formula-similarity/) | Built a two-stage retrieval method for chemical formula data | weighted Jaccard, weighted Euclidean distance, SQL Server | Sanitised internship code |

## Repository map

```text
.
├── case-studies/      # Concise problem-method-result summaries
├── projects/          # Reusable code and project documentation
├── cas爬虫.py          # Legacy chemical-data collection prototype
└── 加权Jaccard香精匹配算法.py  # Legacy entry point; credentials now use environment variables
```

The two root-level scripts are retained to preserve links and commit history. New and substantially revised work follows the English naming and directory conventions under `projects/`.

## Reproducibility and publication policy

- Paths and credentials must come from command-line arguments or environment variables.
- Raw employer data, private databases, student identifiers, assessment briefs, and full assessed reports are not published.
- Team projects link to their original repository and distinguish my contribution from the team's work.
- Each reusable project documents its inputs, environment, commands, outputs, and limitations.

## Contact

- GitHub: [@goriyuki](https://github.com/goriyuki)
- Email: [yijunchen2003@126.com](mailto:yijunchen2003@126.com)

---

中文简介：本仓库用于展示数据科学、预测与运筹优化项目。公开内容均经过脱敏和重构；课程作业题、受限数据及企业内部资料不会上传。
