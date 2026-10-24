# FINA 4930 Mini Project 1 — What drives crowdfunding success?

Regression analysis that predicts the fraction of the funding goal a crowdfunding campaign reaches (`frac_raised`).

- `crowdfunding_analysis.ipynb`: the full analysis (EDA, de-duplication, nested-CV model comparison, tuning, test evaluation, variable relevance)
- `Crowdfunding_Report.docx`: the written report, including policy implications
- `data/crowdfunding_campaign.xlsx`: the source data (Sheet1 = data, Sheet2 = dictionary)
- `figures/`, `results/`: outputs written by the notebook
- `build_report.py`: builds the report from `figures/` and `results/`

## Run

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/jupyter nbconvert --to notebook --execute crowdfunding_analysis.ipynb --inplace
.venv/bin/python build_report.py
```

The notebook takes about 4 minutes, most of it nested cross-validation.
