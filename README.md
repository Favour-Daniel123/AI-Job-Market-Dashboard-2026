# AI Job Market Analysis

An analysis of 10,345 AI and data job postings (2020–2026), delivered as an interactive dashboard, an Excel workbook and a PowerPoint deck. It covers five areas:

1. **Salary prediction models**
2. **Job demand analysis**
3. **Machine learning classification tasks**
4. **Workforce analytics**
5. **Skill demand trend analysis**

Everything is computed from the dataset in `data/`. No outside data or assumptions were added.

## Deliverables

| File | What it is |
|---|---|
| [`docs/index.html`](docs/index.html) | Interactive dashboard (six tabs, charts built with Chart.js). Also published via GitHub Pages, see below. |
| [`outputs/AI_Job_Market_Dashboard.xlsx`](outputs/AI_Job_Market_Dashboard.xlsx) | Workbook with live formulas over the data (Job Demand, Workforce Analytics, Skill Trends), plus static model results (Salary Models, ML Classification). |
| [`outputs/AI_Job_Market_Analysis.pptx`](outputs/AI_Job_Market_Analysis.pptx) | 8-slide presentation of the findings with native, editable charts. |
| [`results/results.json`](results/results.json) | All computed numbers used by the three deliverables. |

## Dataset

`data/1--AI_Job_Market_Dataset.xlsx` has two sheets, RAW and CLEANED, which are identical: 10,345 rows, 19 columns, no missing values, no duplicate rows. The analysis uses the CLEANED sheet.

| Group | Columns |
|---|---|
| Role | `job_title`, `experience_level`, `years_experience`, `education_level` |
| Company | `company_size`, `company_industry`, `country`, `remote_type` |
| Skills (0/1 flags) | `skills_python`, `skills_sql`, `skills_ml`, `skills_deep_learning`, `skills_cloud` |
| Market | `salary`, `job_openings`, `hiring_urgency`, `job_posting_year`, `job_posting_month` |

The file does not state the **currency or period of `salary`**, so salaries are reported as recorded.

## Key findings

**Salary prediction.** Three models and a baseline, 80/20 hold-out plus 5-fold cross-validation:

| Model | Test R² | MAE | RMSE |
|---|---|---|---|
| Mean baseline | -0.002 | 25,839 | 31,584 |
| Linear Regression | 0.992 | 2,449 | 2,847 |
| Random Forest | 0.991 | 2,533 | 2,981 |
| Gradient Boosting | 0.992 | 2,485 | 2,906 |

Permutation importance shows salary is driven almost entirely by **experience level** and **job title**, followed by company size and the Machine Learning, Deep Learning and Cloud skill flags. Country, industry, education, remote type, years of experience, posting date, urgency and openings add essentially nothing. Linear Regression matches the tree models, so the relationship is close to additive.

**Job demand.** Openings are spread very evenly: each job title holds 16.0–17.3% of openings, each industry 16.2–17.7%, each country 14.0–14.5% and each posting year 14.1–14.4%.

**Classification** (test accuracy, best model vs majority-class baseline):

| Task | Baseline | Best model |
|---|---|---|
| Hiring urgency (3 classes) | 50.2% | 100.0% |
| Experience level (3 classes) | 34.0% | 99.1% |
| Salary above median (2 classes) | 50.0% | 96.7% |
| Remote type (3 classes) | 34.0% | 33.5% |

Hiring urgency follows a deterministic rule in this file: every posting with 7 or more openings is High; below that, High when both ML and Deep Learning are required, otherwise Low (7 or fewer years of experience) or Medium (more). The 100% score reflects that rule, not a general hiring insight. Remote type cannot be predicted from the other columns.

**Workforce analytics.** Average salary rises with experience level (Entry ≈ 89k, Mid ≈ 114k, Senior ≈ 138k). AI Engineer and Machine Learning Engineer average about 140k versus about 100k for the other four titles, and multinational corporations pay about 20k more than other company sizes. Country, industry, education and remote type differ by at most about 3k.

**Skill demand trends.** Each skill appears in roughly 47–54% of postings in every year, with no sustained trend. Postings requiring Deep Learning, Machine Learning or Cloud pay about 15.2k, 14.6k and 9.7k more (unadjusted); Python and SQL show no meaningful gap.

## Methodology notes

- **Salary models:** all columns except `job_id` and `salary` as features (categoricals one-hot encoded), `random_state=42`.
- **Classification:** stratified 80/20 split and 5-fold CV, compared with a majority-class baseline. For hiring urgency, experience level and remote type, `salary` is included as a feature; for experience level, `years_experience` is excluded. For the salary-above-median task, `salary` is excluded.
- **Skill premium** is the raw difference in mean salary with vs without the skill. It is not adjusted for job title or experience.
- **Excel workbook:** the demand, workforce and skill sheets are formulas over the `Data` sheet. The model and classification sheets hold static values copied from `results/results.json`.

## Caveats

- Salary currency and period are not stated.
- Postings dated 2026 are included exactly as they appear.
- The patterns are very uniform (flat demand, near-identical skill rates, rule-based urgency), which suggests the dataset may be synthetic. Treat the findings as a description of this file, not of the real labour market.

## Reproduce

Requires Python 3.10+ and Node.js 18+. Tested with Python 3.12, pandas 3.0, scikit-learn 1.8, openpyxl 3.1 and pptxgenjs 4.0.

```bash
git clone https://github.com/<your-username>/<your-repo>.git
cd <your-repo>
pip install -r requirements.txt
./run_all.sh
```

Or run the steps individually: `python src/analysis.py`, `python src/build_dashboard.py`, `python src/build_excel.py`, `npm install && node src/build_deck.js`.

Files freshly written by `build_excel.py` hold formulas without cached values; Excel calculates them when the file is opened.

## Publish the dashboard with GitHub Pages

1. Push the repository to GitHub.
2. Go to **Settings → Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, branch `main`, folder `/docs`, then save.
4. After a minute the dashboard is live at `https://<your-username>.github.io/<your-repo>/`.

The dashboard loads Chart.js from a CDN, so it needs an internet connection.

## Repository structure

```
├── data/                 source dataset
├── src/                  analysis and build scripts
├── results/results.json  computed results
├── outputs/              xlsx, pptx, html deliverables
├── docs/index.html       dashboard (GitHub Pages)
├── run_all.sh
├── requirements.txt
└── package.json
```
