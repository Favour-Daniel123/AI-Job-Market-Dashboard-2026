#!/usr/bin/env bash
# Rebuilds every output from the dataset. Run from the repository root.
set -e
python src/analysis.py          # models + aggregates -> results/results.json
python src/build_dashboard.py   # -> outputs/ai_job_market_dashboard.html
python src/build_excel.py       # -> outputs/AI_Job_Market_Dashboard.xlsx
npm install --silent && node src/build_deck.js   # -> outputs/AI_Job_Market_Analysis.pptx
cp outputs/ai_job_market_dashboard.html docs/index.html
echo "Done. Open the .xlsx in Excel so its formulas calculate."
