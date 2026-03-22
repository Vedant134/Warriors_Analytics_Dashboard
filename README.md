# Warriors Dashboard Analyzer

This data project analyzes Golden State Warriors personnel data to identify elite lineup combinations following the Butler trade. It features a Python pipeline that ingests raw CSV matchup data, strips out unnecessary metrics, cleanly filters the noise, and algorithmically determines the highest performing groupings using advanced NBA stats like Net Rating, True Shooting %, and Offensive/Defensive Ratings.

## Key Project Deliverables

1. **Executive Excel Dashboard (`warriors_lineup_dashboard.xlsx`)**: 
   A professionally designed, highly-stylized Excel layout utilizing corporate aesthetics. It features:
   - Three key KPI scorecards at the top summarizing the peak ratings (Net Rating, Offensive Rating, TS%).
   - Four distinct, visually appealing charts mapping out offensive vs defensive tradeoffs and shooting efficiencies.
   - An "Executive Insights" text panel natively integrated within the spreadsheet boundary.

2. **Executive Summary PDF (`project_summary.pdf`)**: 
   A structured 1-page document summarizing project methodology, ETL data hygiene steps, the precise algorithmic insights generated, and a breakdown of the dashboard elements cleanly presented in a tabular and bulleted corporate format.

## Technology Stack
- **Python**
- **Pandas**
- **XlsxWriter**.
- **ReportLab**

## How To Run
Running the primary script executes the entire data pipeline and produces the Excel and PDF outputs cleanly:
```bash
python generate_dash.py
```
