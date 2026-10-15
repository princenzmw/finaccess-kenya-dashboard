# FinAccess 2024 dashboard data

Prepared from the FinAccess Household Survey 2024 (CBK, KNBS, FSD Kenya): https://finaccess.knbs.or.ke/reports-and-datasets

- `prepare_finaccess.py`: streams the 290 MB workbook and writes the files below (about 30 seconds).
- `output/finaccess_adults.csv`: 19,750 adults (18+), 63 columns, includes weighted columns `w_*`.
- `output/data_dictionary.csv`: every column explained.
- `output/dim_county.csv`: counties with sample sizes.
- `output/reconciliation.csv`: weighted results vs published topline (all match).
- `measures.dax.txt`: Power BI measures. `finaccess_theme.json`: Power BI theme.

Always weight percentages (`w_x` / `weight`). Unweighted formal access is 85.7 %, the correct weighted figure is 84.8 %.
Cite the survey as the source on the dashboard. Check the survey site's terms of use before sharing the data publicly.
