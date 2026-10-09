# LHSD Activities with Catering

Monitoring report for the Local Health Support Division, in the same shape as the division’s Power BI view: overall status, funding source, billing submission, purchase-order status, and the activity register.

Figures come from the **Catering Services** tab of the [2026 activity workbook](https://docs.google.com/spreadsheets/d/12wxTAH5bOuzcm1cQffvF1SaNubqu7XhSQzdlhinS7Gk/edit?usp=sharing). The PDF in this repo was generated on **9 October 2026**.

The July 17, 2026 Power BI snapshot reported 292 activities. The sheet has since been updated. This edition uses the current rows.

Hiring, approved positions, and other issues in that Power BI file sit on other tabs. They are not in this PDF.

## Report

[LHSD-Activities-with-Catering.pdf](output/LHSD-Activities-with-Catering.pdf)

| | |
|---|---|
| Activities in the totals | 384 |
| Conducted | 225 (58.59%) |
| To be conducted | 159 (41.41%) |
| Billing submitted | 97 of 225 |
| With a purchase order | 65 of 159 |
| Set aside | 8 |

Set-aside rows are cancelled, dovetailed into another activity, or recalibrated. They are listed on the last page and left out of the totals.

## How a row is counted

- **Conducted.** Date Conducted is on or before the report date, or the row is flagged Conducted and the date is not later. A remark that says “not conducted” stays open.
- **Billing submitted.** Paid, forwarded to finance, or marked billing on process.
- **With PO.** The PO number cell has a real PO number.
- **Typed target years after 2026** (August 12-13, 2027 and the same pattern) are read as 2026. Those years are an autofill on the sheet. Real Excel dates are kept as stored.

## Regenerate

```bash
pip install -r requirements.txt
python3 generate_report.py
```

The script reads `data/catering-services.xlsx` and writes `output/LHSD-Activities-with-Catering.pdf`. It needs Google Chrome. To refresh the workbook first:

```bash
curl -L -o data/catering-services.xlsx \
  "https://docs.google.com/spreadsheets/d/12wxTAH5bOuzcm1cQffvF1SaNubqu7XhSQzdlhinS7Gk/export?format=xlsx"
python3 generate_report.py
```
