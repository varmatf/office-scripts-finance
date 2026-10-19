# Office Scripts for Finance

Practical, sanitized Office Scripts for finance and FP&A work. Each script uses dummy data and is meant to be adapted, not used as-is.

## Scripts

| Script | What it does |
|---|---|
| [`budget-vs-actuals-variance.ts`](scripts/budget-vs-actuals-variance.ts) | Budget vs actuals variance at account and account-group level, with configurable thresholds |

More scripts will be added over time.

---

## Budget vs Actuals Variance Report

### The problem

Most variance reports are built with lookups and conditional formatting, then rebuilt every month. Thresholds for "what needs explaining" are hard-coded, or live in someone's head.

### What the script does

1. Reads two tables, `tblBudget` and `tblActuals`, each with `AccountGroup`, `Account` and `Amount` columns
2. Joins them by account. Accounts that appear on only one side are kept, so nothing silently drops out
3. Calculates **variance** (Actual − Budget) and **variance %** for each account
4. Rolls everything up to **account-group** level
5. Flags lines where **both** thresholds are breached: the amount *and* the percentage
6. Writes a fresh **Variance Report** sheet with a group summary on top and account detail below

### Why two thresholds?

Either threshold on its own produces noise:

- An amount-only threshold flags large accounts with tiny percentage swings
- A percentage-only threshold flags small accounts where the dollars don't matter

Requiring both keeps the review list focused. In the sample data, *Product Sales* misses budget by 18,500, which is only 7.4%, so it isn't flagged. *Professional Fees* is over by 7,900, which is 56%, so it is.

Lines with no budget but material actuals are flagged as **Review (no budget)**.

### Thresholds live in the workbook, not the code

The `Settings` sheet holds a small table:

| Setting | Value |
|---|---|
| AbsThreshold | 5000 |
| PctThreshold | 0.10 |

Reviewers can change materiality without opening the script.

### The pattern worth copying

```
Read all values once  →  calculate in memory  →  write back once
```

Scripts get slow when they read or write cells one at a time inside a loop. This script makes a handful of calls to Excel regardless of how many accounts there are. Flagged rows are highlighted with a single conditional format rather than formatting rows one by one.

---

## How to use it

1. Download [`sample-data/budget-vs-actuals-sample.xlsx`](sample-data/budget-vs-actuals-sample.xlsx) and open it in Excel (desktop or web)
2. Go to **Automate → New Script**
3. Delete the placeholder code and paste in the contents of [`budget-vs-actuals-variance.ts`](scripts/budget-vs-actuals-variance.ts)
4. Click **Run**
5. A **Variance Report** sheet appears. Run it again after changing data or thresholds; the sheet is rebuilt each time

**Expected flags with the sample data:** Materials, Contractors, Events, Sponsorships, Professional Fees, and Bank Charges (no budget).

### Using your own data

Format your budget and actuals ranges as Excel tables (**Ctrl+T**) and name them `tblBudget` and `tblActuals` under **Table Design → Table Name**. Column order doesn't matter, but the header names must match. If something is missing, the script stops with a message telling you what it couldn't find.

### Running it automatically

Office Scripts can be triggered from Power Automate using the **Excel Online (Business) → Run script** action. For example, a flow could run this script whenever an updated actuals file lands in SharePoint.

---

## Ideas for extending it

- **Favourable / unfavourable:** add an `AccountType` column (Revenue / Expense) and flip the sign for expenses, so positive always means favourable
- **Commentary column:** add a blank column for reviewers to explain flagged lines
- **Prior-year comparison:** add a third table and a second variance set
- **Group-level thresholds:** different materiality for different account groups

## Requirements

- Excel for the web, or Excel desktop on Windows or Mac with a Microsoft 365 business or education licence that includes Office Scripts
- The **Automate** tab must be enabled by your admin

## Disclaimer

All data in this repo is fictional. Scripts are provided as learning examples. Test on a copy of your workbook before using with real data.

## License

MIT. See [LICENSE](LICENSE).
