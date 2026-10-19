/**
 * Budget vs Actuals Variance Report — Office Script
 *
 * What it does
 *   1. Reads two Excel tables: tblBudget and tblActuals
 *      (columns: AccountGroup, Account, Amount)
 *   2. Joins them by Account (accounts missing on either side are kept)
 *   3. Calculates variance and variance % at account level
 *   4. Rolls the results up to account-group level
 *   5. Flags lines that breach BOTH thresholds (amount AND percentage)
 *   6. Writes everything to a fresh "Variance Report" sheet in one pass
 *
 * Thresholds
 *   Read from the tblSettings table (Setting | Value) so reviewers can change
 *   materiality without touching code. If a setting is missing, the defaults
 *   below are used.
 *
 * Pattern worth copying
 *   Read all values once -> calculate in memory -> write back once.
 *   Avoid reading or writing cell by cell inside loops; that is what makes
 *   scripts slow.
 *
 * Sign convention
 *   Variance = Actual - Budget. Whether a positive variance is favourable
 *   depends on the account type (revenue vs expense). See README for an
 *   extension idea.
 */

// ---------- Defaults (used only if tblSettings is missing a value) ----------
const DEFAULT_ABS_THRESHOLD = 5000; // flag if |variance| >= this amount ...
const DEFAULT_PCT_THRESHOLD = 0.1; // ... AND |variance %| >= this (0.10 = 10%)

const OUTPUT_SHEET = "Variance Report";

interface Line {
  group: string;
  account: string;
  budget: number;
  actual: number;
}

interface Settings {
  absThreshold: number;
  pctThreshold: number;
}

function main(workbook: ExcelScript.Workbook) {
  const settings = readSettings(workbook);

  const budgetRows = readAmountTable(workbook, "tblBudget");
  const actualRows = readAmountTable(workbook, "tblActuals");

  // ---- 1. Join budget and actuals by account (full outer join) ----
  const lines = new Map<string, Line>();

  for (const r of budgetRows) {
    const line = getOrCreate(lines, r.account, r.group);
    line.budget += r.amount;
  }
  for (const r of actualRows) {
    const line = getOrCreate(lines, r.account, r.group);
    line.actual += r.amount;
  }

  const accountLines = Array.from(lines.values()).sort(
    (a, b) => a.group.localeCompare(b.group) || a.account.localeCompare(b.account)
  );

  // ---- 2. Roll up to account-group level ----
  const groups = new Map<string, Line>();
  for (const l of accountLines) {
    const g = getOrCreate(groups, l.group, l.group);
    g.budget += l.budget;
    g.actual += l.actual;
  }
  const groupLines = Array.from(groups.values()).sort((a, b) =>
    a.group.localeCompare(b.group)
  );

  // ---- 3. Build output arrays (still in memory) ----
  const accountHeader = ["Account Group", "Account", "Budget", "Actual", "Variance", "Variance %", "Flag"];
  const accountData: (string | number)[][] = [accountHeader];
  for (const l of accountLines) {
    accountData.push(toRow([l.group, l.account], l, settings));
  }

  const groupHeader = ["Account Group", "Budget", "Actual", "Variance", "Variance %", "Flag"];
  const groupData: (string | number)[][] = [groupHeader];
  for (const g of groupLines) {
    groupData.push(toRow([g.group], g, settings));
  }

  // ---- 4. Write to a fresh output sheet in one pass per table ----
  const existing = workbook.getWorksheet(OUTPUT_SHEET);
  if (existing) {
    existing.delete();
  }
  const sheet = workbook.addWorksheet(OUTPUT_SHEET);

  // Header notes showing which thresholds were applied
  sheet.getRange("A1:B1").setValues([["Thresholds applied", ""]]);
  sheet
    .getRange("A2:B3")
    .setValues([
      ["Amount (abs) >=", settings.absThreshold],
      ["Variance % (abs) >=", settings.pctThreshold],
    ]);
  sheet.getRange("A1").getFormat().getFont().setBold(true);
  sheet.getRange("B2").setNumberFormat("#,##0");
  sheet.getRange("B3").setNumberFormat("0.0%");

  // Group summary first (row 5), account detail below it
  const groupStartRow = 4; // zero-based -> row 5
  writeTable(sheet, groupData, groupStartRow, "tblGroupVariance", 1);

  const accountStartRow = groupStartRow + groupData.length + 2;
  writeTable(sheet, accountData, accountStartRow, "tblAccountVariance", 2);

  sheet.getRange("A:G").getFormat().autofitColumns();
  sheet.activate();
}

// ---------------- Helpers ----------------

function getOrCreate(map: Map<string, Line>, key: string, group: string): Line {
  let line = map.get(key);
  if (!line) {
    line = { group: group, account: key, budget: 0, actual: 0 };
    map.set(key, line);
  }
  return line;
}

/** Builds one output row: leading label columns + numbers + flag. */
function toRow(labels: string[], l: Line, s: Settings): (string | number)[] {
  const variance = l.actual - l.budget;
  // Variance % is undefined when budget is zero; leave it blank
  const pct: number | null = l.budget !== 0 ? variance / Math.abs(l.budget) : null;

  let flag = "";
  if (Math.abs(variance) >= s.absThreshold) {
    if (pct === null || Math.abs(pct) >= s.pctThreshold) {
      flag = pct === null ? "Review (no budget)" : "Review";
    }
  }

  return [...labels, l.budget, l.actual, variance, pct === null ? "" : pct, flag];
}

/**
 * Writes a 2D array as an Excel table and formats it.
 * firstNumberCol: zero-based index of the first numeric column (Budget).
 */
function writeTable(
  sheet: ExcelScript.Worksheet,
  data: (string | number)[][],
  startRow: number,
  tableName: string,
  firstNumberCol: number
) {
  const rows = data.length;
  const cols = data[0].length;
  const range = sheet.getRangeByIndexes(startRow, 0, rows, cols);
  range.setValues(data);

  const table = sheet.addTable(range.getAddress(), true);
  table.setName(tableName);

  if (rows > 1) {
    // Budget, Actual, Variance -> number format; Variance % -> percent
    sheet
      .getRangeByIndexes(startRow + 1, firstNumberCol, rows - 1, 3)
      .setNumberFormat("#,##0;[Red](#,##0)");
    sheet
      .getRangeByIndexes(startRow + 1, firstNumberCol + 3, rows - 1, 1)
      .setNumberFormat("0.0%");

    // Highlight flagged rows with one conditional format (no per-row loops)
    const body = sheet.getRangeByIndexes(startRow + 1, 0, rows - 1, cols);
    const flagColLetter = columnLetter(cols - 1);
    const firstDataRow = startRow + 2; // 1-based row number of first data row
    const cf = body.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
    cf.getCustom()
      .getRule()
      .setFormula(`=LEFT($${flagColLetter}${firstDataRow},6)="Review"`);
    cf.getCustom().getFormat().getFill().setColor("#FDE2E1");
  }
}

/** Reads a table with AccountGroup, Account, Amount columns (any order). */
function readAmountTable(
  workbook: ExcelScript.Workbook,
  tableName: string
): { group: string; account: string; amount: number }[] {
  const table = workbook.getTable(tableName);
  if (!table) {
    throw new Error(`Table "${tableName}" not found. Check the table name in Table Design.`);
  }

  const headers = table.getHeaderRowRange().getValues()[0].map((h) => String(h).trim());
  const gi = requireColumn(headers, "AccountGroup", tableName);
  const ai = requireColumn(headers, "Account", tableName);
  const mi = requireColumn(headers, "Amount", tableName);

  if (table.getRowCount() === 0) {
    return [];
  }

  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const out: { group: string; account: string; amount: number }[] = [];
  for (const row of values) {
    const account = String(row[ai]).trim();
    if (account === "") continue; // skip blank rows
    out.push({
      group: String(row[gi]).trim() || "(Unassigned)",
      account: account,
      amount: Number(row[mi]) || 0,
    });
  }
  return out;
}

function requireColumn(headers: string[], name: string, tableName: string): number {
  const i = headers.indexOf(name);
  if (i < 0) {
    throw new Error(`Column "${name}" not found in ${tableName}. Found: ${headers.join(", ")}`);
  }
  return i;
}

/** Reads thresholds from tblSettings (Setting | Value); falls back to defaults. */
function readSettings(workbook: ExcelScript.Workbook): Settings {
  const s: Settings = {
    absThreshold: DEFAULT_ABS_THRESHOLD,
    pctThreshold: DEFAULT_PCT_THRESHOLD,
  };

  const table = workbook.getTable("tblSettings");
  if (!table || table.getRowCount() === 0) {
    return s;
  }

  for (const row of table.getRangeBetweenHeaderAndTotal().getValues()) {
    const k = String(row[0]).trim();
    const v = Number(row[1]);
    if (isNaN(v)) continue;
    if (k === "AbsThreshold") s.absThreshold = v;
    if (k === "PctThreshold") s.pctThreshold = v;
  }
  return s;
}

/** 0 -> A, 1 -> B, ... 26 -> AA */
function columnLetter(index: number): string {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
