import * as XLSX from "xlsx";
import { convertNumber, type ConvertOptions } from "./numberToWords";

export interface ColumnInfo {
  index: number;
  header: string;
  numericRatio: number;
  isCurrencyLike: boolean;
}

export interface DetectionResult {
  headerRow: number;
  columns: ColumnInfo[];
}

const CURRENCY_KEYWORDS = ["amount", "total", "price", "cost", "value", "cash", "payable", "rupees", "rs", "inr"];

function isNumericValue(cell: XLSX.CellObject | undefined): boolean {
  if (!cell || cell.v === undefined || cell.v === null || cell.v === "") return false;
  if (typeof cell.v === "number") return true;
  if (cell.t === "n") return true;
  const s = String(cell.v).trim().replace(/,/g, "");
  return s !== "" && !isNaN(Number(s)) && /^-?\d+(\.\d+)?$/.test(s);
}

/** Detect the header row by scanning for the row with the most non-empty text cells
 *  whose next rows contain numeric data. Returns range.s.r if nothing better is found. */
export function detectHeaderRow(ws: XLSX.WorkSheet): number {
  if (!ws["!ref"]) return 0;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const maxScan = Math.min(range.e.r, range.s.r + 30);

  let best = { row: range.s.r, score: -1 };
  for (let R = range.s.r; R <= maxScan; R++) {
    let textCells = 0;
    let numericCellsInRow = 0;
    let nonEmpty = 0;
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell || cell.v === undefined || cell.v === null || cell.v === "") continue;
      nonEmpty++;
      if (isNumericValue(cell)) numericCellsInRow++;
      else if (typeof cell.v === "string" && cell.v.trim().length > 0) textCells++;
    }
    if (textCells < 2 || numericCellsInRow > textCells) continue;

    // Look at next few rows — must have numeric data
    let belowNumeric = 0;
    const lookahead = Math.min(R + 5, range.e.r);
    for (let R2 = R + 1; R2 <= lookahead; R2++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        if (isNumericValue(ws[XLSX.utils.encode_cell({ r: R2, c: C })])) belowNumeric++;
      }
    }
    if (belowNumeric === 0) continue;

    const score = textCells * 10 + nonEmpty + belowNumeric;
    if (score > best.score) best = { row: R, score };
  }
  return best.row;
}

export function detectNumericColumns(ws: XLSX.WorkSheet, headerRow?: number): DetectionResult {
  if (!ws["!ref"]) return { headerRow: 0, columns: [] };
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const hr = headerRow ?? detectHeaderRow(ws);
  const cols: ColumnInfo[] = [];

  for (let C = range.s.c; C <= range.e.c; C++) {
    const headerCell = ws[XLSX.utils.encode_cell({ r: hr, c: C })];
    const header = headerCell ? String(headerCell.v ?? "").trim() : "";
    if (!header) continue;
    if (/in words$/i.test(header)) continue;

    let numeric = 0;
    let total = 0;
    for (let R = hr + 1; R <= range.e.r; R++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
      if (!cell || cell.v === undefined || cell.v === null || cell.v === "") continue;
      total++;
      if (isNumericValue(cell)) numeric++;
    }
    const ratio = total === 0 ? 0 : numeric / total;
    const isCurrencyLike = CURRENCY_KEYWORDS.some((k) => header.toLowerCase().includes(k));
    if (ratio >= 0.6 && total > 0) {
      cols.push({ index: C, header, numericRatio: ratio, isCurrencyLike });
    }
  }
  return { headerRow: hr, columns: cols };
}

/** Insert a new column at position `insertAt` containing converted words for column `sourceCol`. */
function insertWordsColumn(
  ws: XLSX.WorkSheet,
  sourceCol: number,
  insertAt: number,
  newHeader: string,
  opts: ConvertOptions,
): { skipped: number } {
  const range = XLSX.utils.decode_range(ws["!ref"]!);
  let skipped = 0;

  // Shift all cells at column >= insertAt one column to the right.
  // Iterate from rightmost column to avoid overwriting.
  for (let C = range.e.c; C >= insertAt; C--) {
    for (let R = range.s.r; R <= range.e.r; R++) {
      const oldAddr = XLSX.utils.encode_cell({ r: R, c: C });
      const newAddr = XLSX.utils.encode_cell({ r: R, c: C + 1 });
      if (ws[oldAddr]) {
        ws[newAddr] = ws[oldAddr];
        delete ws[oldAddr];
      } else {
        delete ws[newAddr];
      }
    }
  }

  // Update merges
  if (ws["!merges"]) {
    for (const m of ws["!merges"]) {
      if (m.s.c >= insertAt) m.s.c += 1;
      if (m.e.c >= insertAt) m.e.c += 1;
    }
  }

  // Update column widths/info
  if (ws["!cols"]) {
    ws["!cols"].splice(insertAt, 0, { wch: Math.max(newHeader.length + 4, 30) });
  }

  // Write header at insertAt
  const headerAddr = XLSX.utils.encode_cell({ r: range.s.r, c: insertAt });
  ws[headerAddr] = { t: "s", v: newHeader };

  // Write converted values. Note: sourceCol may have shifted if sourceCol >= insertAt (it was, since insertAt = sourceCol+1, so source unchanged).
  for (let R = range.s.r + 1; R <= range.e.r; R++) {
    const srcAddr = XLSX.utils.encode_cell({ r: R, c: sourceCol });
    const cell = ws[srcAddr];
    const newAddr = XLSX.utils.encode_cell({ r: R, c: insertAt });
    if (!cell || cell.v === undefined || cell.v === null || cell.v === "") {
      skipped++;
      continue;
    }
    const num = typeof cell.v === "number" ? cell.v : Number(cell.v);
    if (isNaN(num)) {
      skipped++;
      continue;
    }
    ws[newAddr] = { t: "s", v: convertNumber(num, opts) };
  }

  // Update !ref
  range.e.c += 1;
  ws["!ref"] = XLSX.utils.encode_range(range);

  return { skipped };
}

export function processSheet(
  ws: XLSX.WorkSheet,
  selectedCols: number[],
  opts: ConvertOptions,
  headerRow?: number,
): { skipped: number } {
  const hr = headerRow ?? detectHeaderRow(ws);
  const sorted = [...selectedCols].sort((a, b) => b - a);
  let skipped = 0;
  for (const col of sorted) {
    const headerAddr = XLSX.utils.encode_cell({ r: hr, c: col });
    const headerVal = ws[headerAddr]?.v ?? "Column";
    const newHeader = `${String(headerVal)} in Words`;

    const nextAddr = XLSX.utils.encode_cell({ r: hr, c: col + 1 });
    if (ws[nextAddr] && /in words$/i.test(String(ws[nextAddr].v ?? ""))) continue;

    const r = insertWordsColumn(ws, col, col + 1, newHeader, opts, hr);
    skipped += r.skipped;
  }
  return { skipped };
}
