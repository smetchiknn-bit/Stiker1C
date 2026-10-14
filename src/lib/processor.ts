import type * as ExcelJS from "exceljs";
import { getExcelJS } from "./excel";

/* ... (шапка без изменений) ... */

export type Raw = string | number | boolean | Date | null;

/* ---------- цветовые константы макроса ---------- */
export const FILL_RED = "FF0000";
export const FILL_GREEN = "92D050";
export const FILL_GOLD = "FFC000";
export const FILL_YELLOW = "FFFF00";
export const FILL_GR = { theme: 4, tint: 0.599993896298105 };
export const FILL_L3 = { theme: 9, tint: 0.799981688894314 };
export const FILL_GR_HEX = "DEEBF7";
export const FILL_L3_HEX = "E2EFDA";
export const BORDER_GRAY = { theme: 1, tint: -0.349986266670736 };

/* Колонки, удаляемые в конце Нумерация_1С (исходные H,I,J + M..R). */
const DELETED = new Set([8, 9, 10, 13, 14, 15, 16, 17, 18]);
/* S (19) — колонка "ИД.Бюджетная.Статья.1С"; её НЕ удаляем, если она есть. */

const CODE_SRC_COL = 11; // K — код уровня в исходной сетке

/* ---------- типы ---------- */
interface CellData { /* ... без изменений ... */ }
type Row = Map<number, CellData>;
export interface MergeRect { r1: number; c1: number; r2: number; c2: number; }

export interface Analysis {
  fileName: string;
  baseName: string;
  outName: string;
  sheetNames: string[];
  dataRows: number;
  stikerRaw: Raw;
  stikerNum: number | null;
  hasGR: boolean;
  levelCounts: Record<string, number>;
  hasBudgetColumn: boolean;
  budgetSrcCol: number | null;  // 1-based номер колонки бюджетной статьи в исходной сетке
}

/* ... остальные интерфейсы без изменений ... */

/* ---------- утилиты ---------- */
const round2 = (x: number) => Math.round((x + Number.EPSILON) * 100) / 100;

export function money(n: number | null | undefined): string { /* ... */ }
export function num(v: Raw): number { /* ... */ }
function toNumericOrNull(v: Raw): number | null { /* ... */ }
function effectiveValue(v: ExcelJS.CellValue | null | undefined): Raw { /* ... */ }
function isEmptyV(v: Raw | undefined): boolean { return v == null || v === ""; }
export function disp(v: Raw): string { /* ... */ }

const OFFICE_THEME = [ /* ... */ ];
function themeToHex(idx: number, tint?: number): string { /* ... */ }
function a1ToRC(a: string): { row: number; col: number } | null { /* ... */ }
function styleFillHex(style: any): string | undefined { /* ... */ }
function findSheet(wb: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined { /* ... */ }
function baseNameOf(fileName: string): string { /* ... */ }

/* ---------- 1. анализ донора ---------- */
const normalizeHeader = (s: string) => s.toLowerCase().replace(/[\s._\-]/g, "");

/** Ищет 1-based номер колонки "ИД.Бюджетная.Статья.1С" по заголовку в строке 1. */
function findBudgetColumnIndex(ws: ExcelJS.Worksheet): number | null {
  const cols = Math.max(ws.columnCount, 1);
  for (let c = 1; c <= cols; c++) {
    const v = effectiveValue(ws.getCell(1, c).value);
    if (typeof v !== "string") continue;
    const h = normalizeHeader(v);
    if (h.includes("ид") && h.includes("бюджетнаястатья")) return c;
  }
  return null;
}

export function analyzeDonor(wb: ExcelJS.Workbook, fileName: string): Analysis {
  const sheetNames = wb.worksheets.map((w) => w.name);
  const ws = findSheet(wb, "Свод");
  if (!ws) { /* ... ошибка как была ... */ }

  const lastRowK = lastNonEmptyRowOf(ws, 11);
  const lastRowA = lastNonEmptyRowOf(ws, 1);
  const stikerRaw = effectiveValue(ws.getCell(2, 10).value);

  let hasGR = false;
  const levelCounts: Record<string, number> = {};
  for (let r = 2; r <= lastRowK; r++) {
    const v = effectiveValue(ws.getCell(r, 11).value);
    if (typeof v === "string" && v !== "") {
      levelCounts[v] = (levelCounts[v] || 0) + 1;
      if (v === "ГР") hasGR = true;
    }
  }

  const budgetSrcCol = findBudgetColumnIndex(ws);
  const hasBudgetColumn = budgetSrcCol != null;

  return {
    fileName,
    baseName: baseNameOf(fileName),
    outName: baseNameOf(fileName) + "_1С.xlsx",
    sheetNames,
    dataRows: Math.max(lastRowA - 1, 0),
    stikerRaw,
    stikerNum: stikerRaw == null || stikerRaw === "" ? 0 : toNumericOrNull(stikerRaw),
    hasGR,
    levelCounts,
    hasBudgetColumn,
    budgetSrcCol,
  };
}

function lastNonEmptyRowOf(ws: ExcelJS.Worksheet, col: number): number { /* ... */ }

/* ---------- доступ к сетке ---------- */
function cell(row: Row | undefined, c: number): CellData | undefined { return row?.get(c); }
function val(grid: Row[], r: number, c: number): Raw { return cell(grid[r - 1], c)?.v ?? null; }
function setV(row: Row, c: number, v: Raw): CellData { /* ... */ }
function fillRow(row: Row, hex: string) { /* ... */ }
function fillRowTheme(row: Row, t: { theme: number; tint: number }) { /* ... */ }
function lastNonEmpty(grid: Row[], c: number): number { /* ... */ }

/* ---------- 2. Нумерация_1С ---------- */
const HIERARCHY = [ /* ... */ ];
function runNumbering(grid: Row[], inserts: number[]): { inserted: number; nElemets: number } { /* ... без изменений ... */ }

/* ---------- 3. удаление служебных колонок ---------- */
function deletedBeforeIn(set: Set<number>, c: number): number {
  let n = 0;
  set.forEach((d) => { if (d < c) n++; });
  return n;
}
function mapColIn(set: Set<number>, c: number): number {
  return c - deletedBeforeIn(set, c);
}

function deleteColumns(
  grid: Row[],
  widths: any[],
  merges: MergeRect[],
  inserts: number[],
  deletedSet: Set<number>,
) {
  for (const row of grid) {
    const entries = [...row.entries()].filter(([c]) => !deletedSet.has(c));
    row.clear();
    for (const [c, cd] of entries) row.set(mapColIn(deletedSet, c), cd);
  }
  const newWidths: any[] = [];
  widths.forEach((w, i) => {
    const c = i + 1;
    if (!deletedSet.has(c)) newWidths[mapColIn(deletedSet, c) - 1] = w;
  });
  widths.length = 0;
  widths.push(...newWidths);
  for (const m of merges) {
    for (const ins of inserts) {
      if (m.r1 >= ins) m.r1++;
      if (m.r2 >= ins) m.r2++;
    }
  }
}

/* ---------- 4. Сквозные_формулы_1С ---------- */
function setFormula(grid: Row[], r: number, c: number, f: string, fr: number) { /* ... */ }
function runFormulas(grid: Row[], nElement: number) { /* ... */ }

/* ---------- 5. Формат_1С ---------- */
function runFormat(grid: Row[], nElemets: number, hasBudgetColumn: boolean, budgetColAfter: number | null) {
  for (let r = 1; r <= nElemets; r++) {
    while (grid.length < r) grid.push(new Map());
    const row = grid[r - 1];
    const g7 = row.get(7) ?? row.set(7, { v: null }).get(7)!;
    g7.fill = FILL_RED;
    g7.fillT = undefined;
    if (!g7.numFmt) g7.numFmt = "@";

    // Зелёная заливка: колонки ТА (8), КЕР.в.СР (9) + бюджетная колонка, если есть
    const greenCols = new Set<number>([8, 9]);
    if (budgetColAfter != null) greenCols.add(budgetColAfter);

    for (const c of greenCols) {
      const cd = row.get(c) ?? row.set(c, { v: null }).get(c)!;
      cd.fill = FILL_GREEN;
      cd.fillT = undefined;
      cd.gBorder = true;
    }
    const b = row.get(2) ?? row.set(2, { v: null }).get(2)!;
    b.numFmt = "0_ ;-0 ";
    if (r >= 2) b.hAlign = "left";
  }
}

/* ---------- 6. группировка_1С ---------- */
function runGrouping(grid: Row[], nElement: number, codeColAfter: number) {
  const lastRow = nElement;
  const dots: number[] = new Array(lastRow + 1).fill(0);
  for (let r = 2; r <= lastRow; r++) {
    const v = val(grid, r, 1);
    const s = v == null ? "" : String(v);
    dots[r] = s.length - s.split(".").join("").length;
  }
  const levels = new Array(lastRow + 1).fill(0);
  let groups = 0;
  for (let level = 10; level >= 3; level--) {
    let i = 2;
    while (i <= lastRow) {
      if (dots[i] === level) {
        const y0 = i;
        i++;
        while (i <= lastRow) {
          const isGR = val(grid, i, codeColAfter) === "ГР";
          if (dots[i] > level - 1 || isGR) i++;
          else break;
        }
        if (y0 < i) {
          groups++;
          for (let rr = y0; rr <= i - 1; rr++) levels[rr] = Math.min(7, levels[rr] + 1);
        }
      } else i++;
    }
  }
  return { levels, groups };
}

/* ---------- 7. сборка выходной книги ---------- */
const GRAY_BORDER = { style: "thin", color: BORDER_GRAY } as any;
function writeCell(target: ExcelJS.Cell, cd: CellData) { /* ... */ }
function buildWorkbook(/* ... */) { /* ... без изменений ... */ }

/* ---------- превью ---------- */
function previewFillHex(cd: CellData): string | undefined { /* ... */ }
function buildPreview(grid: Row[], levels: number[], upto: number): PreviewRow[] { /* ... */ }
function safePreview(grid: Row[], levels: number[], upto: number): PreviewRow[] { /* ... */ }

/* ---------- 8. ФИНАЛЬНАЯ обработка колонки [ИД.Бюджетная.Статья.1С] ---------- */
/**
 * Продлевает значения [ИД.Бюджетная.Статья.1С] со строк КЕР на строки ТМЦ
 * и очищает ячейки на всех прочих строках.
 *  - budgetCol — номер колонки бюджетной статьи в итоговой сетке (1-based)
 *  - codeCol   — номер колонки с кодами уровней (КЕР/ТМЦ/ГР) в итоговой сетке
 *  - upToRow   — последняя строка данных
 */
function propagateBudgetColumn(
  grid: Row[],
  budgetCol: number,
  codeCol: number,
  upToRow: number,
): void {
  let lastBudget: Raw = null;
  const limit = Math.min(upToRow, grid.length);
  for (let r = 2; r <= limit; r++) {
    const row = grid[r - 1];
    if (!row) continue;

    const code = row.get(codeCol)?.v;
    const codeS = typeof code === "string" ? code : null;

    if (codeS === "КЕР") {
      const cur = row.get(budgetCol)?.v;
      if (!isEmptyV(cur)) lastBudget = cur;
    }

    if (codeS === "КЕР" || codeS === "ТМЦ") {
      if (!isEmptyV(lastBudget)) {
        setV(row, budgetCol, lastBudget);
      } else if (row.has(budgetCol)) {
        setV(row, budgetCol, null);
      }
    } else {
      if (row.has(budgetCol)) {
        setV(row, budgetCol, null);
      }
    }
  }
}

/* ---------- главный вход ---------- */
export async function processWorkbook(wb: ExcelJS.Workbook, a: Analysis): Promise<Outcome> {
  const E = await getExcelJS();
  const ws = findSheet(wb, "Свод")!;
  const R = Math.max(ws.rowCount, 1);
  const C = Math.max(ws.columnCount, 12);

  // материализация сетки
  const grid: Row[] = [];
  const heights: (number | undefined)[] = [];
  for (let r = 1; r <= R; r++) {
    const row = new Map<number, CellData>();
    const xrow = ws.getRow(r);
    heights.push(xrow.height);
    xrow.eachCell({ includeEmpty: false }, (xcell, col) => {
      row.set(col, {
        v: effectiveValue(xcell.value),
        style: JSON.parse(JSON.stringify(xcell.style ?? {})),
      });
    });
    grid.push(row);
  }

  const widths: any[] = [];
  for (let c = 1; c <= Math.max(C, 19); c++) {
    const col = ws.getColumn(c);
    widths.push({ w: (col as any).isCustomWidth ? col.width : undefined, hidden: !!col.hidden });
  }

  const merges: MergeRect[] = [];
  const rawMerges: any[] = ((ws.model as any)?.merges as any[]) ?? [];
  for (const s of rawMerges) {
    const str = typeof s === "string" ? s : typeof s?.range === "string" ? s.range : null;
    if (!str) continue;
    const parts = str.split(":");
    if (parts.length !== 2) continue;
    const p1 = a1ToRC(parts[0]);
    const p2 = a1ToRC(parts[1]);
    if (p1 && p2) merges.push({ r1: p1.row, c1: p1.col, r2: p2.row, c2: p2.col });
  }

  // заголовки 1С
  const headers = [
    "ИД 1С",
    "Структура/Статья ССР",
    "ЕдИзм",
    "Количество",
    "СМР на 1 ед. руб. (с НДС)",
    "ТМЦ на 1 ед. руб. (с НДС)",
    "Код ССР, ИД КЕР, ИД ТМЦ",
  ];
  if (grid.length === 0) grid.push(new Map());
  headers.forEach((h, i) => setV(grid[0], i + 1, h));

  // динамический набор удаляемых колонок: не удаляем бюджетную колонку
  const deletedSet = new Set(DELETED);
  if (a.budgetSrcCol != null) deletedSet.delete(a.budgetSrcCol);

  // куда переедет бюджетная колонка и колонка кодов после удаления
  const budgetColAfter = a.budgetSrcCol != null ? mapColIn(deletedSet, a.budgetSrcCol) : null;
  const codeColAfter = mapColIn(deletedSet, CODE_SRC_COL);

  if (budgetColAfter != null) {
    setV(grid[0], budgetColAfter, "ИД.Бюджетная.Статья.1С");
  }

  // Нумерация_1С
  const lastRowA0 = lastNonEmpty(grid, 1);
  let y = 0;
  for (let rr = 2; rr <= lastRowA0; rr++) if (val(grid, rr, CODE_SRC_COL) === "КЕР") y++;
  const inserts: number[] = [];
  const numInfo = runNumbering(grid, inserts);
  const nElemets = numInfo.nElemets;

  // удаление служебных колонок
  deleteColumns(grid, widths, merges, inserts, deletedSet);

  const nElement = lastNonEmpty(grid, 1);

  // формулы, формат, группировка
  const sums = runFormulas(grid, nElement);
  runFormat(grid, nElemets, a.hasBudgetColumn, budgetColAfter);
  const grp = runGrouping(grid, nElement, codeColAfter);

  // ФИНАЛЬНЫЙ ШАГ: продлеваем бюджетную колонку КЕР → ТМЦ
  if (budgetColAfter != null) {
    propagateBudgetColumn(grid, budgetColAfter, codeColAfter, nElement);
  }

  // фильтруем объединения
  const okMerges = merges
    .filter((m) => !deletedSet.has(m.c1) && !deletedSet.has(m.c2))
    .map((m) => ({
      r1: m.r1,
      c1: mapColIn(deletedSet, m.c1),
      r2: m.r2,
      c2: mapColIn(deletedSet, m.c2),
    }))
    .filter((m) => m.c1 <= m.c2 && m.r1 <= m.r2);

  const outWorkbook = buildWorkbook(E, grid, heights, widths, okMerges, grp.levels);

  const totalNum = sums.grand;
  const diff = a.stikerNum != null ? round2(a.stikerNum - totalNum) : null;
  const match = diff == null ? null : Math.abs(diff) < 0.005;

  return {
    analysis: a,
    outWorkbook,
    totalNum,
    diff,
    match,
    rows: safePreview(grid, grp.levels, Math.min(sums.totalRow, 90)),
    cols: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
    stats: {
      totalRows: nElement,
      insertedRows: y,
      groups: grp.groups,
      ker: y,
      tmz: a.levelCounts["ТМЦ"] ?? 0,
    },
  };
}
