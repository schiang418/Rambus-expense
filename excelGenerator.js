import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Category → column index mapping (1-based, matching template)
const CATEGORY_COL = {
  'Airfare':         5,
  'Transportation':  6,
  'Lodging':         7,
  'Travel-Other':    8,
  'Meals':           9,
  'Entertainment':  10,
  'Telephone':      11,
  'Office Supplies':12,
  'Others':         13,
};

// Reference cell dimensions (from original template Receipts sheet)
const REF_COL_WIDTH  = 69;      // character units
const REF_ROW_HEIGHT = 408.6;   // points
// Pixel equivalents for image sizing
const CELL_W_PX = Math.round(REF_COL_WIDTH * 7 + 5);   // ≈ 488
const CELL_H_PX = Math.round(REF_ROW_HEIGHT * 96 / 72); // ≈ 545

export async function generateExcel(employeeName, receipts) {
  // Sort chronologically
  const sorted = [...receipts].sort((a, b) => {
    if (!a.date) return 1;
    if (!b.date) return -1;
    return a.date.localeCompare(b.date);
  });

  // Assign Ref# (max 7 per row, A-G)
  sorted.forEach((r, i) => {
    r.ref = `${Math.floor(i / 7) + 1}${String.fromCharCode(65 + (i % 7))}`;
  });

  const firstDate = sorted[0]?.date || '';
  const lastDate  = sorted[sorted.length - 1]?.date || '';

  // ─── Load template ────────────────────────────────────────────────────────
  const templatePath = path.join(__dirname, 'template', 'SamuelJuneJulylocalexpenses.xlsx');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);

  // ─── TWD sheet ────────────────────────────────────────────────────────────
  const wsTWD = wb.getWorksheet('TWD');

  // Employee name (row 5, col 2)
  wsTWD.getCell(5, 2).value = employeeName;

  // For Period (row 5, col 10)
  wsTWD.getCell(5, 10).value = `${firstDate} to ${lastDate}`;

  // Clear data rows 8-24
  for (let row = 8; row <= 24; row++) {
    for (let col = 1; col <= 15; col++) {
      const cell = wsTWD.getCell(row, col);
      // Preserve formulas in subtotal/total rows
      if (typeof cell.value !== 'object' || !cell.value?.formula) {
        cell.value = null;
      }
    }
  }

  // Write expense rows
  sorted.forEach((r, i) => {
    const row = 8 + i;
    wsTWD.getCell(row, 1).value = r.date;
    wsTWD.getCell(row, 2).value = `${r.category} – ${r.merchant}`;
    wsTWD.getCell(row, 3).value = r.ref;
    const catCol = CATEGORY_COL[r.category] || 13;
    wsTWD.getCell(row, catCol).value = Number(r.amount) || 0;
  });

  // ─── Receipts sheet ───────────────────────────────────────────────────────
  const wsReceipts = wb.getWorksheet('Receipts');

  // Set column widths and row heights to match reference
  for (let c = 1; c <= 7; c++) {
    wsReceipts.getColumn(c).width = REF_COL_WIDTH;
  }
  const maxRow = Math.ceil(sorted.length / 7) + 1;
  for (let r = 1; r <= maxRow; r++) {
    wsReceipts.getRow(r).height = REF_ROW_HEIGHT;
  }

  // Remove existing images
  wsReceipts.getImages().forEach(img => {
    const idx = wsReceipts._images?.indexOf(img);
    if (idx > -1) wsReceipts._images.splice(idx, 1);
  });
  // Clear via workbook image list
  if (wb._media) {
    // keep non-receipt images; we'll just add new ones
  }

  // Embed receipt images
  for (const r of sorted) {
    if (!r.imageBase64) continue;

    const rowNum = parseInt(r.ref[0]);
    const colNum = r.ref.charCodeAt(1) - 65; // 0-based

    // Scale image to fill cell
    const imgBuffer = Buffer.from(r.imageBase64, 'base64');
    
    // Get image dimensions using ExcelJS approach
    // We'll embed at cell size directly
    const imageId = wb.addImage({
      base64: r.imageBase64,
      extension: 'jpeg',
    });

    // Use twoCell anchor so image fills the cell
    wsReceipts.addImage(imageId, {
      tl: { col: colNum, row: rowNum - 1 },
      br: { col: colNum + 1, row: rowNum },
      editAs: 'oneCell',
    });
  }

  // ─── Return buffer ────────────────────────────────────────────────────────
  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}
