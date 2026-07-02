import ExcelJS from 'exceljs';
import sharp from 'sharp';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── Column mapping: original A-M preserved exactly, RMB added at N ───────────
const CATEGORY_COL = {
  'Airfare':          5,   // E
  'Transportation':   6,   // F
  'Lodging':          7,   // G
  'Travel-Other':     8,   // H
  'Meals':            9,   // I
  'Entertainment':   10,   // J
  'Telephone':       11,   // K
  'Office Supplies': 12,   // L
  'Others':          13,   // M
  'RMB':             14,   // N ← new column
};

// Receipts tab: exact column widths from reference template (A-G)
const RECEIPT_COL_WIDTHS = [68.3, 71.3, 69.6, 74.3, 67.3, 64.0, 64.0];
const RECEIPT_ROW_HEIGHT = 408.6; // points
// Pixel conversion
const COL_PX = RECEIPT_COL_WIDTHS.map(w => Math.round(w * 7 + 5));
const ROW_PX = Math.round(RECEIPT_ROW_HEIGHT * 96 / 72);

const COL_LETTERS = ['A','B','C','D','E','F','G'];

function refLabel(idx) {
  return `${Math.floor(idx / 7) + 1}${COL_LETTERS[idx % 7]}`;
}

// Deep-copy a cell's style onto another cell
function copyStyle(srcCell, dstCell) {
  try {
    if (srcCell.font)      dstCell.font      = { ...srcCell.font };
    if (srcCell.border)    dstCell.border    = JSON.parse(JSON.stringify(srcCell.border));
    if (srcCell.fill)      dstCell.fill      = JSON.parse(JSON.stringify(srcCell.fill));
    if (srcCell.alignment) dstCell.alignment = { ...srcCell.alignment };
    if (srcCell.numFmt)    dstCell.numFmt    = srcCell.numFmt;
  } catch (e) {
    // silently skip style copy errors
  }
}

// Resize image buffer to fill cell while preserving aspect ratio
async function resizeToFill(buffer, cellW, cellH) {
  try {
    const meta = await sharp(buffer).metadata();
    const scale = Math.min(cellW / meta.width, cellH / meta.height);
    const newW = Math.round(meta.width * scale);
    const newH = Math.round(meta.height * scale);
    const resized = await sharp(buffer)
      .resize(newW, newH, { fit: 'fill' })
      .jpeg({ quality: 85 })
      .toBuffer();
    return { buffer: resized, width: newW, height: newH };
  } catch {
    return { buffer, width: cellW, height: cellH };
  }
}

export async function generateExcel(employeeName, receipts) {
  // Sort chronologically by date string
  const sorted = [...receipts].sort((a, b) => {
    if (!a.date) return 1;
    if (!b.date) return -1;
    return a.date.localeCompare(b.date);
  });

  // Assign Ref# (1A, 1B ... 1G, 2A, 2B ...)
  sorted.forEach((r, i) => { r.ref = refLabel(i); });

  const firstDate = sorted[0]?.date || '';
  const lastDate  = sorted[sorted.length - 1]?.date || '';
  const n = sorted.length;

  // ── Row layout (dynamic based on receipt count) ────────────────────────────
  const DATA_START  = 8;
  const DATA_END    = DATA_START + n - 1;
  const SUBTOTAL    = DATA_END + 1;
  const REQ_ROW     = SUBTOTAL + 1;
  const TOTAL_ROW   = SUBTOTAL + 3;
  const REVIEW_ROW  = SUBTOTAL + 4;
  const REIMB_ROW   = SUBTOTAL + 5;
  const APPROVAL    = SUBTOTAL + 7;

  // ── Load template ──────────────────────────────────────────────────────────
  const templatePath = path.join(__dirname, 'template', 'SamuelJuneJulylocalexpenses.xlsx');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const wsTWD = wb.getWorksheet('TWD');

  // ── Capture styles from reference BEFORE modifying anything ───────────────
  // Data row styles (row 8, cols 1-14; col 14 reuses col 13)
  const dataStyles = {};
  for (let c = 1; c <= 14; c++) {
    dataStyles[c] = wsTWD.getCell(8, Math.min(c, 13));
  }

  // Subtotal row styles (row 25)
  const subtotalStyles = {};
  for (let c = 1; c <= 14; c++) {
    subtotalStyles[c] = wsTWD.getCell(25, Math.min(c, 13));
  }

  // Header row 7 col 13 style (for RMB header N7)
  const header7N = wsTWD.getCell(7, 13);

  // Bottom section styles (rows 26-32)
  const bottomStyles = {};
  for (const refRow of [26, 27, 28, 29, 30, 31, 32]) {
    bottomStyles[refRow] = {};
    for (let c = 1; c <= 14; c++) {
      bottomStyles[refRow][c] = wsTWD.getCell(refRow, Math.min(c, 13));
    }
  }

  // Currency box value cell style (M2)
  const currencyValCell = wsTWD.getCell(2, 13);

  // ── Clear rows 8-50 ────────────────────────────────────────────────────────
  for (let r = 8; r <= 50; r++) {
    for (let c = 1; c <= 14; c++) {
      wsTWD.getCell(r, c).value = null;
    }
  }

  // ── Add RMB to Currency/Rate box at N2/N3 (L2/M2/L3/M3 unchanged) ─────────
  wsTWD.getColumn(14).width = 12.71;
  copyStyle(currencyValCell, wsTWD.getCell(2, 14));
  wsTWD.getCell(2, 14).value = 'RMB';
  copyStyle(currencyValCell, wsTWD.getCell(3, 14));
  wsTWD.getCell(3, 14).value = 4.33;

  // ── Add RMB header at N7 ───────────────────────────────────────────────────
  copyStyle(header7N, wsTWD.getCell(7, 14));
  wsTWD.getCell(7, 14).value = 'RMB';

  // ── Update employee name and period ───────────────────────────────────────
  wsTWD.getCell(5, 2).value = employeeName || 'Samuel Chiang';
  wsTWD.getCell(5, 10).value = `${firstDate}-${lastDate}`;

  // ── Write data rows ────────────────────────────────────────────────────────
  sorted.forEach((r, i) => {
    const row = DATA_START + i;
    for (let c = 1; c <= 14; c++) {
      copyStyle(dataStyles[c], wsTWD.getCell(row, c));
    }
    wsTWD.getRow(row).height = 18.75;

    wsTWD.getCell(row, 1).value = r.date;
    wsTWD.getCell(row, 2).value = r.description || r.merchant || '';
    wsTWD.getCell(row, 3).value = r.ref;
    wsTWD.getCell(row, 4).value = { formula: `SUM(E${row}:N${row})` };
    const catCol = CATEGORY_COL[r.category] || 13;
    wsTWD.getCell(row, catCol).value = Number(r.amount) || 0;
  });

  // ── Write subtotal row ─────────────────────────────────────────────────────
  for (let c = 1; c <= 14; c++) {
    copyStyle(subtotalStyles[c], wsTWD.getCell(SUBTOTAL, c));
  }
  wsTWD.getRow(SUBTOTAL).height = 18.75;
  wsTWD.getCell(SUBTOTAL, 3).value = 'total';
  wsTWD.getCell(SUBTOTAL, 4).value = { formula: `SUM(E${SUBTOTAL}:N${SUBTOTAL})` };
  for (let c = 5; c <= 14; c++) {
    const colLetter = String.fromCharCode(64 + c);
    wsTWD.getCell(SUBTOTAL, c).value = { formula: `SUM(${colLetter}${DATA_START}:${colLetter}${DATA_END})` };
  }

  // ── Write bottom section ───────────────────────────────────────────────────
  const bottomMap = {
    [REQ_ROW]:     26,
    [REQ_ROW + 1]: 27,
    [TOTAL_ROW]:   28,
    [REVIEW_ROW]:  29,
    [REIMB_ROW]:   30,
    [REIMB_ROW+1]: 31,
    [APPROVAL]:    32,
  };

  for (const [newRowStr, refRow] of Object.entries(bottomMap)) {
    const nr = parseInt(newRowStr);
    for (let c = 1; c <= 14; c++) {
      copyStyle(bottomStyles[refRow][c], wsTWD.getCell(nr, c));
    }
    wsTWD.getRow(nr).height = undefined; // auto-height, prevents squashing
  }

  wsTWD.getCell(REQ_ROW, 1).value = 'REQUESTED BY:';
  wsTWD.getCell(REQ_ROW, 2).value = employeeName || 'Samuel Chiang';
  wsTWD.getCell(REQ_ROW, 6).value = '*64020 Transportation – Other (example rental car, taxi etc)';

  wsTWD.getCell(TOTAL_ROW, 13).value = 'Total Expenses:';
  wsTWD.getCell(TOTAL_ROW, 14).value = { formula: `D${SUBTOTAL}` };
  wsTWD.getCell(TOTAL_ROW, 14).numFmt = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';

  wsTWD.getCell(REVIEW_ROW, 1).value = 'REVIEWED BY:';
  wsTWD.getCell(REVIEW_ROW, 2).value = 'DATE';
  wsTWD.getCell(REVIEW_ROW, 13).value = 'Less Advance';
  wsTWD.getCell(REVIEW_ROW, 14).value = 0;
  wsTWD.getCell(REVIEW_ROW, 14).numFmt = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';

  wsTWD.getCell(REIMB_ROW, 13).value = 'Total Reimbursable Expenses:';
  wsTWD.getCell(REIMB_ROW, 14).value = { formula: `N${TOTAL_ROW}-N${REVIEW_ROW}` };
  wsTWD.getCell(REIMB_ROW, 14).numFmt = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';
  wsTWD.getCell(REIMB_ROW, 14).border = {
    top:    { style: 'thin' },
    bottom: { style: 'double' },
  };

  wsTWD.getCell(APPROVAL, 1).value = 'APPROVAL BY:';
  wsTWD.getCell(APPROVAL, 2).value = 'DATE';

  // ── Receipts sheet ─────────────────────────────────────────────────────────
  const wsRec = wb.getWorksheet('Receipts');

  // Set column widths A-G to match reference
  RECEIPT_COL_WIDTHS.forEach((w, i) => {
    wsRec.getColumn(i + 1).width = w;
  });

  // Set row heights
  const numReceiptRows = Math.ceil(n / 7);
  for (let r = 1; r <= numReceiptRows; r++) {
    wsRec.getRow(r).height = RECEIPT_ROW_HEIGHT;
  }

  // Clear existing images
  if (wsRec._images) wsRec._images.length = 0;

  // Embed receipt images resized to fill each cell
  for (let i = 0; i < sorted.length; i++) {
    const r = sorted[i];
    if (!r.imageBase64) continue;

    const colIdx = i % 7;           // 0-based column
    const rowIdx = Math.floor(i / 7); // 0-based row

    const cellW = COL_PX[colIdx];
    const cellH = ROW_PX;

    try {
      const rawBuffer = Buffer.from(r.imageBase64, 'base64');
      const { buffer: resizedBuf, width: imgW, height: imgH } = await resizeToFill(rawBuffer, cellW, cellH);

      const imageId = wb.addImage({
        buffer: resizedBuf,
        extension: 'jpeg',
      });

      wsRec.addImage(imageId, {
        tl: { col: colIdx, row: rowIdx },
        ext: { width: imgW, height: imgH },
        editAs: 'oneCell',
      });
    } catch (e) {
      console.error(`Failed to embed image for ${r.ref}:`, e.message);
    }
  }

  // ── Return buffer ──────────────────────────────────────────────────────────
  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}
