import ExcelJS from 'exceljs';
import sharp from 'sharp';
import path from 'path';
import { fileURLToPath } from 'url';
import { normalizeDate, isIsoDate } from './dateUtils.js';
import { FX_MARKUP } from './fx.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── Column mapping: original A-M (13 cols), all amounts in TWD ───────────────
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
  'RMB':             13,   // M (treated as Others/TWD)
};

// Receipts tab: exact column widths from reference (A-G)
const RECEIPT_COL_WIDTHS = [68.3, 71.3, 69.6, 74.3, 67.3, 64.0, 64.0];
const RECEIPT_ROW_HEIGHT = 408.6; // points
const COL_PX = RECEIPT_COL_WIDTHS.map(w => Math.round(w * 7 + 5));
const ROW_PX = Math.round(RECEIPT_ROW_HEIGHT * 96 / 72);

const COL_LETTERS = ['A','B','C','D','E','F','G'];

function refLabel(idx) {
  return `${Math.floor(idx / 7) + 1}${COL_LETTERS[idx % 7]}`;
}

// Snapshot a cell's full style as a plain object (detached from the live cell)
function snapStyle(cell) {
  try {
    return JSON.parse(JSON.stringify(cell.style || {}));
  } catch {
    return {};
  }
}

// Replace a cell's entire style (clears anything the snapshot doesn't define)
function applyStyle(dstCell, style) {
  try {
    dstCell.style = JSON.parse(JSON.stringify(style));
  } catch {
    // silently skip style errors
  }
}

// Resize image buffer to fill cell while preserving aspect ratio.
// .rotate() with no argument bakes in the EXIF orientation — phone photos are
// usually stored sideways with a rotation tag that Excel ignores.
async function resizeToFill(buffer, cellW, cellH) {
  try {
    const meta = await sharp(buffer).metadata();
    const sideways = meta.orientation >= 5 && meta.orientation <= 8;
    const srcW = sideways ? meta.height : meta.width;
    const srcH = sideways ? meta.width : meta.height;
    const scale = Math.min(cellW / srcW, cellH / srcH);
    const newW = Math.round(srcW * scale);
    const newH = Math.round(srcH * scale);
    const resized = await sharp(buffer)
      .rotate()
      .resize(newW, newH, { fit: 'fill' })
      .jpeg({ quality: 85 })
      .toBuffer();
    return { buffer: resized, width: newW, height: newH };
  } catch {
    return { buffer, width: cellW, height: cellH };
  }
}

export async function generateExcel(employeeName, receipts) {
  // Normalize every date to ISO first (cached parses and hand edits can arrive
  // as 08Jul2026, 7/1/26, 115/4/28 ...), then sort chronologically
  const sorted = receipts.map(r => ({ ...r, date: normalizeDate(r.date) || '' })).sort((a, b) => {
    if (!a.date) return 1;
    if (!b.date) return -1;
    return a.date.localeCompare(b.date);
  });

  // Assign Ref# (1A, 1B ... 1G, 2A, 2B ...)
  sorted.forEach((r, i) => { r.ref = refLabel(i); });

  // Period spans only the receipts with a recognizable date
  const isoDates = sorted.map(r => r.date).filter(isIsoDate);
  const firstDate = isoDates[0] || '';
  const lastDate  = isoDates[isoDates.length - 1] || '';
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

  // ── Capture styles from reference rows BEFORE modifying anything ───────────
  // Snapshots must be detached plain objects: with 18+ receipts the data rows
  // land on the template's own rows 25-32, so live cell references would be
  // overwritten before the subtotal/bottom sections are written.
  const dataStyles = {};
  for (let c = 1; c <= 13; c++) {
    dataStyles[c] = snapStyle(wsTWD.getCell(8, c));
  }

  // Subtotal row styles (row 25)
  const subtotalStyles = {};
  for (let c = 1; c <= 13; c++) {
    subtotalStyles[c] = snapStyle(wsTWD.getCell(25, c));
  }

  // Bottom section styles (rows 26-32)
  const bottomStyles = {};
  for (const refRow of [26, 27, 28, 29, 30, 31, 32]) {
    bottomStyles[refRow] = {};
    for (let c = 1; c <= 13; c++) {
      bottomStyles[refRow][c] = snapStyle(wsTWD.getCell(refRow, c));
    }
  }

  // ── Clear rows 8-60: values AND styles, so template formatting from the
  //    original row positions doesn't bleed through at the new positions ──────
  for (let r = 8; r <= 60; r++) {
    for (let c = 1; c <= 20; c++) {
      const cell = wsTWD.getCell(r, c);
      cell.value = null;
      cell.style = {};
    }
  }

  // ── Update employee name and period ───────────────────────────────────────
  const slashDate = d => /^\d{4}-\d{2}-\d{2}/.test(d || '') ? d.substring(0, 10).replace(/-/g, '/') : (d || '');
  wsTWD.getCell(5, 2).value = employeeName || 'Samuel Chiang';
  wsTWD.getCell(5, 10).value = `${slashDate(firstDate)}-${slashDate(lastDate)}`;

  // ── Currency / Rate box (L2:M3): the marked-up rate(s) applied to foreign receipts
  const fxUsed = {};
  for (const r of sorted) {
    if (r.currency && r.currency !== 'TWD' && r.fxRate) {
      (fxUsed[r.currency] = fxUsed[r.currency] || new Set()).add(r.fxRate);
    }
  }
  const fxCurrencies = Object.keys(fxUsed);
  const fmtRates = c => [...fxUsed[c]].sort((a, b) => a - b).map(x => x.toFixed(2)).join(' / ');
  if (fxCurrencies.length === 0) {
    wsTWD.getCell(2, 13).value = 'TWD';
    wsTWD.getCell(3, 13).value = 'N/A';
  } else if (fxCurrencies.length === 1) {
    const c = fxCurrencies[0];
    wsTWD.getCell(2, 13).value = c;
    // a single rate stays numeric so the template's #,##0.00 format applies
    wsTWD.getCell(3, 13).value = fxUsed[c].size === 1 ? [...fxUsed[c]][0] : fmtRates(c);
  } else {
    wsTWD.getCell(2, 13).value = fxCurrencies.join(' / ');
    wsTWD.getCell(3, 13).value = fxCurrencies.map(c => `${c} ${fmtRates(c)}`).join(' / ');
  }

  // Fine print to the right of the box (N2) explaining the card-fee markup
  const note = wsTWD.getCell(2, 14);
  if (fxCurrencies.length > 0) {
    note.value = `* Rate = bank rate on the receipt date + ${(FX_MARKUP * 100).toFixed(1)}% credit card fee (foreign-currency receipts)`;
    note.font = { name: 'Arial', size: 8, italic: true, color: { argb: 'FF718096' } };
    note.alignment = { vertical: 'middle', horizontal: 'left' };
  } else {
    note.value = null;
  }

  // ── Write data rows ────────────────────────────────────────────────────────
  sorted.forEach((r, i) => {
    const row = DATA_START + i;
    for (let c = 1; c <= 13; c++) {
      applyStyle(wsTWD.getCell(row, c), dataStyles[c]);
    }
    wsTWD.getRow(row).height = 18.75;

    // Real date value, consistently formatted as m/d/yyyy (e.g. 8/7/2026)
    const dateCell = wsTWD.getCell(row, 1);
    if (isIsoDate(r.date)) {
      const [y, mo, d] = r.date.split('-').map(Number);
      dateCell.value = new Date(Date.UTC(y, mo - 1, d));
      dateCell.numFmt = 'm/d/yyyy';
    } else {
      dateCell.value = r.date || '';
    }

    // Explanation: English summary first, then the receipt's own wording,
    // e.g. "parking fee - 新竹停車場 停車費" or "hotel stay - Hyatt House San Jose"
    const merchant = (r.merchant || '').trim();
    const desc = (r.description || '').trim();
    const original = (r.original || '').trim();
    const local = original && original !== merchant && !merchant.includes(original)
      ? `${merchant} ${original}`.trim()
      : merchant;
    const explanation = desc && local && desc !== local ? `${desc} - ${local}` : (desc || local);
    // Foreign receipts: note the original amount and the rate used, e.g. "(USD 200.00 @ 32.76)"
    const isForeign = r.currency && r.currency !== 'TWD';
    const fxNote = isForeign
      ? ` (${r.currency} ${(Number(r.amount) || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} @ ${r.fxRate ? r.fxRate.toFixed(2) : 'rate unavailable'})`
      : '';
    wsTWD.getCell(row, 2).value = explanation + fxNote;
    wsTWD.getCell(row, 3).value = r.ref;
    wsTWD.getCell(row, 4).value = { formula: `SUM(E${row}:M${row})` };
    const catCol = CATEGORY_COL[r.category] || 13;
    wsTWD.getCell(row, catCol).value = Number(isForeign ? r.amountTwd : r.amount) || 0;
  });

  // ── Write subtotal row ─────────────────────────────────────────────────────
  for (let c = 1; c <= 13; c++) {
    applyStyle(wsTWD.getCell(SUBTOTAL, c), subtotalStyles[c]);
  }
  wsTWD.getRow(SUBTOTAL).height = 18.75;
  wsTWD.getCell(SUBTOTAL, 3).value = 'total';
  wsTWD.getCell(SUBTOTAL, 4).value = { formula: `SUM(E${SUBTOTAL}:M${SUBTOTAL})` };
  for (let c = 5; c <= 13; c++) {
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
    for (let c = 1; c <= 13; c++) {
      applyStyle(wsTWD.getCell(nr, c), bottomStyles[refRow][c]);
    }
    wsTWD.getRow(nr).height = undefined; // auto-height
  }

  wsTWD.getCell(REQ_ROW, 1).value = 'REQUESTED BY:';
  wsTWD.getCell(REQ_ROW, 2).value = employeeName || 'Samuel Chiang';
  wsTWD.getCell(REQ_ROW, 6).value = '*64020 Transportation – Other (example rental car, taxi etc)';

  const moneyFmt = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)';

  wsTWD.getCell(TOTAL_ROW, 12).value = 'Total Expenses:';
  wsTWD.getCell(TOTAL_ROW, 13).value = { formula: `D${SUBTOTAL}` };
  wsTWD.getCell(TOTAL_ROW, 13).numFmt = moneyFmt;

  wsTWD.getCell(REVIEW_ROW, 1).value = 'REVIEWED BY:';
  wsTWD.getCell(REVIEW_ROW, 2).value = 'DATE';
  wsTWD.getCell(REVIEW_ROW, 12).value = 'Less Advance';
  wsTWD.getCell(REVIEW_ROW, 13).value = 0;
  wsTWD.getCell(REVIEW_ROW, 13).numFmt = moneyFmt;

  wsTWD.getCell(REIMB_ROW, 12).value = 'Total Reimbursable Expenses:';
  wsTWD.getCell(REIMB_ROW, 13).value = { formula: `M${TOTAL_ROW}-M${REVIEW_ROW}` };
  wsTWD.getCell(REIMB_ROW, 13).numFmt = moneyFmt;
  wsTWD.getCell(REIMB_ROW, 13).border = {
    top:    { style: 'thin' },
    bottom: { style: 'double' },
  };

  wsTWD.getCell(APPROVAL, 1).value = 'APPROVAL BY:';
  wsTWD.getCell(APPROVAL, 2).value = 'DATE';

  // ── Receipts sheet ─────────────────────────────────────────────────────────
  const wsRec = wb.getWorksheet('Receipts');

  // Set column widths A-G
  RECEIPT_COL_WIDTHS.forEach((w, i) => {
    wsRec.getColumn(i + 1).width = w;
  });

  // Set row heights
  const numReceiptRows = Math.ceil(n / 7);
  for (let r = 1; r <= numReceiptRows; r++) {
    wsRec.getRow(r).height = RECEIPT_ROW_HEIGHT;
  }

  // Drop the template's own receipt images: clear the sheet anchors AND the
  // workbook media list, otherwise the old pictures ride along in every file
  if (wsRec._images) wsRec._images.length = 0;
  wb.media.length = 0;

  // Embed receipt images resized to fill each cell
  for (let i = 0; i < sorted.length; i++) {
    const r = sorted[i];
    if (!r.imageBase64) continue;

    const colIdx = i % 7;
    const rowIdx = Math.floor(i / 7);
    const cellW  = COL_PX[colIdx];
    const cellH  = ROW_PX;

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
