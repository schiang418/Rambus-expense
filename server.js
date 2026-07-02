import express from 'express';
import multer from 'multer';
import AdmZip from 'adm-zip';
import axios from 'axios';
import path from 'path';
import { fileURLToPath } from 'url';
import { v4 as uuidv4 } from 'uuid';
import { generateExcel } from './excelGenerator.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '100mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Multer: accept multiple files, store in memory
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 }   // 100 MB per file
});

const IMAGE_EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tiff', '.tif']);

const MIME_BY_EXT = {
  '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png',
  '.webp': 'image/webp', '.gif': 'image/gif', '.bmp': 'image/bmp',
  '.tiff': 'image/tiff', '.tif': 'image/tiff'
};

function isImage(filename) {
  return IMAGE_EXTS.has(path.extname(filename).toLowerCase());
}

function mimeFor(filename) {
  return MIME_BY_EXT[path.extname(filename).toLowerCase()] || 'image/jpeg';
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ─── ROC date converter ───────────────────────────────────────────────────────
function convertRocToWestern(dateStr) {
  if (!dateStr) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) return dateStr.substring(0, 10);
  // Years > 1000 are already western (receipts sometimes print 2026/06/16) — only
  // small years are ROC and need +1911
  const yr = y => (+y > 1000 ? +y : +y + 1911);
  const pad = n => String(+n).padStart(2, '0');
  let m = dateStr.match(/(\d+)年(\d{1,2})-\d{1,2}月/);           // 115年03-04月 (billing range)
  if (m) return `${yr(m[1])}-${pad(m[2])}-01`;
  m = dateStr.match(/(\d+)年(\d{1,2})月(\d{1,2})[日號]?/);        // 115年04月28日
  if (m) return `${yr(m[1])}-${pad(m[2])}-${pad(m[3])}`;
  m = dateStr.match(/(\d+)年(\d{1,2})月/);                        // 115年04月 (day unknown)
  if (m) return `${yr(m[1])}-${pad(m[2])}-01`;
  const CN = {零:0,一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10,廿:20,卅:30};
  m = dateStr.match(/(\d+)年([一二三四五六七八九十廿卅]+)月([一二三四五六七八九十廿卅]+)[號日]/);
  if (m) {
    // 十/廿/卅 are bases: 廿八 = 20+8, 十五 = 10+5, 四 = 4
    const cn2int = s => [...s].reduce((v, c) => (CN[c] >= 10 ? v + CN[c] : v + (CN[c] || 0)), 0);
    return `${yr(m[1])}-${pad(cn2int(m[2]))}-${pad(cn2int(m[3]))}`;
  }
  m = dateStr.match(/(\d+)\/(\d{1,2})\/(\d{1,2})/);               // 115/4/28 or 2026/06/16
  if (m) return `${yr(m[1])}-${pad(m[2])}-${pad(m[3])}`;
  return dateStr;
}

// ─── LLM vision: parse one receipt image (with retry + exponential backoff) ───
async function parseReceiptWithLLM(imageBase64, mimeType = 'image/jpeg', retries = 3) {
  const apiUrl = (process.env.FORGE_API_URL || 'https://api.openai.com').replace(/\/$/, '');
  const apiKey = process.env.FORGE_API_KEY || process.env.OPENAI_API_KEY;
  const model = process.env.FORGE_MODEL || 'gpt-5.4';

  const prompt = `You are an expert at reading Taiwan receipts (both traditional Chinese and English).
Analyze this receipt image and extract the following information.
Return ONLY a valid JSON object with exactly these fields (no markdown, no explanation):
{
  "date": "the date on the receipt as printed (e.g. 2026-04-14 or 115年04月14日 or 115/4/14)",
  "amount": <number, total amount paid in TWD, no currency symbol>,
  "merchant": "merchant/business name",
  "category": "one of: Airfare | Transportation | Lodging | Travel-Other | Meals | Entertainment | Telephone | Office Supplies | Others",
  "description": "brief description of the expense (10 words max)"
}

Category rules:
- Meals: restaurants, cafes, food, catering, 餐廳, 美食, 飲食
- Transportation: taxi, bus, HSR, MRT, parking, 計程車, 停車, 交通
- Office Supplies: printing, copying, stationery, 列印, 影印, 文具, 輸出
- Lodging: hotel, 飯店, 旅館
- Airfare: airline, 航空
- Telephone: phone bill, SIM, 電信
- Entertainment: client entertainment, events
- Others: anything else`;

  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const response = await axios.post(
        `${apiUrl}/v1/chat/completions`,
        {
          model,
          messages: [{
            role: 'user',
            content: [
              { type: 'image_url', image_url: { url: `data:${mimeType};base64,${imageBase64}`, detail: 'high' } },
              { type: 'text', text: prompt }
            ]
          }],
          // GPT-5-family models reject `max_tokens` and non-default `temperature`;
          // reasoning tokens also count toward the completion limit, so keep headroom.
          max_completion_tokens: 2000
        },
        {
          headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
          timeout: 60000
        }
      );
      const content = response.data.choices[0].message.content.trim();
      const cleaned = content.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();
      return JSON.parse(cleaned);
    } catch (err) {
      const status = err.response?.status;
      if ((status === 429 || status === 500 || status === 503 || err.code === 'ECONNABORTED') && attempt < retries) {
        const retryAfter = Number(err.response?.headers?.['retry-after']);
        const wait = retryAfter > 0 ? retryAfter * 1000 : attempt * 10000;
        console.warn(`Attempt ${attempt}/${retries} failed (${status || err.code}), retrying in ${wait/1000}s...`);
        await sleep(wait);
      } else {
        // Surface the API's actual error message, not just "status code NNN"
        const apiMsg = err.response?.data?.error?.message;
        if (apiMsg) err.message = `${apiMsg} (HTTP ${status})`;
        throw err;
      }
    }
  }
}

// ─── POST /api/parse ──────────────────────────────────────────────────────────
app.post('/api/parse', upload.array('files', 200), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    // Collect all images from every uploaded file
    const images = [];

    for (const file of req.files) {
      const ext = path.extname(file.originalname).toLowerCase();

      if (ext === '.zip') {
        const zip = new AdmZip(file.buffer);
        for (const entry of zip.getEntries()) {
          if (!entry.isDirectory && isImage(entry.entryName)) {
            const baseName = path.basename(entry.entryName);
            if (baseName.startsWith('._') || baseName.startsWith('.')) continue;
            images.push({
              name: entry.entryName,
              buffer: entry.getData(),
              mimeType: mimeFor(entry.entryName)
            });
          }
        }
      } else if (isImage(file.originalname)) {
        images.push({
          name: file.originalname,
          buffer: file.buffer,
          mimeType: file.mimetype || 'image/jpeg'
        });
      }
    }

    if (images.length === 0) {
      return res.status(400).json({ error: 'No receipt images found in the uploaded files' });
    }

    console.log(`Processing ${images.length} receipt image(s) from ${req.files.length} uploaded file(s)`);

    // Parse images in small parallel batches to stay under API rate limits
    const results = [];
    const BATCH = Number(process.env.PARSE_CONCURRENCY) || 3;
    for (let i = 0; i < images.length; i += BATCH) {
      const batch = images.slice(i, i + BATCH);
      console.log(`Parsing batch ${Math.floor(i/BATCH)+1}: images ${i+1}–${Math.min(i+BATCH, images.length)}/${images.length}`);
      const batchResults = await Promise.all(batch.map(async (img) => {
        try {
          const b64 = img.buffer.toString('base64');
          const parsed = await parseReceiptWithLLM(b64, img.mimeType);
          const westernDate = convertRocToWestern(parsed.date);
          return {
            id: uuidv4(),
            filename: img.name,
            imageBase64: b64,
            date: westernDate || parsed.date,
            amount: parsed.amount,
            merchant: parsed.merchant,
            category: parsed.category || 'Others',
            description: parsed.description || parsed.merchant
          };
        } catch (e) {
          console.error(`Failed to parse ${img.name}: [${e.response?.status}] ${e.response?.data ? JSON.stringify(e.response.data) : e.message}`);
          return {
            id: uuidv4(),
            filename: img.name,
            imageBase64: img.buffer.toString('base64'),
            date: '',
            amount: 0,
            merchant: img.name,
            category: 'Others',
            description: 'Parse failed – please fill manually',
            error: e.message
          };
        }
      }));
      results.push(...batchResults);
      // Small pause between batches
      if (i + BATCH < images.length) await sleep(1000);
    }

    // Sort chronologically
    results.sort((a, b) => {
      if (!a.date) return 1;
      if (!b.date) return -1;
      return a.date.localeCompare(b.date);
    });

    console.log(`Done: ${results.length} receipts parsed`);
    res.json({ receipts: results, totalFiles: req.files.length, totalImages: images.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ─── POST /api/generate ───────────────────────────────────────────────────────
app.post('/api/generate', async (req, res) => {
  try {
    const { employeeName, receipts } = req.body;
    if (!receipts || receipts.length === 0) {
      return res.status(400).json({ error: 'No receipts provided' });
    }

    const excelBuffer = await generateExcel(employeeName || 'Samuel Chiang', receipts);

    const sorted = [...receipts].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    // Dates may still contain non-ASCII (e.g. unconverted ROC dates like 115年04月14日),
    // which is illegal in a Content-Disposition header — keep only YYYY-MM-DD-shaped dates.
    const safeDate = d => (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}/.test(d)) ? d.substring(0, 10) : 'unknown';
    const period = `${safeDate(sorted[0]?.date)}_to_${safeDate(sorted[sorted.length-1]?.date)}`;
    const filename = `Expense_Report_${period}.xlsx`.replace(/[^\w.\-]/g, '_');

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(excelBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Rambus Expense server running on port ${PORT}`));
