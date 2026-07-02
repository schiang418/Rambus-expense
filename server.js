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

function isImage(filename) {
  return IMAGE_EXTS.has(path.extname(filename).toLowerCase());
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ─── ROC date converter ───────────────────────────────────────────────────────
function convertRocToWestern(dateStr) {
  if (!dateStr) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) return dateStr.substring(0, 10);
  let m = dateStr.match(/(\d+)年(\d{1,2})-\d{1,2}月/);
  if (m) return `${+m[1]+1911}-${String(+m[2]).padStart(2,'0')}-01`;
  const CN = {零:0,一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10,廿:20,卅:30};
  m = dateStr.match(/(\d+)年([一二三四五六七八九十廿卅]+)月([一二三四五六七八九十廿卅]+)[號日]/);
  if (m) {
    const cn2int = s => [...s].reduce((v,c) => v*10 + (CN[c]||0), 0);
    return `${+m[1]+1911}-${String(cn2int(m[2])).padStart(2,'0')}-${String(cn2int(m[3])).padStart(2,'0')}`;
  }
  m = dateStr.match(/(\d+)\/(\d{1,2})\/(\d{1,2})/);
  if (m) return `${+m[1]+1911}-${String(+m[2]).padStart(2,'0')}-${String(+m[3]).padStart(2,'0')}`;
  return dateStr;
}

// ─── LLM vision: parse one receipt image (with retry + exponential backoff) ───
async function parseReceiptWithLLM(imageBase64, mimeType = 'image/jpeg', retries = 4) {
  const apiUrl = (process.env.FORGE_API_URL || 'https://api.openai.com').replace(/\/$/, '');
  const apiKey = process.env.FORGE_API_KEY || process.env.OPENAI_API_KEY;

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
          model: 'gpt-4o',
          messages: [{
            role: 'user',
            content: [
              { type: 'image_url', image_url: { url: `data:${mimeType};base64,${imageBase64}`, detail: 'high' } },
              { type: 'text', text: prompt }
            ]
          }],
          max_tokens: 300,
          temperature: 0
        },
        {
          headers: { Authorization: `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
          timeout: 30000
        }
      );
      const content = response.data.choices[0].message.content.trim();
      const cleaned = content.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();
      return JSON.parse(cleaned);
    } catch (err) {
      const status = err.response?.status;
      // Retry on rate limit (429), server error (500/503), or timeout
      if ((status === 429 || status === 500 || status === 503 || err.code === 'ECONNABORTED') && attempt < retries) {
        const wait = attempt * 10000; // 10s, 20s, 30s
        console.warn(`Attempt ${attempt}/${retries} failed (${status || err.code}) for image, retrying in ${wait/1000}s...`);
        await sleep(wait);
      } else {
        throw err;
      }
    }
  }
}

// ─── POST /api/parse ──────────────────────────────────────────────────────────
// Accepts: multipart with field "files" (multiple ZIPs and/or images)
// Returns: JSON array of parsed receipt entries
app.post('/api/parse', upload.array('files', 200), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    // Collect all images from every uploaded file
    const images = [];  // { name, buffer, mimeType }

    for (const file of req.files) {
      const ext = path.extname(file.originalname).toLowerCase();

      if (ext === '.zip') {
        // Extract all images from the ZIP
        const zip = new AdmZip(file.buffer);
        for (const entry of zip.getEntries()) {
          if (!entry.isDirectory && isImage(entry.entryName)) {
            // Skip macOS metadata files
            const baseName = path.basename(entry.entryName);
            if (baseName.startsWith('._') || baseName.startsWith('.')) continue;
            images.push({
              name: entry.entryName,
              buffer: entry.getData(),
              mimeType: 'image/jpeg'
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
      // Silently skip unsupported file types
    }

    if (images.length === 0) {
      return res.status(400).json({ error: 'No receipt images found in the uploaded files' });
    }

    console.log(`Processing ${images.length} receipt image(s) from ${req.files.length} uploaded file(s)`);

    // Parse each image sequentially with a delay between calls to avoid rate limits
    const results = [];
    for (let i = 0; i < images.length; i++) {
      const img = images[i];
      console.log(`Parsing ${i + 1}/${images.length}: ${img.name}`);
      try {
        const b64 = img.buffer.toString('base64');
        const parsed = await parseReceiptWithLLM(b64, img.mimeType);
        const westernDate = convertRocToWestern(parsed.date);
        results.push({
          id: uuidv4(),
          filename: img.name,
          imageBase64: b64,
          date: westernDate || parsed.date,
          amount: parsed.amount,
          merchant: parsed.merchant,
          category: parsed.category || 'Others',
          description: parsed.description || parsed.merchant
        });
      } catch (e) {
        console.error(`Failed to parse ${img.name}:`, e.message);
        results.push({
          id: uuidv4(),
          filename: img.name,
          imageBase64: img.buffer.toString('base64'),
          date: '',
          amount: 0,
          merchant: img.name,
          category: 'Others',
          description: 'Parse failed – please fill manually',
          error: e.message
        });
      }
      // 2s delay between requests to stay within rate limits
      if (i < images.length - 1) await sleep(2000);
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
    const period = `${sorted[0]?.date || 'unknown'}_to_${sorted[sorted.length-1]?.date || 'unknown'}`;
    const filename = `Expense_Report_${period}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(excelBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => console.log(`Rambus Expense server running on port ${PORT}`));
