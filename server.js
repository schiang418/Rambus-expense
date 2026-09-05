import express from 'express';
import multer from 'multer';
import AdmZip from 'adm-zip';
import axios from 'axios';
import fs from 'fs';
import fsp from 'fs/promises';
import sharp from 'sharp';
import path from 'path';
import { fileURLToPath } from 'url';
import { v4 as uuidv4 } from 'uuid';
import { generateExcel } from './excelGenerator.js';
import { normalizeDate, isIsoDate } from './dateUtils.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 3000;

// Persistent storage for cumulative reports — point DATA_DIR at the Railway
// volume mount path (e.g. /data) so receipts survive restarts/redeploys
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const REPORTS_DIR = path.join(DATA_DIR, 'reports');
fs.mkdirSync(REPORTS_DIR, { recursive: true });

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

// ─── Report store helpers ─────────────────────────────────────────────────────
function slugify(name) {
  return name.trim().toLowerCase()
    .replace(/[^a-z0-9一-鿿]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 60) || `report-${Date.now()}`;
}

// Slugs are path components — reject anything that could traverse directories
function safeSlug(s) {
  return /^[a-z0-9一-鿿-]{1,60}$/.test(s || '') ? s : null;
}

function reportDir(slug) {
  return path.join(REPORTS_DIR, slug);
}

async function loadManifest(slug) {
  try {
    return JSON.parse(await fsp.readFile(path.join(reportDir(slug), 'manifest.json'), 'utf8'));
  } catch {
    return null;
  }
}

async function saveManifest(manifest) {
  manifest.updatedAt = new Date().toISOString();
  await fsp.writeFile(
    path.join(reportDir(manifest.slug), 'manifest.json'),
    JSON.stringify(manifest, null, 2)
  );
}

function reportSummary(m) {
  return {
    slug: m.slug,
    name: m.name,
    createdAt: m.createdAt,
    updatedAt: m.updatedAt,
    generatedAt: m.generatedAt || null,
    imageCount: m.images.length,
  };
}

// Phone photos carry an EXIF orientation tag instead of upright pixels.
// Bake the rotation in once so the stored file, the LLM input, and the Excel
// embed all agree on which way is up.
async function normalizeOrientation(img) {
  try {
    const meta = await sharp(img.buffer).metadata();
    if (meta.orientation && meta.orientation > 1) {
      img.buffer = await sharp(img.buffer).rotate().jpeg({ quality: 92 }).toBuffer();
      img.mimeType = 'image/jpeg';
    }
  } catch {
    // not a format sharp understands — leave it as-is
  }
  return img;
}

// Collect receipt images from uploaded files (ZIPs are expanded)
async function extractImages(files) {
  const images = [];
  for (const file of files) {
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
  return Promise.all(images.map(normalizeOrientation));
}

// Parse images through the LLM in small parallel batches.
// Returns an array aligned with the input: {date, amount, merchant, category, description, error?}
async function parseImagesBatch(images) {
  const out = [];
  const BATCH = Number(process.env.PARSE_CONCURRENCY) || 3;
  for (let i = 0; i < images.length; i += BATCH) {
    const batch = images.slice(i, i + BATCH);
    console.log(`Parsing batch ${Math.floor(i/BATCH)+1}: images ${i+1}–${Math.min(i+BATCH, images.length)}/${images.length}`);
    const batchResults = await Promise.all(batch.map(async (img) => {
      try {
        const parsed = await parseReceiptWithLLM(img.buffer.toString('base64'), img.mimeType);
        const westernDate = normalizeDate(parsed.date);
        return {
          date: westernDate || parsed.date || '',
          amount: parsed.amount,
          merchant: parsed.merchant,
          category: parsed.category || 'Others',
          description: parsed.description || parsed.merchant
        };
      } catch (e) {
        console.error(`Failed to parse ${img.name}: [${e.response?.status}] ${e.response?.data ? JSON.stringify(e.response.data) : e.message}`);
        return {
          date: '',
          amount: 0,
          merchant: img.name,
          category: 'Others',
          description: 'Parse failed – please fill manually',
          error: e.message
        };
      }
    }));
    out.push(...batchResults);
    if (i + BATCH < images.length) await sleep(1000);
  }
  return out;
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ─── LLM vision: parse one receipt image (with retry + exponential backoff) ───
async function parseReceiptWithLLM(imageBase64, mimeType = 'image/jpeg', retries = 3) {
  const apiUrl = (process.env.FORGE_API_URL || 'https://api.openai.com').replace(/\/$/, '');
  const apiKey = process.env.FORGE_API_KEY || process.env.OPENAI_API_KEY;
  const model = process.env.FORGE_MODEL || 'gpt-5.4';

  const prompt = `You are an expert at reading Taiwan receipts (both traditional Chinese and English).
Analyze this receipt image and extract the following information.
Return ONLY a valid JSON object with exactly these fields (no markdown, no explanation):
{
  "date": "YYYY-MM-DD",
  "amount": <number, total amount paid in TWD, no currency symbol>,
  "merchant": "merchant/business name as printed on the receipt (Chinese is fine)",
  "category": "one of: Airfare | Transportation | Lodging | Travel-Other | Meals | Entertainment | Telephone | Office Supplies | Others",
  "description": "concise English summary of what was purchased, 3-8 words"
}

Date rules:
- Always output ISO format YYYY-MM-DD.
- Taiwan receipts often use the ROC/Minguo calendar: add 1911 to the year (115年04月14日 or 115/4/14 → 2026-04-14).
- US receipts use month/day/year (07/26/2026 → 2026-07-26). Written months: 08Jul2026 or Jun 9, 2026 → 2026-07-08 / 2026-06-09.

Description rules:
- ALWAYS write the description in English, even when the receipt is entirely in Chinese.
- Summarize the purpose of the expense; do not transliterate or translate the receipt text verbatim.
- Good examples: "parking fee", "gasoline for company car", "mobile phone monthly bill", "lunch with client", "gym membership fee", "printing services", "HSR ticket Hsinchu to Taipei".

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

// ─── POST /api/parse (one-shot mode) ──────────────────────────────────────────
app.post('/api/parse', upload.array('files', 200), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No files uploaded' });
    }

    const images = await extractImages(req.files);
    if (images.length === 0) {
      return res.status(400).json({ error: 'No receipt images found in the uploaded files' });
    }

    console.log(`Processing ${images.length} receipt image(s) from ${req.files.length} uploaded file(s)`);
    const parsed = await parseImagesBatch(images);

    const results = images.map((img, i) => ({
      id: uuidv4(),
      filename: img.name,
      imageBase64: img.buffer.toString('base64'),
      ...parsed[i]
    }));

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

// ─── Cumulative reports (persisted on the Railway volume) ─────────────────────

// List all reports
app.get('/api/reports', async (req, res) => {
  try {
    const entries = await fsp.readdir(REPORTS_DIR, { withFileTypes: true });
    const reports = [];
    for (const e of entries) {
      if (!e.isDirectory()) continue;
      const m = await loadManifest(e.name);
      if (m) reports.push(reportSummary(m));
    }
    reports.sort((a, b) => (b.updatedAt || b.createdAt).localeCompare(a.updatedAt || a.createdAt));
    res.json({ reports });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Create a report (or return the existing one with the same name)
app.post('/api/reports', async (req, res) => {
  try {
    const name = (req.body?.name || '').trim();
    if (!name) return res.status(400).json({ error: 'Report name is required' });
    const slug = slugify(name);
    let m = await loadManifest(slug);
    if (!m) {
      await fsp.mkdir(reportDir(slug), { recursive: true });
      m = {
        slug,
        name,
        createdAt: new Date().toISOString(),
        generatedAt: null,
        images: []
      };
      await saveManifest(m);
    }
    res.json(reportSummary(m));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Report detail (image list without pixel data)
app.get('/api/reports/:slug', async (req, res) => {
  const slug = safeSlug(req.params.slug);
  const m = slug && await loadManifest(slug);
  if (!m) return res.status(404).json({ error: 'Report not found' });
  res.json({
    ...reportSummary(m),
    images: m.images.map(i => ({
      id: i.id,
      originalName: i.originalName,
      mimeType: i.mimeType,
      size: i.size,
      addedAt: i.addedAt,
      parsed: i.parsed ? { ...i.parsed } : null
    }))
  });
});

// Add ZIPs / photos to a report (accumulates on the volume)
app.post('/api/reports/:slug/upload', upload.array('files', 200), async (req, res) => {
  try {
    const slug = safeSlug(req.params.slug);
    const m = slug && await loadManifest(slug);
    if (!m) return res.status(404).json({ error: 'Report not found' });

    const images = await extractImages(req.files || []);
    if (images.length === 0) {
      return res.status(400).json({ error: 'No receipt images found in the uploaded files' });
    }

    for (const img of images) {
      const id = uuidv4();
      // Stored extension must match the (possibly re-encoded) bytes, not the upload name
      const ext = Object.entries(MIME_BY_EXT).find(([, m]) => m === img.mimeType)?.[0] || '.jpg';
      const storedName = `${id}${ext}`;
      await fsp.writeFile(path.join(reportDir(slug), storedName), img.buffer);
      m.images.push({
        id,
        storedName,
        originalName: img.name,
        mimeType: img.mimeType,
        size: img.buffer.length,
        addedAt: new Date().toISOString(),
        parsed: null
      });
    }
    await saveManifest(m);
    console.log(`Report "${m.name}": added ${images.length} image(s), now ${m.images.length} total`);
    res.json({ added: images.length, imageCount: m.images.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Serve a stored receipt image (thumbnails / preview)
app.get('/api/reports/:slug/images/:id', async (req, res) => {
  const slug = safeSlug(req.params.slug);
  const m = slug && await loadManifest(slug);
  const img = m?.images.find(i => i.id === req.params.id);
  if (!img) return res.status(404).json({ error: 'Image not found' });
  try {
    const buf = await fsp.readFile(path.join(reportDir(slug), img.storedName));
    res.setHeader('Content-Type', img.mimeType || 'image/jpeg');
    res.setHeader('Cache-Control', 'private, max-age=3600');
    res.send(buf);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Remove one receipt from a report
app.delete('/api/reports/:slug/images/:id', async (req, res) => {
  try {
    const slug = safeSlug(req.params.slug);
    const m = slug && await loadManifest(slug);
    if (!m) return res.status(404).json({ error: 'Report not found' });
    const idx = m.images.findIndex(i => i.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Image not found' });
    const [img] = m.images.splice(idx, 1);
    await fsp.rm(path.join(reportDir(slug), img.storedName), { force: true });
    await saveManifest(m);
    res.json({ imageCount: m.images.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Delete a whole report
app.delete('/api/reports/:slug', async (req, res) => {
  try {
    const slug = safeSlug(req.params.slug);
    if (!slug || !(await loadManifest(slug))) return res.status(404).json({ error: 'Report not found' });
    await fsp.rm(reportDir(slug), { recursive: true, force: true });
    res.json({ deleted: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// The trigger: parse everything accumulated in a report.
// Results are cached per image in the manifest, so only new (or previously
// failed) receipts hit the LLM on subsequent runs.
app.post('/api/reports/:slug/parse', async (req, res) => {
  try {
    const slug = safeSlug(req.params.slug);
    const m = slug && await loadManifest(slug);
    if (!m) return res.status(404).json({ error: 'Report not found' });
    if (m.images.length === 0) return res.status(400).json({ error: 'This report has no receipts yet' });

    // force=true re-parses everything (e.g. after a prompt change); otherwise
    // only new or previously-failed receipts hit the LLM
    const force = req.body?.force === true;
    const toParse = m.images.filter(i => force || !i.parsed || i.parsed.error);
    if (toParse.length > 0) {
      console.log(`Report "${m.name}": parsing ${toParse.length} of ${m.images.length} receipt(s)`);
      const buffers = await Promise.all(toParse.map(async i => ({
        name: i.originalName,
        buffer: await fsp.readFile(path.join(reportDir(slug), i.storedName)),
        mimeType: i.mimeType
      })));
      const parsed = await parseImagesBatch(buffers);
      toParse.forEach((i, idx) => { i.parsed = parsed[idx]; });
      await saveManifest(m);
    }

    const receipts = await Promise.all(m.images.map(async i => ({
      id: i.id,
      filename: i.originalName,
      imageBase64: (await fsp.readFile(path.join(reportDir(slug), i.storedName))).toString('base64'),
      date: i.parsed?.date || '',
      amount: i.parsed?.amount ?? 0,
      merchant: i.parsed?.merchant || i.originalName,
      category: i.parsed?.category || 'Others',
      description: i.parsed?.description || '',
      ...(i.parsed?.error ? { error: i.parsed.error } : {})
    })));

    receipts.sort((a, b) => {
      if (!a.date) return 1;
      if (!b.date) return -1;
      return a.date.localeCompare(b.date);
    });

    res.json({
      receipts,
      totalFiles: m.images.length,
      totalImages: m.images.length,
      parsedNow: toParse.length,
      fromCache: m.images.length - toParse.length
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ─── POST /api/generate ───────────────────────────────────────────────────────
app.post('/api/generate', async (req, res) => {
  try {
    const { employeeName, receipts, reportSlug } = req.body;
    if (!receipts || receipts.length === 0) {
      return res.status(400).json({ error: 'No receipts provided' });
    }

    const excelBuffer = await generateExcel(employeeName || 'Samuel Chiang', receipts);

    // When generating from a cumulative report, stamp it as generated
    if (reportSlug) {
      const slug = safeSlug(reportSlug);
      const m = slug && await loadManifest(slug);
      if (m) {
        m.generatedAt = new Date().toISOString();
        await saveManifest(m);
      }
    }

    const sorted = [...receipts].sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    // Dates may still contain non-ASCII (e.g. unconverted ROC dates like 115年04月14日),
    // which is illegal in a Content-Disposition header — keep only YYYY-MM-DD-shaped dates.
    const safeDate = d => { const n = normalizeDate(d); return isIsoDate(n) ? n : 'unknown'; };
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
