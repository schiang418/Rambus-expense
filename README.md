# Rambus Taiwan Expense Report Generator

A web app that parses receipt images (JPEG or ZIP) using AI vision, lets you review and edit the extracted data, and generates a fully populated Rambus Taiwan expense report Excel file for download.

## Features

- **Cumulative expense reports** — create a named report (e.g. "July 2026 Taiwan"), keep adding ZIPs or receipt photos to it over days or weeks, then trigger **Process All** when you're ready. Receipts are stored on a persistent volume, and parse results are cached per image so re-processing only sends *new* receipts to the AI.
- **Quick mode** — one-shot upload → parse → review → download, no storage.
- **iPhone-ready (PWA)** — open the app in Safari on your phone, tap Share → **Add to Home Screen**, and it installs like an app: snap receipt photos with the camera button, they upload straight into the open report, and you can generate the Excel from the phone too. Desktop and phone always see the same reports because everything lives on the server.
- AI-powered extraction of date, amount, merchant, and category from each receipt
- Automatic ROC (Republic of China) calendar date conversion (e.g. 115年 → 2026)
- Editable expense table — correct any field before export
- Generates the exact Rambus Taiwan Excel template (TWD tab + Receipts tab with embedded images)
- Auto-populates "For Period" date range and employee name

## Environment Variables

| Variable | Description |
|---|---|
| `FORGE_API_URL` | OpenAI-compatible API base URL (e.g. `https://api.openai.com`) |
| `FORGE_API_KEY` | API key for the LLM vision model |
| `FORGE_MODEL` | Vision model to use (default: `gpt-5.4`) |
| `PARSE_CONCURRENCY` | Receipts parsed in parallel (default: 3; lower it if you hit 429 rate limits) |
| `DATA_DIR` | Directory for stored reports — set to the Railway volume mount path, e.g. `/data` |
| `PORT` | Server port (Railway sets this automatically) |

## Railway Deployment

1. Go to [railway.app](https://railway.app) and sign in
2. Click **New Project → Deploy from GitHub repo**
3. Select `schiang418/Rambus-expense`
4. In **Variables**, add:
   - `FORGE_API_KEY` = your OpenAI API key
   - `FORGE_API_URL` = `https://api.openai.com` (or your proxy URL)
   - `DATA_DIR` = `/data`
5. **Attach a volume** (required for cumulative reports to survive redeploys):
   right-click the service → **Attach Volume** → set **Mount Path** to `/data`
6. Railway will auto-detect Node.js and deploy — done!

## Install on iPhone

1. Open the app URL in **Safari** on your iPhone
2. Tap the **Share** button → **Add to Home Screen**
3. Launch it from the home-screen icon — full screen, with a 📷 **Take Photo** button that opens the camera directly

## API

| Method & Path | Purpose |
|---|---|
| `POST /api/parse` | One-shot: upload ZIPs/images, parse all, return receipts |
| `POST /api/generate` | Build the Excel from a receipts payload (optional `reportSlug` stamps the report as generated) |
| `GET /api/reports` | List reports |
| `POST /api/reports` | Create a report `{name}` |
| `GET /api/reports/:slug` | Report detail + image list |
| `POST /api/reports/:slug/upload` | Add ZIPs/images to a report |
| `GET /api/reports/:slug/images/:id` | Fetch a stored receipt image |
| `DELETE /api/reports/:slug/images/:id` | Remove one receipt |
| `DELETE /api/reports/:slug` | Delete a report and its receipts |
| `POST /api/reports/:slug/parse` | The trigger: parse everything accumulated (cached per image), return receipts for review |

## Local Development

```bash
npm install
FORGE_API_KEY=sk-... node server.js
```

Then open http://localhost:3000
