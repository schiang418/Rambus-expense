# Rambus Taiwan Expense Report Generator

A web app that parses receipt images (JPEG or ZIP) using AI vision, lets you review and edit the extracted data, and generates a fully populated Rambus Taiwan expense report Excel file for download.

## Features

- Upload a ZIP file or JPEG receipt images via drag-and-drop
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
| `PORT` | Server port (Railway sets this automatically) |

## Railway Deployment

1. Go to [railway.app](https://railway.app) and sign in
2. Click **New Project → Deploy from GitHub repo**
3. Select `schiang418/Rambus-expense`
4. In **Variables**, add:
   - `FORGE_API_KEY` = your OpenAI API key
   - `FORGE_API_URL` = `https://api.openai.com` (or your proxy URL)
5. Railway will auto-detect Node.js and deploy — done!

## Local Development

```bash
npm install
FORGE_API_KEY=sk-... node server.js
```

Then open http://localhost:3000
