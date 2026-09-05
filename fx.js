// Foreign-currency support: look up the TWD exchange rate for a receipt's date,
// add a markup for credit-card conversion fees, and convert amounts to TWD.
//
// Rate sources, in order:
//   1. Bank of Taiwan historical closing rates (spot selling). Their site sits
//      behind a JavaScript bot-challenge that blocks most server requests, so
//      this usually fails — it's tried first in case it works from the host.
//   2. @fawazahmed0/currency-api — free, keyless daily mid-market rates that
//      include TWD and CNY, with history by date.
// Results are cached on disk so a date+currency is only fetched once.
import axios from 'axios';
import fs from 'fs';
import fsp from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const CACHE_FILE = path.join(DATA_DIR, 'fx-cache.json');

// Card issuers typically add ~1.5% on top of the bank rate
export const FX_MARKUP = Number(process.env.FX_MARKUP ?? 0.015);

// Receipt currency → {Bank of Taiwan code, currency-api code}
const CODES = {
  USD: { bot: 'USD', api: 'usd' },
  RMB: { bot: 'CNY', api: 'cny' },
};

export function normCurrency(c) {
  const raw = String(c || 'TWD');
  if (/人民|RMB|CNY|RENMINBI/i.test(raw)) return 'RMB';
  if (/台|臺|NT\$|TWD|NTD/i.test(raw)) return 'TWD';
  const s = raw.toUpperCase().replace(/[^A-Z$¥]/g, '');
  if (s.startsWith('US') || s === '$') return 'USD';
  if (s === '¥' || s.startsWith('CN')) return 'RMB';
  return 'TWD';
}

const round2 = n => Math.round(n * 100) / 100;

let cache = null;
async function loadCache() {
  if (cache) return cache;
  try {
    cache = JSON.parse(await fsp.readFile(CACHE_FILE, 'utf8'));
  } catch {
    cache = {};
  }
  return cache;
}

async function saveCache() {
  try {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    await fsp.writeFile(CACHE_FILE, JSON.stringify(cache, null, 2));
  } catch (e) {
    console.warn('fx cache not saved:', e.message);
  }
}

function shiftDate(iso, days) {
  const d = new Date(`${iso}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().substring(0, 10);
}

// Bank of Taiwan CSV: rows look like  "USD,本行買入,31.895,32.245,..." and
// "USD,本行賣出,32.565,32.345,..." (cash then spot). We want spot selling.
async function fetchBankOfTaiwan(currency, date) {
  const code = CODES[currency].bot;
  const res = await axios.get(`https://rate.bot.com.tw/xrt/flcsv/0/${date}`, {
    timeout: 8000,
    responseType: 'text',
    headers: { 'User-Agent': 'Mozilla/5.0', Accept: 'text/csv,*/*' },
  });
  const body = String(res.data);
  if (/Challenge Validation|<html/i.test(body)) throw new Error('Bank of Taiwan blocked the request');
  const line = body.split(/\r?\n/).find(l => l.startsWith(`${code},`) && l.includes('賣出'));
  if (!line) throw new Error(`no ${code} row`);
  const nums = line.split(',').slice(2).map(Number).filter(n => Number.isFinite(n) && n > 0);
  if (nums.length < 2) throw new Error(`could not parse ${code} row`);
  return { rate: nums[1], source: 'Bank of Taiwan spot selling' };
}

async function fetchCurrencyApi(currency, date) {
  const base = CODES[currency].api;
  const urls = [
    `https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@${date}/v1/currencies/${base}.json`,
    `https://${date}.currency-api.pages.dev/v1/currencies/${base}.json`,
  ];
  let lastErr;
  for (const url of urls) {
    try {
      const res = await axios.get(url, { timeout: 8000 });
      const rate = res.data?.[base]?.twd;
      if (rate > 0) return { rate, source: 'currency-api mid-market' };
      lastErr = new Error('no twd rate in response');
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr;
}

// Base (un-marked-up) TWD rate for one unit of `currency` on `date`,
// stepping back to earlier days when the date has no data yet.
export async function getBaseRate(currency, date) {
  currency = normCurrency(currency);
  if (currency === 'TWD') return { rate: 1, source: 'n/a', date };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) throw new Error('a valid receipt date is needed to look up the exchange rate');

  await loadCache();
  const key = `${date}:${currency}`;
  if (cache[key]) return cache[key];

  let d = date;
  let lastErr;
  for (let back = 0; back < 7; back++) {
    for (const fetcher of [fetchBankOfTaiwan, fetchCurrencyApi]) {
      try {
        const { rate, source } = await fetcher(currency, d);
        const entry = { rate, source, date: d, fetchedAt: new Date().toISOString() };
        cache[key] = entry;
        await saveCache();
        return entry;
      } catch (e) {
        lastErr = e;
      }
    }
    d = shiftDate(d, -1);
  }
  throw new Error(`no ${currency} rate found for ${date} (${lastErr?.message || 'unknown error'})`);
}

// Marked-up rate rounded to 2 decimals, as a person would write it in the report
export async function getRate(currency, date) {
  const base = await getBaseRate(currency, date);
  return { ...base, baseRate: base.rate, rate: round2(base.rate * (1 + FX_MARKUP)), markup: FX_MARKUP };
}

// Fill amountTwd / fxRate on every receipt. A receipt with a manual
// `fxOverride` uses that rate as-is (no extra markup).
export async function applyFx(receipts) {
  for (const r of receipts) {
    r.currency = normCurrency(r.currency);
    const amount = Number(r.amount) || 0;
    delete r.fxError;

    if (r.currency === 'TWD') {
      r.fxRate = null;
      r.fxSource = null;
      r.amountTwd = amount;
      continue;
    }

    if (Number(r.fxOverride) > 0) {
      r.fxRate = Number(r.fxOverride);
      r.fxSource = 'manual';
    } else {
      try {
        const info = await getRate(r.currency, r.date);
        r.fxRate = info.rate;
        r.fxSource = `${info.source}${info.date !== r.date ? ` (${info.date})` : ''} +${(FX_MARKUP * 100).toFixed(1)}%`;
      } catch (e) {
        r.fxRate = null;
        r.fxSource = null;
        r.fxError = e.message;
      }
    }
    r.amountTwd = r.fxRate ? Math.round(amount * r.fxRate) : 0;
  }
  return receipts;
}
