// Normalize the many date formats found on receipts to ISO YYYY-MM-DD.
// Returns the input unchanged when no known pattern matches, so callers can
// decide whether to keep it as free text.
export function normalizeDate(dateStr) {
  if (!dateStr) return null;
  dateStr = String(dateStr).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(dateStr)) return dateStr.substring(0, 10);

  // Years > 1000 are already western (receipts sometimes print 2026/06/16) — only
  // small years are ROC and need +1911
  const yr = y => (+y > 1000 ? +y : +y + 1911);
  const pad = n => String(+n).padStart(2, '0');
  const iso = (y, mo, d) => {
    if (+mo < 1 || +mo > 12 || +d < 1 || +d > 31) return dateStr;
    return `${y}-${pad(mo)}-${pad(d)}`;
  };

  let m = dateStr.match(/(\d+)年(\d{1,2})-\d{1,2}月/);           // 115年03-04月 (billing range)
  if (m) return iso(yr(m[1]), m[2], 1);
  m = dateStr.match(/(\d+)年(\d{1,2})月(\d{1,2})[日號]?/);        // 115年04月28日
  if (m) return iso(yr(m[1]), m[2], m[3]);
  m = dateStr.match(/(\d+)年(\d{1,2})月/);                        // 115年04月 (day unknown)
  if (m) return iso(yr(m[1]), m[2], 1);

  const CN = {零:0,一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9,十:10,廿:20,卅:30};
  m = dateStr.match(/(\d+)年([一二三四五六七八九十廿卅]+)月([一二三四五六七八九十廿卅]+)[號日]/);
  if (m) {
    // 十/廿/卅 are bases: 廿八 = 20+8, 十五 = 10+5, 四 = 4
    const cn2int = s => [...s].reduce((v, c) => (CN[c] >= 10 ? v + CN[c] : v + (CN[c] || 0)), 0);
    return iso(yr(m[1]), cn2int(m[2]), cn2int(m[3]));
  }

  m = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);            // US style 07/26/2026
  if (m) return iso(m[3], m[1], m[2]);
  m = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);           // US style 07/26/26
  if (m) return iso(`20${m[3]}`, m[1], m[2]);
  m = dateStr.match(/(\d+)\/(\d{1,2})\/(\d{1,2})/);               // 115/4/28 or 2026/06/16
  if (m) return iso(yr(m[1]), m[2], m[3]);

  const MON = {jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
  m = dateStr.match(/(\d{1,2})\s*([A-Za-z]{3})[a-z]*\.?,?\s*(\d{4})/); // 08Jul2026, 8 Jul 2026
  if (m && MON[m[2].toLowerCase()]) return iso(m[3], MON[m[2].toLowerCase()], m[1]);
  m = dateStr.match(/([A-Za-z]{3})[a-z]*\.?\s+(\d{1,2}),?\s*(\d{4})/); // Jun 9, 2026
  if (m && MON[m[1].toLowerCase()]) return iso(m[3], MON[m[1].toLowerCase()], m[2]);
  m = dateStr.match(/^(\d{1,2})-(\d{1,2})-(\d{2})$/);               // 08-07-26 (US M-D-YY)
  if (m) return iso(`20${m[3]}`, m[1], m[2]);
  m = dateStr.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})$/);         // 2026.07.26
  if (m) return iso(m[1], m[2], m[3]);

  return dateStr;
}

export function isIsoDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s || '');
}
