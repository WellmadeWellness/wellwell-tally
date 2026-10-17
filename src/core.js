/* Wellwell Monthly Tally – parsing and calculation core.
   Runs in the browser (inlined) and in Node (for tests). Needs the SheetJS XLSX object. */
(function (root) {
  'use strict';

  // ---------- helpers ----------
  const S = v => (v === null || v === undefined) ? '' : String(v).replace(/\t/g, ' ').replace(/\s+/g, ' ').trim();
  const isNum = v => typeof v === 'number' && isFinite(v);
  const num = v => isNum(v) ? v : (typeof v === 'string' && /^\s*-?\d+(\.\d+)?\s*$/.test(v) ? Number(v) : null);
  const pad = n => String(n).padStart(2, '0');
  function iso(y, m, d) { return `${y}-${pad(m)}-${pad(d)}`; }
  function toISO(XLSX, v) {
    if (isNum(v) && v > 20000 && v < 80000) { const p = XLSX.SSF.parse_date_code(v); return iso(p.y, p.m, p.d); }
    if (isNum(v) && v > 19000000 && v < 21000000) { const s = String(v); return iso(s.slice(0, 4), s.slice(4, 6), s.slice(6, 8)); }
    const s = S(v); let m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (m) return iso(m[1], m[2], m[3]);
    return null;
  }
  function hash(str) { let h = 2166136261; for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(36); }
  function normName(p) {
    let s = S(p).replace(/^\[(.*?)\]\s*/, '$1 ');
    return s.replace(/\s+/g, ' ').trim();
  }
  function rowsOf(XLSX, ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false }); }
  function findHeader(rows, need, avoid, maxScan) {
    for (let i = 0; i < Math.min(rows.length, maxScan || 15); i++) {
      const r = rows[i].map(S);
      if (need.every(n => r.includes(n)) && !(avoid || []).some(a => r.includes(a))) {
        const map = {}; r.forEach((v, j) => { if (v && !(v in map)) map[v] = j; });
        return { idx: i, map };
      }
    }
    return null;
  }
  const firstCell = rows => { for (const r of rows) for (const c of r) if (c !== null && S(c) !== '') return c; return null; };

  // ---------- channel rules for parcel order numbers (learned from the July 위탁 sheet) ----------
  const ORDER_RULES = [
    [/^3[34]\d{8}$/, '카카오톡 스토어'], [/^20\d{14}$/, '스마트스토어-B'], [/^20\d{8}$/, '이지웰몰'], [/^20\d{15}$/, '11번가'],
    [/^4[45]\d{8}$/, '지마켓'], [/^25\d{8}$/, '옥션'], [/^20\d{12}-\d{3}/, '홈앤쇼핑'], [/^20\d{24}$/, '홈앤쇼핑'],
    [/^(56|26)\d{10}$/, '농수산쇼핑몰'], [/^(2[89]|3[34])\d{7}$/, '오늘의집'], [/^PO/, 'TEMU'], [/^\d{13,14}$/, '쿠팡'], [/^20\d{16}$/, '샵바이(파트너사)']
  ];
  const NON_SALE = ['cs', '샘플', '인플샘플', '인플 샘플', '직원구매', '택배사고', 'cs디스이즈부담', '우리부담cs', '구매인증', '체험단'];
  const B2B_NAMES = ['녹선', '국민상회(썬스토리)', '삼성웰스토리', '웰스토리', '삼영군납', '슈퍼키친'];
  function defaultChannelType(ch) {
    if (/^벌크/.test(ch) || /^오프라인/.test(ch) || B2B_NAMES.includes(ch)) return 'B2B';
    if (NON_SALE.some(n => ch === n || ch.includes('샘플') || ch.endsWith('체험단'))) return 'Non-sale';
    return 'Online';
  }

  // ---------- format parsers ----------
  function parseBmart(XLSX, rows, ctx) {
    const h = findHeader(rows, ['업체', '발주번호', '상품명', '발주수량']); if (!h) return null;
    const date = toISO(XLSX, firstCell(rows.slice(0, h.idx)));
    const out = [];
    for (let i = h.idx + 1; i < rows.length; i++) {
      const r = rows[i], g = k => r[h.map[k]];
      if (S(r[0]) === '합계') break;
      const prod = S(g('상품명')); if (!prod || !isNum(g('발주수량'))) continue;
      out.push({ cat: 'Retail', channel: 'B마트', date, product: prod, qty: g('발주수량'), boxes: g('박스수량'),
        ref: S(g('발주번호')), center: S(g('입고')), supplier: S(g('업체')),
        key: ['B마트', S(g('발주번호')), S(g('입고')), prod, g('발주수량')].join('|') });
    }
    return { kind: 'Retail · B마트 order', date, lines: out };
  }
  function parseCoupang(XLSX, rows, ctx) {
    const h = findHeader(rows, ['물류센터', '발주번호', '상품명', '발주수량']); if (!h) return null;
    const date = toISO(XLSX, firstCell(rows.slice(0, h.idx)));
    const out = [];
    for (let i = h.idx + 1; i < rows.length; i++) {
      const r = rows[i], g = k => r[h.map[k]];
      const prod = S(g('상품명')); if (!prod || !isNum(g('발주수량'))) continue;
      const sup = S(g('구분'));
      out.push({ cat: 'Retail', channel: '쿠팡', date, product: prod, qty: g('발주수량'), boxes: g('박스수량'),
        ref: S(g('발주번호')), center: S(g('물류센터')), supplier: sup, thirdParty: !!sup && sup !== '웰웰',
        key: ['쿠팡', S(g('발주번호')), S(g('물류센터')), prod, g('발주수량')].join('|') });
    }
    return { kind: 'Retail · 쿠팡 order', date, lines: out };
  }
  function parseKurly(XLSX, rows, ctx) {
    if (!findHeader(rows, ['입고예정일', '상품명', '낱개'], null, 40)) return null;
    const out = []; let map = null, storage = '냉장';
    for (const r of rows) {
      const t = r.map(S); const first = t.find(x => x);
      if (first && first.startsWith('*') && first.includes('컬리')) { storage = first.includes('냉동') ? '냉동' : '냉장'; map = null; continue; }
      if (t.includes('상품명') && t.includes('낱개')) { map = {}; t.forEach((v, j) => { if (v && !(v in map)) map[v] = j; }); continue; }
      if (!map) continue;
      const g = k => r[map[k]]; const prod = S(g('상품명'));
      if (!prod || !isNum(g('낱개'))) continue;
      const date = toISO(XLSX, g('입고예정일'));
      out.push({ cat: 'Retail', channel: '마켓컬리', date, product: prod, qty: g('낱개'), boxes: g('박스'),
        ref: S(g('상품코드')), center: S(g('입고지')), supplier: '경한글로벌', storage,
        key: ['마켓컬리', date, S(g('상품코드')), S(g('입고지')), g('낱개')].join('|') });
    }
    const dates = out.map(l => l.date).filter(Boolean).sort();
    return { kind: 'Retail · 마켓컬리 order', date: dates[0] || null, lines: out };
  }
  function shipDateFromName(name, fallbackYear) {
    const base = name.replace(/\.[^.]+$/, '');
    let m = base.match(/(?:^|\D)(\d{1,2})\.(\d{1,2})(?:\D|$)/);
    if (m) return iso(fallbackYear, m[1], m[2]);
    const all = [...base.matchAll(/(?:^|\D)(\d{4})(?=\D|$)/g)].map(x => x[1]).filter(x => +x.slice(0, 2) >= 1 && +x.slice(0, 2) <= 12 && +x.slice(2) >= 1 && +x.slice(2) <= 31);
    if (all.length) { const x = all[all.length - 1]; return iso(fallbackYear, x.slice(0, 2), x.slice(2)); }
    return null;
  }
  function parseParcel(XLSX, rows, ctx) {
    const h = findHeader(rows, ['상품명', '수량', '주문번호'], ['택배사', '운송장번호', '송장번호', '결제일'], 3); if (!h) return null;
    const dcol = '주문일자' in h.map ? '주문일자' : ('날짜' in h.map ? '날짜' : null);
    const out = [];
    for (let i = h.idx + 1; i < rows.length; i++) {
      const r = rows[i], g = k => (k in h.map) ? r[h.map[k]] : null;
      const prod = S(g('상품명')); const q = num(g('수량'));
      if (!prod || q === null) continue;
      const ord = S(g('주문번호')), mk = S(g('업체명')), chCol = S(g('채널'));
      let channel = null, how = 'named';
      if (ord && !/^\d/.test(ord)) channel = ord;
      else if (chCol) channel = chCol;
      else {
        const id = mk.split(' ')[0];
        for (const [re, c] of ORDER_RULES) if (re.test(id)) { channel = c; how = 'order-number format'; break; }
        if (!channel && id && !/^\d/.test(id)) channel = id;
      }
      if (!channel) { channel = 'Unidentified'; how = 'unknown'; }
      const od = dcol ? toISO(XLSX, g(dcol)) : null;
      const who = hash(S(g('수취인명')) + '|' + S(g('우편번호')));   // one-way; names and addresses are never kept
      out.push({ cat: 'Parcel', channel, how, orderDate: od, product: prod, qty: q, ref: ord, market: mk,
        key: ['P', ord, mk, prod, q, who].join('|') });
    }
    const ods = out.map(l => l.orderDate).filter(d => d && d > '2015').sort();
    const year = ods.length ? ods[ods.length - 1].slice(0, 4) : String(new Date().getFullYear());
    let ship = shipDateFromName(ctx.fileName, year) || (ods.length ? ods[ods.length - 1] : null);
    out.forEach(l => { l.date = ship; });
    return { kind: 'Parcel · 택배 발주', date: ship, lines: out, shipFromName: !!shipDateFromName(ctx.fileName, year) };
  }
  const RETAIL_WORDS = /(쿠팡|b마트|비마트|컬리)/i;
  function parseOther(XLSX, rows, ctx) {
    const h = findHeader(rows, ['상품명', '발주수량', '박스수량'], ['업체', '물류센터'], 6); if (!h) return null;
    const title = S(firstCell(rows.slice(0, h.idx)));
    const m = title.match(/^(\d{1,2})\s*\/\s*(\d{1,2})\s*(?:\([^)]*\))?\s*(.*?)\s*(발주|출고|입고)/);
    if (!m) return null;
    const cust = S(m[3]).replace(/\s*\(.*\)$/, '');
    if (RETAIL_WORDS.test(cust)) return { kind: 'skip', reason: `Old-format ${cust} list – Retail comes from the order files` };
    const out = [];
    let expYear = null;
    for (let i = h.idx + 1; i < rows.length; i++) {
      const r = rows[i], g = k => (k in h.map) ? r[h.map[k]] : null;
      if (S(r[0]) === '합계' || S(r.find(x => x !== null)) === '합계') break;
      const prod = S(g('상품명')); const q = g('발주수량');
      if (!prod || !isNum(q) || /^#/.test(prod)) continue;
      const exp = toISO(XLSX, g('유통기한')) || toISO(XLSX, g('유통기한가이드'));
      if (exp && !expYear) expYear = exp;
      out.push({ cat: 'Other', channel: cust, product: prod, qty: q, boxes: g('박스수량'), expiry: exp, center: S(g('입고처')) });
    }
    // year: the expiry date is the only year in these sheets – an order is shipped before it expires
    let y = new Date().getFullYear();
    if (expYear) { const ey = +expYear.slice(0, 4), md = expYear.slice(5); y = (md >= iso(2000, m[1], m[2]).slice(5)) ? ey : ey - 1; }
    const date = iso(y, m[1], m[2]);
    out.forEach(l => { l.date = date; l.key = ['O', cust, date, l.product, l.qty].join('|'); });
    return { kind: `Other · ${cust}`, date, lines: out };
  }

  // ---------- workbook ----------
  function parseWorkbook(XLSX, data, fileName, today) {
    const wb = XLSX.read(data, { type: 'array', cellDates: false });
    const res = { fileName, sheets: [], lines: [], notes: [], kinds: new Set() };
    let parcelTaken = false; const oldSheets = []; const skipped = [];
    const now = today || new Date().toISOString().slice(0, 10);
    if (wb.SheetNames.includes('총 합계') && wb.SheetNames.includes('위탁')) {
      res.settlement = parseSettlement(XLSX, wb, fileName);
      const st = res.settlement;
      res.notes.push(st ? `Month-end settlement for ${st.month} (${st.site}), total ${Math.round(st.total).toLocaleString('ko-KR')}. Not added to the tally – used in the Compare tab.` : 'Looks like a settlement file, but its 총 합계 sheet could not be read.');
      res.kinds = ['Settlement']; return res;
    }
    for (const name of wb.SheetNames) {
      const rows = rowsOf(XLSX, wb.Sheets[name]); if (!rows.length) continue;
      let r = parseBmart(XLSX, rows) || parseCoupang(XLSX, rows) || parseKurly(XLSX, rows);
      if (!r && !parcelTaken) { r = parseParcel(XLSX, rows, { fileName }); if (r) parcelTaken = true; }
      if (!r) r = parseOther(XLSX, rows, { fileName });
      if (!r) { skipped.push(name); continue; }
      if (r.kind === 'skip') { oldSheets.push(name); continue; }
      if (r.date && (daysBetween(r.date, now) > 365 || daysBetween(now, r.date) > 45)) { oldSheets.push(name); continue; }
      r.lines.forEach(l => { l.sheet = name; l.file = fileName; });
      res.sheets.push({ name, kind: r.kind, lines: r.lines.length, units: r.lines.reduce((a, l) => a + l.qty, 0), date: r.date, shipFromName: r.shipFromName });
      res.kinds.add(r.kind.split(' · ')[0]);
      res.lines.push(...r.lines);
    }
    if (oldSheets.length) res.notes.push(`${oldSheets.length} old sheet${oldSheets.length > 1 ? 's' : ''} skipped (dated outside the last year, or an old outbound list): ${oldSheets.slice(0, 4).join(', ')}${oldSheets.length > 4 ? ' …' : ''}`);
    if (skipped.length) res.minor = `${skipped.length} other sheet${skipped.length > 1 ? 's' : ''} not used (pivots, print or template sheets)`;
    if (!res.lines.length) res.notes.push('No order lines recognised in this file.');
    res.kinds = [...res.kinds];
    return res;
  }
  function daysBetween(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 864e5); }

  // ---------- pricing and totals ----------
  function priceFor(product, prices) {
    const n = normName(product);
    if (prices[n]) return { price: prices[n][0], src: prices[n][1], matched: n };
    const gift = /사은품/.test(n);
    const c2 = n.replace(/\(사은품\)/g, '').replace(/\*[^*]*\*/g, '').replace(/\s+/g, ' ').trim();
    if (prices[c2]) return { price: prices[c2][0], src: (gift ? 'Free gift at cost of ' : 'Same item as ') + c2, matched: c2, gift };
    const c3 = c2.replace(/\((?:[^()]*뚜껑[^()]*)\)/g, '').replace(/\s+/g, ' ').trim();
    if (prices[c3]) return { price: prices[c3][0], src: 'Same item as ' + c3, matched: c3, gift };
    return { price: null, src: null, matched: c3, gift };
  }
  function compute(files, prices, channelTypes, opts) {
    opts = opts || {};
    const seen = new Map(); const all = [];
    for (const f of files) for (const l of f.lines) {
      const x = Object.assign({}, l);
      if (f.shipOverride && x.cat === 'Parcel') x.date = f.shipOverride;
      x.month = x.date ? x.date.slice(0, 7) : 'unknown';
      const p = priceFor(x.product, prices);
      x.price = p.price; x.priceSrc = p.src; x.gift = !!p.gift; x.priceKey = p.matched;
      x.value = p.price ? p.price * x.qty : 0;
      x.type = x.cat === 'Parcel' ? (channelTypes[x.channel] || defaultChannelType(x.channel)) : (x.cat === 'Retail' ? 'Retail' : 'B2B');
      x.group = x.cat;   // every line stays in the group of the file it came from (B2B lines in 택배 files stay in 택배)
      x.excluded = null; x.fileId = f.id;
      if (x.thirdParty && !opts.includeThirdParty) x.excluded = 'Third-party supplier (' + x.supplier + ')';
      const prev = seen.get(x.key);
      if (prev && prev !== f.fileName) x.excluded = 'Duplicate of a line in ' + prev;
      else if (!prev) seen.set(x.key, f.fileName);
      all.push(x);
    }
    return all;
  }
  // ---------- month-end settlement file (Phase 3) ----------
  function parseSettlement(XLSX, wb, fileName) {
    const tot = rowsOf(XLSX, wb.Sheets['총 합계']);
    const title = S(firstCell(tot));
    const mm = title.match(/(\d{1,2})\s*월/); if (!mm) return null;
    const site = (wb.SheetNames.includes('부자재 의왕') || /의왕/.test(title)) ? 'Uiwang' : 'Sokcho';
    const items = {}; let total = null, started = false;
    for (const r of tot) {
      const t = r.map(S);
      if (t.includes('채널') && t.includes('금액')) { started = true; continue; }
      if (!started) continue;
      const name = t.find(x => x); const val = r.find(isNum);
      if (!name || val === undefined) continue;
      if (/총\s*합계/.test(name)) { total = val; break; }
      items[name] = (items[name] || 0) + val;
    }
    if (total === null) total = Object.values(items).reduce((a, b) => a + b, 0);
    // consignment (위탁) by channel, and the year
    const consignment = {}; let year = null;
    const wt = rowsOf(XLSX, wb.Sheets['위탁']); const h = findHeader(wt, ['상품명', '수량', '금액'], null, 5);
    if (h) for (let i = h.idx + 1; i < wt.length; i++) {
      const r = wt[i]; const amt = r[h.map['금액']]; if (!isNum(amt) || !S(r[h.map['상품명']])) continue;
      const ch = S(r[h.map['채널']]) || '(no channel)'; consignment[ch] = (consignment[ch] || 0) + amt;
      if (!year && '출고일자' in h.map) { const d = toISO(XLSX, r[h.map['출고일자']]); if (d) year = d.slice(0, 4); }
    }
    // 경한글로벌 at factory price (+VAT), so it can be compared on the same basis
    let gyeonghanCost = null;
    if (wb.SheetNames.includes('경한글로벌')) {
      for (const r of rowsOf(XLSX, wb.Sheets['경한글로벌'])) {
        const t = r.map(S); const j = t.indexOf('(+vat)');
        if (j >= 0) { const v = r.slice(j + 1).find(isNum); if (v) { gyeonghanCost = v; break; } }
      }
    }
    year = year || String(new Date().getFullYear());
    return { month: `${year}-${pad(mm[1])}`, site, items, total, consignment, gyeonghanCost, fileName };
  }
  // Map a settlement onto Retail / 택배 / 기타 on the factory-price basis.
  function settlementGroups(st, channelTypes) {
    const g = { Retail: 0, Parcel: 0, Other: 0, NotInOrders: 0 }; const detail = [];
    const add = (grp, label, v, note) => { g[grp] += v; detail.push({ group: grp, label, value: v, note: note || '' }); };
    for (const [k, v] of Object.entries(st.items)) {
      if (/쿠팡/.test(k)) add('Retail', k, v);
      else if (/[Bb]마트|비마트/.test(k)) add('Retail', k, v);
      else if (/경한/.test(k)) {
        if (st.gyeonghanCost) { add('Retail', k + ' (at factory price)', st.gyeonghanCost, 'Settlement ' + Math.round(v).toLocaleString('ko-KR') + ' – compared at factory price'); add('NotInOrders', k + ' margin above factory price', v - st.gyeonghanCost); }
        else add('Retail', k, v);
      }
      else if (/^위탁/.test(k)) {
        let b2b = 0;
        for (const [ch, a] of Object.entries(st.consignment)) { const t = channelTypes[ch] || defaultChannelType(ch); if (t === 'B2B') b2b += a; }
        add('Parcel', '위탁', v, b2b ? 'Includes B2B / 벌크 / 오프라인 ' + Math.round(b2b).toLocaleString('ko-KR') : '');
      }
      else if (/N배송|수출|시장|직원/.test(k)) add('Other', k, v);
      else add('NotInOrders', k, v, /부자재/.test(k) ? 'Packaging fee – not in order files' : (/조정/.test(k) ? 'Management adjustment' : ''));
    }
    return { groups: g, detail };
  }
  root.Tally = { parseWorkbook, compute, priceFor, normName, defaultChannelType, ORDER_RULES, hash, settlementGroups };
})(typeof window !== 'undefined' ? window : globalThis);
