/* global Papa, Chart, pako */

// 分享链接：若 URL 含 #r=xxx，则解码并展示报告（与本地结论一致）
(function checkReportHash() {
  const hash = location.hash;
  const m = hash && hash.startsWith('#r=') ? hash.slice(3) : null;
  if (!m || typeof pako === 'undefined') return;
  try {
    const binary = atob(m.replace(/-/g, '+').replace(/_/g, '/'));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const decompressed = pako.ungzip(bytes);
    const html = new TextDecoder().decode(decompressed);
    document.open();
    document.write(html);
    document.close();
    window.__reportView = true;
  } catch (_) {}
})();

const $ = (id) => document.getElementById(id);

const state = {
  rows: [],
  lastFileName: null,
  charts: {
    opsTrend: null,
    sourceTrend: null,
    placementTrend: null,
    latestPosQuadrant: null,
  },
  selectedSources: new Set(),
  selectedPlacements: new Set(),
  topPlacementFilter: '(全部)',
  topSourceFilter: '(全部)',
  topProjectsShowAll: false,
  topProjectsByWeek: false,
  latestWeekRange: null,
  boundDirHandle: null,
};

/** 汇总页「导出整站快照」iframe 读取（const state 不会挂到 window） */
if (typeof window !== 'undefined') window.__SNAPSHOT_OPS__ = state;

const DB_NAME = 'ops-dashboard-local-db';
const DB_STORE = 'kv';
const DB_KEY_DIR_HANDLE = 'boundDirHandle';

const OPS_SUBDIR_CANDIDATES = ['运营宣推', '运营宣推大盘', '运营宣推上'];

const COLS = {
  gran: '时间维度',
  date: '日期',
  topicId: '专题id',
  topicName: '专题名称',
  source: '宣发来源',
  placement: '资源位',

  g_imp: '全局曝光uv',
  g_clk: '全局点击uv',
  g_aread: '全局实际阅读uv',
  g_payu: '全局付费用户',
  g_rev: '全局付费解锁收入',

  ops_imp: '全局曝光uv-运营宣推',
  ops_clk: '全局点击uv-运营宣推',
  ops_read: '运营宣推贡献-阅读uv',
  ops_payu: '运营宣推贡献-付费用户',
  ops_rev: '运营宣推贡献-付费解锁收入',
};

function num(v) {
  if (v == null) return 0;
  const s = String(v).trim();
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function safeDiv(a, b) {
  const den = Number(b);
  if (!Number.isFinite(den) || den === 0) return 0;
  return Number(a) / den;
}

function fmtInt(n) {
  const x = Math.round(Number(n) || 0);
  return x.toLocaleString('zh-CN');
}

function fmtMoney(n, digits = 2) {
  const x = Number(n) || 0;
  return x.toLocaleString('zh-CN', { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

function fmtRate(x, digits = 2) {
  return `${(Number(x) * 100).toFixed(digits)}%`;
}

function pickScopeCols(scope) {
  if (scope === 'ops') {
    return { imp: COLS.ops_imp, clk: COLS.ops_clk, read: COLS.ops_read, payu: COLS.ops_payu, rev: COLS.ops_rev };
  }
  return { imp: COLS.g_imp, clk: COLS.g_clk, read: COLS.g_aread, payu: COLS.g_payu, rev: COLS.g_rev };
}

function rowGranValue(row) {
  if (Object.prototype.hasOwnProperty.call(row, COLS.gran)) return String(row[COLS.gran] ?? '').trim();
  const alt = `\ufeff${COLS.gran}`;
  if (Object.prototype.hasOwnProperty.call(row, alt)) return String(row[alt] ?? '').trim();
  return '';
}

function isDailyRow(row) {
  const g = rowGranValue(row);
  if (!g) return true;
  return g === '日';
}

function computeFunnelAgg(rows, scopeCols) {
  const imp = rows.reduce((s, r) => s + num(r[scopeCols.imp]), 0);
  const clk = rows.reduce((s, r) => s + num(r[scopeCols.clk]), 0);
  const read = rows.reduce((s, r) => s + num(r[scopeCols.read]), 0);
  const payu = rows.reduce((s, r) => s + num(r[scopeCols.payu]), 0);
  const rev = rows.reduce((s, r) => s + num(r[scopeCols.rev]), 0);
  return { imp, clk, read, payu, rev };
}

function computeDerived({ imp, clk, read, payu, rev }) {
  return {
    ctr: safeDiv(clk, imp),
    read_rate: safeDiv(read, clk),
    pay_rate: safeDiv(payu, read),
    rev_per_imp: safeDiv(rev, imp),
  };
}

function groupBy(rows, keyFn) {
  const m = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return m;
}

function buildTable(el, columns, data) {
  if (!el) return;
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  for (const c of columns) {
    const th = document.createElement('th');
    th.textContent = c.label;
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  const tbody = document.createElement('tbody');
  data.forEach((row, i) => {
    const tr = document.createElement('tr');
    for (const c of columns) {
      const td = document.createElement('td');
      const v = typeof c.value === 'function' ? c.value(row, i, data) : row[c.value];
      td.innerHTML = typeof v === 'string' ? v : (v ?? '');
      if (c.className) td.className = c.className;
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });

  el.innerHTML = '';
  el.appendChild(thead);
  el.appendChild(tbody);
}

function kpi(label, value) {
  return `
    <div class="kpi">
      <div class="kpi__label">${label}</div>
      <div class="kpi__value">${value}</div>
    </div>
  `;
}

function currentControls() {
  return {
    scope: $('metricScope').value,
    effMetric: $('effMetric').value,
    minImp: Number($('minImp').value || 0),
  };
}

function effLabel(metric) {
  switch (metric) {
    case 'ctr': return 'CTR';
    case 'read_rate': return '阅读率';
    case 'pay_rate': return '付费率';
    case 'rev_per_imp': return '每曝光收入';
    default: return metric;
  }
}

function filterRowsByDate(rows, startDate, endDate) {
  let s = startDate || '';
  let e = endDate || '';
  if (s && e && s > e) {
    const t = s;
    s = e;
    e = t;
  }
  return rows.filter((r) => {
    const d = String(r[COLS.date] ?? '').trim();
    if (!d) return false;
    if (s && d < s) return false;
    if (e && d > e) return false;
    return true;
  });
}

function daysBetweenISO(a, b) {
  if (!a || !b) return null;
  const da = new Date(`${a}T00:00:00`);
  const db = new Date(`${b}T00:00:00`);
  if (Number.isNaN(da.getTime()) || Number.isNaN(db.getTime())) return null;
  return Math.round((db - da) / (24 * 3600 * 1000));
}

function wow(curr, prev) {
  if (!prev || prev === 0) return null;
  return (curr - prev) / prev;
}

function fmtWow(x) {
  if (x == null || !Number.isFinite(x)) return 'N/A';
  const sign = x >= 0 ? '+' : '';
  return `${sign}${(x * 100).toFixed(2)}%`;
}

function openLocalDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbSet(key, value) {
  const db = await openLocalDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite');
    tx.objectStore(DB_STORE).put(value, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

async function dbGet(key) {
  const db = await openLocalDb();
  const val = await new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readonly');
    const req = tx.objectStore(DB_STORE).get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  db.close();
  return val;
}


function parseCsvRows(text) {
  return new Promise((resolve, reject) => {
    if (typeof Papa === 'undefined') {
      reject(new Error('CSV parser unavailable'));
      return;
    }
    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      complete: (res) => resolve(res.data || []),
      error: (err) => reject(err),
    });
  });
}

function rowDedupKey(r) {
  const keys = [
    '\ufeff时间维度', '时间维度', '日期', '专题id', '专题名称', '宣发来源', '资源位',
    '全局曝光uv', '全局点击uv', '全局实际阅读uv', '全局付费用户', '全局付费解锁收入',
    '全局曝光uv-运营宣推', '全局点击uv-运营宣推', '运营宣推贡献-阅读uv',
    '运营宣推贡献-付费用户', '运营宣推贡献-付费解锁收入',
  ];
  return keys.map((k) => String(r[k] ?? '')).join('|');
}

function mergeRowsDedup(rows) {
  const map = new Map();
  for (const r of rows) {
    const k = rowDedupKey(r);
    if (!map.has(k)) map.set(k, r);
  }
  return Array.from(map.values());
}

function render() {
  const { scope, effMetric, minImp } = currentControls();
  const scopeCols = pickScopeCols(scope);
  const dailyRows = state.rows.filter(isDailyRow);

  if (!dailyRows.length) {
    $('statusHint').textContent = '未读取到有效数据行（请确认 CSV）。';
    return;
  }

  $('statusHint').textContent = `已加载：${state.lastFileName || 'CSV'}；数据行：${dailyRows.length.toLocaleString('zh-CN')}。`;

  const allDates = Array.from(new Set(dailyRows.map((r) => String(r[COLS.date] ?? '').trim()).filter(Boolean))).sort();
  const minDate = allDates[0] || '';
  const maxDate = allDates[allDates.length - 1] || '';
  const recent7Start = allDates[Math.max(0, allDates.length - 7)] || minDate;

  const dateInputDefaults = [
    ['latestStartDate', minDate], ['latestEndDate', maxDate],
    ['opsStartDate', minDate], ['opsEndDate', maxDate],
    ['srcStartDate', minDate], ['srcEndDate', maxDate],
    ['posStartDate', minDate], ['posEndDate', maxDate],
    ['topStartDate', recent7Start], ['topEndDate', maxDate],
    ['matrixStartDate', minDate], ['matrixEndDate', maxDate],
  ];
  dateInputDefaults.forEach(([id, def]) => {
    const el = $(id);
    if (!el) return;
    el.min = minDate;
    el.max = maxDate;
    if (!el.value) el.value = def;
  });

  // Helpers：自然周 = 周一至周日 7 天为一周，周起始为周一
  const toWeekStartLocal = (dateStr) => {
    const d = new Date(`${dateStr}T00:00:00`);
    if (Number.isNaN(d.getTime())) return dateStr;
    const day = d.getDay(); // 0=周日,1=周一,...,6=周六
    const deltaToMon = day === 0 ? -6 : 1 - day; // 归一到当周周一
    d.setDate(d.getDate() + deltaToMon);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const da = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${da}`;
  };

  const aggregateSeries = (series, gran) => {
    if (gran === 'day') return series;
    const byWeek = new Map();
    for (const x of series) {
      const wk = toWeekStartLocal(x.date);
      if (!byWeek.has(wk)) byWeek.set(wk, []);
      byWeek.get(wk).push(x);
    }
    const weeks = Array.from(byWeek.keys()).sort();
    return weeks.map((wk) => {
      const arr = byWeek.get(wk);
      const days = arr.length;
      const sum = arr.reduce((m, x) => ({
        date: wk,
        days,
        gImp: m.gImp + x.gImp,
        imp: m.imp + x.imp,
        clk: m.clk + x.clk,
        read: m.read + x.read,
        payu: m.payu + x.payu,
        rev: m.rev + x.rev,
      }), { date: wk, days, gImp: 0, imp: 0, clk: 0, read: 0, payu: 0, rev: 0 });
      const derv = computeDerived(sum);
      const opsShare = safeDiv(sum.imp, sum.gImp);
      const base = { date: wk, gImp: sum.gImp, ...sum, ...derv, opsShare, days };
      if (gran === 'week_avg') {
        return {
          ...base,
          gImp: safeDiv(sum.gImp, days),
          imp: safeDiv(sum.imp, days),
          clk: safeDiv(sum.clk, days),
          read: safeDiv(sum.read, days),
          payu: safeDiv(sum.payu, days),
          rev: safeDiv(sum.rev, days),
        };
      }
      return base;
    });
  };

  // 0) 最新周资源效率 + 调配建议（固定运营宣推口径）；结论在 module1 顶部，优先用 ops 日期与图表保持一致
  try {
    const opsStart = $('opsStartDate')?.value || $('latestStartDate')?.value;
    const opsEnd = $('opsEndDate')?.value || $('latestEndDate')?.value;
    const latestRows = filterRowsByDate(dailyRows, opsStart, opsEnd);
    const latestNoDataInRange = !latestRows.length;
    const opsCols = pickScopeCols('ops');
    const byDate = groupBy(latestRows, (r) => String(r[COLS.date] ?? '').trim());
    const dateKeys = Array.from(byDate.keys()).filter(Boolean).sort();

    const weekBucket = new Map(); // weekStart -> {dates:Set, rows:[]}
    for (const d of dateKeys) {
      const wk = toWeekStartLocal(d);
      if (!weekBucket.has(wk)) weekBucket.set(wk, { weekStart: wk, dates: new Set(), rows: [] });
      const b = weekBucket.get(wk);
      b.dates.add(d);
      b.rows.push(...byDate.get(d));
    }

    const weekArr = Array.from(weekBucket.values()).sort((a, b) => a.weekStart.localeCompare(b.weekStart));
    let latestWeek = null;
    const fullWeeks = weekArr.filter((w) => w.dates.size >= 7);
    if (fullWeeks.length) latestWeek = fullWeeks[fullWeeks.length - 1];
    else if (weekArr.length) latestWeek = weekArr[weekArr.length - 1];

    if (!latestWeek) {
      state.latestWeekRange = null;
      $('latestWeekSummary').innerHTML = latestNoDataInRange ? '当前筛选日期范围内没有数据，请调整开始/结束日期。' : '未找到可计算的周数据。';
      const kwEl = $('latestWeekKpis'); if (kwEl) kwEl.innerHTML = '';
      const w1 = $('conclusionSection1Wrap');
      const w2 = $('conclusionSection2Wrap');
      if (w1) w1.style.display = 'none';
      if (w2) w2.style.display = 'none';
      $('latestWeekWowTable').innerHTML = '';
      $('latestWeekTable').innerHTML = '';
      $('latestWeekSourceTable').innerHTML = '';
      const tall = $('latestWeekSourceTableAll');
      if (tall) tall.innerHTML = '';
      const swEl = $('sourceWeeklyTable');
      if (swEl) swEl.innerHTML = '';
      $('latestWeekTopProjectsTable').innerHTML = '';
      if ($('latestPosQuadrantFallback')) $('latestPosQuadrantFallback').innerHTML = '';
      if ($('latestPosQuadrantHint')) $('latestPosQuadrantHint').textContent = '';
      const c2 = state.charts.latestPosQuadrant;
      if (c2) { c2.destroy(); state.charts.latestPosQuadrant = null; }
    } else {
      const days = latestWeek.dates.size;
      const weekRows = latestWeek.rows;
      const weekStart = latestWeek.weekStart;
      const latestWeekIdx = weekArr.findIndex((w) => w.weekStart === latestWeek.weekStart);
      let prevWeek = latestWeekIdx > 0 ? weekArr[latestWeekIdx - 1] : null;
      if (prevWeek) {
        const diff = daysBetweenISO(prevWeek.weekStart, weekStart);
        if (diff !== 7) prevWeek = null;
      }

      // week range
      const ws = new Date(`${weekStart}T00:00:00`);
      const we = new Date(ws);
      we.setDate(ws.getDate() + 6);
      const weekEnd = `${we.getFullYear()}-${String(we.getMonth() + 1).padStart(2, '0')}-${String(we.getDate()).padStart(2, '0')}`;
      state.latestWeekRange = { start: weekStart, end: weekEnd };

      const gImp = weekRows.reduce((s, r) => s + num(r[COLS.g_imp]), 0);
      const a = computeFunnelAgg(weekRows, opsCols);
      const d = computeDerived(a);
      const opsShareGlobal = safeDiv(a.imp, gImp);

      let prevA = null;
      let prevD = null;
      let prevOpsShareGlobal = null;
      if (prevWeek) {
        const pgImp = prevWeek.rows.reduce((s, r) => s + num(r[COLS.g_imp]), 0);
        prevA = computeFunnelAgg(prevWeek.rows, opsCols);
        prevD = computeDerived(prevA);
        prevOpsShareGlobal = safeDiv(prevA.imp, pgImp);
      }

      const kwEl = $('latestWeekKpis');
      if (kwEl) kwEl.innerHTML = [
        kpi('周区间', `<span class="num">${weekStart} ~ ${weekEnd}</span>${days < 7 ? '<span class="muted">（非完整周）</span>' : ''}`),
        kpi('运营宣推曝光占比（周）', fmtRate(opsShareGlobal, 2)),
        kpi('运营宣推每曝光收入（周）', `<span class="num">${fmtMoney(d.rev_per_imp, 4)}</span>`),
        kpi('运营宣推p-CTR（周）', fmtRate(d.ctr, 2)),
      ].join('');

      const wowRows = [
        { metric: '每曝光收入', curr: d.rev_per_imp, prev: prevD?.rev_per_imp, wow: wow(d.rev_per_imp, prevD?.rev_per_imp) },
        { metric: 'p-CTR', curr: d.ctr, prev: prevD?.ctr, wow: wow(d.ctr, prevD?.ctr) },
        { metric: '阅读率', curr: d.read_rate, prev: prevD?.read_rate, wow: wow(d.read_rate, prevD?.read_rate) },
        { metric: '付费率', curr: d.pay_rate, prev: prevD?.pay_rate, wow: wow(d.pay_rate, prevD?.pay_rate) },
        { metric: '运营宣推曝光占比', curr: opsShareGlobal, prev: prevOpsShareGlobal, wow: wow(opsShareGlobal, prevOpsShareGlobal) },
      ];
      const wowCols = [
        { label: '指标', value: (r) => r.metric },
        { label: '当前周', className: 'num', value: (r) => (r.metric === '每曝光收入' ? fmtMoney(r.curr, 4) : fmtRate(r.curr, 2)) },
        { label: '上周', className: 'num', value: (r) => {
          if (r.prev == null) return 'N/A';
          return r.metric === '每曝光收入' ? fmtMoney(r.prev, 4) : fmtRate(r.prev, 2);
        } },
        { label: '环比', className: 'num', value: (r) => fmtWow(r.wow) },
      ];
      buildTable($('latestWeekWowTable'), wowCols, wowRows);

      // 资源位为「其它」等的不在数据结论范围内
      const excludePlacement = new Set(['其它', '其他', '(空)']);
      const weekRowsForPos = weekRows.filter((r) => !excludePlacement.has(String(r[COLS.placement] ?? '').trim() || '(空)'));

      // by placement
      const byPos = groupBy(weekRowsForPos, (r) => String(r[COLS.placement] ?? '').trim() || '(空)');
      const posRows = Array.from(byPos.entries()).map(([pos, rows]) => {
        const pa = computeFunnelAgg(rows, opsCols);
        const pd = computeDerived(pa);
        const shareOps = safeDiv(pa.imp, a.imp);
        let tag = '保持观察';
        if (pd.rev_per_imp >= d.rev_per_imp * 1.2 && shareOps >= 0.05) tag = '高效&有量（建议加量）';
        else if (pd.rev_per_imp <= d.rev_per_imp * 0.7 && shareOps >= 0.05) tag = '低效大盘（建议挪量）';
        else if (pd.rev_per_imp >= d.rev_per_imp * 1.2) tag = '高效小盘（可试探加量）';
        return { pos, ...pa, ...pd, shareOps, tag };
      }).sort((x, y) => y.rev_per_imp - x.rev_per_imp);

      // source wow
      const srcCurrMap = new Map();
      const srcPrevMap = new Map();
      const bySrcCurr = groupBy(weekRows, (r) => String(r[COLS.source] ?? '').trim() || '(空)');
      bySrcCurr.forEach((rows, src) => {
        const sa = computeFunnelAgg(rows, opsCols);
        srcCurrMap.set(src, { ...sa, ...computeDerived(sa) });
      });
      if (prevWeek) {
        const bySrcPrev = groupBy(prevWeek.rows, (r) => String(r[COLS.source] ?? '').trim() || '(空)');
        bySrcPrev.forEach((rows, src) => {
          const sa = computeFunnelAgg(rows, opsCols);
          srcPrevMap.set(src, { ...sa, ...computeDerived(sa) });
        });
      }
      const srcTop = Array.from(srcCurrMap.entries())
        .map(([src, m]) => ({ src, ...m }))
        .sort((x, y) => y.imp - x.imp)
        .slice(0, 5);

      const srcAllWow = Array.from(srcCurrMap.entries())
        .map(([src, m]) => {
          const p = srcPrevMap.get(src);
          return {
            src,
            pctr: m.ctr,
            pctr_wow: wow(m.ctr, p?.ctr),
            erpi: m.rev_per_imp,
            erpi_wow: wow(m.rev_per_imp, p?.rev_per_imp),
            imp: m.imp,
          };
        })
        .sort((x, y) => y.imp - x.imp);

      const srcCols = [
        { label: '宣发来源', value: (r) => r.src },
        { label: '运营宣推曝光', className: 'num', value: (r) => fmtInt(r.imp) },
        { label: 'p-CTR', className: 'num', value: (r) => fmtRate(r.pctr, 2) },
        { label: 'p-CTR环比', className: 'num', value: (r) => fmtWow(r.pctr_wow) },
        { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.erpi, 4) },
        { label: '每曝光收入环比', className: 'num', value: (r) => fmtWow(r.erpi_wow) },
      ];

      // 宣发来源周度明细表（多周 × 曝光量级pv / p-ctr / 每曝光收入 + 周环比）
      const swEl = $('sourceWeeklyTable');
      if (swEl && weekArr.length >= 1) {
        const weeksToShow = weekArr.slice(-2).reverse();
        const weekLabel = (wk) => {
          const m = wk.slice(5, 7);
          const d = wk.slice(8, 10);
          return `${m}${d}周`;
        };
        const byWeekBySrc = new Map();
        for (const w of weekArr.slice(-2)) {
          const bySrc = groupBy(w.rows, (r) => String(r[COLS.source] ?? '').trim() || '(空)');
          const totalImp = w.rows.reduce((s, r) => s + num(r[opsCols.imp]), 0);
          const srcMap = new Map();
          bySrc.forEach((rows, src) => {
            const a = computeFunnelAgg(rows, opsCols);
            const d = computeDerived(a);
            const shareOps = safeDiv(a.imp, totalImp);
            srcMap.set(src, { imp: a.imp, ctr: d.ctr, rev_per_imp: d.rev_per_imp, shareOps });
          });
          const totalA = computeFunnelAgg(w.rows, opsCols);
          const totalD = computeDerived(totalA);
          srcMap.set('运营全局', { imp: totalImp, ctr: totalD.ctr, rev_per_imp: totalD.rev_per_imp, shareOps: 1 });
          byWeekBySrc.set(w.weekStart, srcMap);
        }
        const latestWk = weeksToShow[0];
        const prevWk = weeksToShow.length >= 2 ? weeksToShow[1] : null;
        const allSources = new Set();
        weekArr.slice(-2).forEach((w) => {
          const bySrc = groupBy(w.rows, (r) => String(r[COLS.source] ?? '').trim() || '(空)');
          bySrc.forEach((_, src) => allSources.add(src));
        });
        const otherSources = Array.from(allSources).filter((s) => s !== '运营全局');
        const latestSrcMap = byWeekBySrc.get(latestWk.weekStart);
        const prevSrcMap = prevWk ? byWeekBySrc.get(prevWk.weekStart) : null;
        otherSources.sort((a, b) => {
          const impA = latestSrcMap?.get(a)?.imp ?? 0;
          const impB = latestSrcMap?.get(b)?.imp ?? 0;
          return impB - impA;
        });
        const sourceOrder = ['运营全局', ...otherSources];

        let maxAbsChangeSrc = null;
        let maxAbsChangeVal = -1;
        if (prevWk && prevSrcMap) {
          for (const src of otherSources) {
            const currM = latestSrcMap?.get(src);
            const prevM = prevSrcMap?.get(src);
            if (currM && prevM) {
              const impWow = wow(currM.imp, prevM.imp);
              const ctrWow = wow(currM.ctr, prevM.ctr);
              const revWow = wow(currM.rev_per_imp, prevM.rev_per_imp);
              const shareWow = wow(currM.shareOps, prevM.shareOps);
              const absMax = Math.max(
                impWow != null ? Math.abs(impWow) : 0,
                shareWow != null ? Math.abs(shareWow) : 0,
                ctrWow != null ? Math.abs(ctrWow) : 0,
                revWow != null ? Math.abs(revWow) : 0,
              );
              if (absMax > maxAbsChangeVal) {
                maxAbsChangeVal = absMax;
                maxAbsChangeSrc = src;
              }
            }
          }
        }

        const cellCls = (curr, prev) => {
          if (prev == null) return '';
          return curr > prev ? 'bad' : curr < prev ? 'good' : '';
        };

        let html = '<thead><tr><th rowspan="2">宣发来源</th>';
        weeksToShow.forEach((w) => {
          html += `<th colspan="4">${weekLabel(w.weekStart)}</th>`;
        });
        html += '<th colspan="5">周环比</th></tr><tr>';
        weeksToShow.forEach(() => {
          html += '<th class="num">曝光量级pv</th><th class="num">宣推占比</th><th class="num">p-ctr</th><th class="num">每曝光收入</th>';
        });
        html += '<th class="num">曝光量级pv</th><th class="num">宣推占比</th><th class="num">p-ctr</th><th class="num">每曝光收入</th><th class="num">每曝光收入绝对值</th></tr></thead><tbody>';

        for (const src of sourceOrder) {
          const rowHighlight = src === maxAbsChangeSrc ? 'rowHighlightMax' : '';
          html += `<tr${rowHighlight ? ` class="${rowHighlight}"` : ''}><td><strong>${src}</strong></td>`;
          weeksToShow.forEach((w, wi) => {
            const m = byWeekBySrc.get(w.weekStart)?.get(src);
            if (m) {
              const isLatest = wi === 0 && prevWk;
              const prevM = prevWk && wi === 0 ? prevSrcMap?.get(src) : null;
              let impSp = fmtInt(m.imp);
              let shareSp = fmtRate(m.shareOps, 2);
              let ctrSp = fmtRate(m.ctr, 2);
              let revSp = fmtMoney(m.rev_per_imp, 4);
              if (isLatest && prevM) {
                const impCls = cellCls(m.imp, prevM.imp);
                const shareCls = cellCls(m.shareOps, prevM.shareOps);
                const ctrCls = cellCls(m.ctr, prevM.ctr);
                const revCls = cellCls(m.rev_per_imp, prevM.rev_per_imp);
                if (impCls) impSp = `<span class="${impCls}">${impSp}</span>`;
                if (shareCls) shareSp = `<span class="${shareCls}">${shareSp}</span>`;
                if (ctrCls) ctrSp = `<span class="${ctrCls}">${ctrSp}</span>`;
                if (revCls) revSp = `<span class="${revCls}">${revSp}</span>`;
              }
              html += `<td class="num">${impSp}</td><td class="num">${shareSp}</td><td class="num">${ctrSp}</td><td class="num">${revSp}</td>`;
            } else {
              html += '<td class="num">-</td><td class="num">-</td><td class="num">-</td><td class="num">-</td>';
            }
          });
          if (prevWk) {
            const currM = latestSrcMap?.get(src);
            const prevM = prevSrcMap?.get(src);
            if (currM && prevM) {
              const impWow = wow(currM.imp, prevM.imp);
              const ctrWow = wow(currM.ctr, prevM.ctr);
              const revWow = wow(currM.rev_per_imp, prevM.rev_per_imp);
              const shareWow = wow(currM.shareOps, prevM.shareOps);
              const revAbs = currM.rev_per_imp - prevM.rev_per_imp;
              html += `<td class="num">${fmtWow(impWow)}</td><td class="num">${fmtWow(shareWow)}</td><td class="num">${fmtWow(ctrWow)}</td><td class="num">${fmtWow(revWow)}</td><td class="num">${revAbs >= 0 ? '+' : ''}${fmtMoney(revAbs, 4)}</td>`;
            } else {
              html += '<td class="num">-</td><td class="num">-</td><td class="num">-</td><td class="num">-</td><td class="num">-</td>';
            }
          } else {
            html += '<td class="num">-</td><td class="num">-</td><td class="num">-</td><td class="num">-</td><td class="num">-</td>';
          }
          html += '</tr>';
        }
        html += '</tbody>';
        swEl.innerHTML = html;
      }

      // resource usage summary top5 by exposure
      const posUsageTop = [...posRows].sort((x, y) => y.imp - x.imp).slice(0, 5);
      const prevPosMap = new Map();
      if (prevWeek) {
        const prevWeekRowsForPos = prevWeek.rows.filter((r) => !excludePlacement.has(String(r[COLS.placement] ?? '').trim() || '(空)'));
        const byPosPrev = groupBy(prevWeekRowsForPos, (r) => String(r[COLS.placement] ?? '').trim() || '(空)');
        byPosPrev.forEach((rows, pos) => {
          const pa = computeFunnelAgg(rows, opsCols);
          const pd = computeDerived(pa);
          prevPosMap.set(pos, { ...pa, ...pd });
        });
      }
      const posAllWow = [...posRows].map((x) => {
        const p = prevPosMap.get(x.pos);
        return {
          pos: x.pos,
          imp: x.imp,
          shareOps: x.shareOps,
          pctr: x.ctr,
          pctr_wow: wow(x.ctr, p?.ctr),
          erpi: x.rev_per_imp,
          erpi_wow: wow(x.rev_per_imp, p?.rev_per_imp),
          tag: x.tag,
        };
      }).sort((a, b) => b.imp - a.imp);

      // 资源位表格：资源位、曝光量级、曝光占比、pctr、每曝光收入、周环比、本周建议（曝光占比=各资源位运营宣推/运营宣推总曝光）
      const weekCols = [
        { label: '资源位', value: (r) => r.pos },
        { label: '曝光量级', className: 'num', value: (r) => fmtInt(r.imp) },
        { label: '曝光占比', className: 'num', value: (r) => fmtRate(r.shareOps ?? 0, 2) },
        { label: 'p-CTR', className: 'num', value: (r) => fmtRate(r.pctr ?? r.ctr, 2) },
        { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.erpi ?? r.rev_per_imp, 4) },
        { label: '周环比', className: 'num', value: (r) => (r.erpi_wow != null ? fmtWow(r.erpi_wow) : '-') },
        { label: '本周建议', value: (r) => r.tag ?? '' },
      ];
      buildTable($('latestWeekTable'), weekCols, posAllWow);

      // top5 projects: 按(项目,宣发来源,资源位)粒度，含效率表现（自身历史+同来源同资源位对比，曝光量级归一）
      const excludeSrc = new Set(['其它', '其他', '(空)']);
      const weekRowsForTop5 = weekRows.filter(
        (r) => !excludePlacement.has(String(r[COLS.placement] ?? '').trim() || '(空)')
      );
      const byTopicSrcPos = groupBy(weekRowsForTop5, (r) => {
        const tid = String(r[COLS.topicId] ?? '').trim() || String(r[COLS.topicName] ?? '').trim();
        const name = String(r[COLS.topicName] ?? '').trim() || '(空)';
        const src = String(r[COLS.source] ?? '').trim() || '(空)';
        const pos = String(r[COLS.placement] ?? '').trim() || '(空)';
        return `${tid}||${name}||${src}||${pos}`;
      });
      let topicTopRaw = Array.from(byTopicSrcPos.entries()).map(([key, rows]) => {
        const ta = computeFunnelAgg(rows, opsCols);
        const td = computeDerived(ta);
        const [, name, src, pos] = key.split('||');
        const tid = key.split('||')[0];
        return { tid, name, src, pos, ...ta, ...td };
      }).filter((x) => x.imp >= Math.max(10000, Math.floor(minImp / 2)))
        .filter((x) => !excludeSrc.has(x.src))
        .sort((x, y) => y.rev_per_imp - x.rev_per_imp);

      // 效率表现：自身历史均值 + 同来源同资源位同曝光档位对比
      const getEfficiencyLabel = (item) => {
        const impLo = item.imp * 0.5;
        const impHi = item.imp * 2;
        const histErpis = [];
        let peerErpis = [];
        for (const w of weekArr) {
          if (w.weekStart === weekStart) continue;
          const matches = w.rows.filter((r) => {
            if (excludePlacement.has(String(r[COLS.placement] ?? '').trim() || '(空)')) return false;
            const n = String(r[COLS.topicName] ?? '').trim();
            const p = String(r[COLS.placement] ?? '').trim() || '(空)';
            return (n === item.name || String(r[COLS.topicId] ?? '').trim() === item.tid) && p === item.pos;
          });
          if (matches.length) {
            const agg = computeFunnelAgg(matches, opsCols);
            const der = computeDerived(agg);
            if (agg.imp > 0) histErpis.push(der.rev_per_imp);
          }
        }
        const histAvg = histErpis.length ? histErpis.reduce((s, v) => s + v, 0) / histErpis.length : null;
        const byPeer = groupBy(
          weekRowsForTop5.filter((r) => String(r[COLS.source] ?? '').trim() === item.src && (String(r[COLS.placement] ?? '').trim() || '(空)') === item.pos),
          (r) => String(r[COLS.topicName] ?? '').trim()
        );
        byPeer.forEach((rows, name) => {
          if (name === item.name) return;
          const agg = computeFunnelAgg(rows, opsCols);
          if (agg.imp >= impLo && agg.imp <= impHi && agg.imp > 0) peerErpis.push(agg.rev / agg.imp);
        });
        const cur = item.rev_per_imp;
        const hist = histAvg != null && histAvg > 0 ? (cur >= histAvg * 1.05 ? '高于自身历史' : cur <= histAvg * 0.95 ? '低于自身历史' : '持平自身历史') : null;
        const peerP50 = peerErpis.length ? peerErpis.sort((a, b) => a - b)[Math.floor(peerErpis.length / 2)] : null;
        const peer = peerP50 != null && peerP50 > 0 ? (cur >= peerP50 * 1.05 ? '高于同档位P50' : cur <= peerP50 * 0.95 ? '低于同档位P50' : '持平同档位P50') : null;
        const parts = [hist, peer].filter(Boolean);
        return parts.length ? parts.join('；') : '-';
      };

      const topicTop = topicTopRaw.slice(0, 5).map((x) => ({ ...x, effLabel: getEfficiencyLabel(x) }));
      const topicCols = [
        { label: '项目名', value: (r) => r.name },
        { label: '宣发来源', value: (r) => r.src },
        { label: '资源场景', value: (r) => r.pos },
        { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.rev_per_imp, 4) },
        { label: 'p-CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
        { label: '效率表现', value: (r) => r.effLabel || '-' },
      ];
      buildTable($('latestWeekTopProjectsTable'), topicCols, topicTop);

      buildTable($('latestWeekSourceTableAll'), srcCols, srcAllWow);

      // 资源位效率-规模象限图（最新周）
      const quadEl = $('latestPosQuadrantChart');
      const quadHintEl = $('latestPosQuadrantHint');
      const quadFallbackEl = $('latestPosQuadrantFallback');
      if (quadFallbackEl) {
        quadFallbackEl.innerHTML = `<div class="fallbackList">${
          posRows.slice(0, 10).map((x) =>
            `<div class="fallbackItem"><span>${x.pos}</span><span>曝光占比 ${fmtRate(x.shareOps,2)} ｜ 每曝光收入 ${fmtMoney(x.rev_per_imp,4)}</span></div>`
          ).join('')
        }</div>`;
      }

      if (typeof Chart !== 'undefined' && quadEl) {
        const points = posRows.slice(0, 12).map((x) => ({
          x: x.shareOps,
          y: x.rev_per_imp,
          r: Math.max(4, Math.min(18, Math.sqrt(x.imp || 0) / 120)),
          pos: x.pos,
          imp: x.imp,
        }));
        const ctx = quadEl.getContext('2d');
        if (state.charts.latestPosQuadrant) state.charts.latestPosQuadrant.destroy();
        state.charts.latestPosQuadrant = new Chart(ctx, {
          type: 'bubble',
          data: { datasets: [{ label: '资源位', data: points, backgroundColor: 'rgba(79,124,255,.45)', borderColor: 'rgba(79,124,255,.9)' }] },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              x: { title: { display: true, text: '曝光占比' }, ticks: { callback: (v) => `${Number(v * 100).toFixed(0)}%` }, grid: { color: 'rgba(15,23,42,.06)' } },
              y: { title: { display: true, text: '每曝光收入' }, ticks: { callback: (v) => Number(v).toFixed(4) }, grid: { color: 'rgba(15,23,42,.06)' } },
            },
            plugins: {
              tooltip: {
                callbacks: {
                  label: (ctx2) => {
                    const p = ctx2.raw || {};
                    return `${p.pos || ''} | 占比${fmtRate(p.x || 0, 2)} | 每曝光收入${fmtMoney(p.y || 0, 4)} | 曝光${fmtInt(p.imp || 0)}`;
                  },
                },
              },
            },
          },
        });
        if (quadHintEl) {
          quadHintEl.innerHTML = '横轴=曝光占比，纵轴=每曝光收入，气泡大小≈曝光量。<br/>' +
            '<span class="muted">四象限建议：右上(高占比+高效率)=维持/加量；左上(低占比+高效率)=可加量；右下(高占比+低效率)=优先挪量；左下(低占比+低效率)=观察或收缩。</span>';
        }
        if (quadFallbackEl) quadFallbackEl.style.display = 'none';
      } else {
        if (quadHintEl) quadHintEl.textContent = '图表库未加载，已显示列表版资源位效率。';
        if (quadFallbackEl) quadFallbackEl.style.display = '';
      }

      const high = posRows.filter((r) => r.tag.includes('高效')).slice(0, 3);
      const low = posRows.filter((r) => r.tag.includes('低效')).slice(0, 3);

      // 波动阈值：1和2无明显波动则不展示
      const wowRev = wow(d.rev_per_imp, prevD?.rev_per_imp);
      const wowCtr = wow(d.ctr, prevD?.ctr);
      const wowShare = wow(opsShareGlobal, prevOpsShareGlobal);
      const hasSignificantWow = prevD && [wowRev, wowCtr, wowShare].some((v) => v != null && Math.abs(v) >= 0.05);
      const srcTopWithWow = srcTop.map((x) => {
        const p = srcPrevMap.get(x.src);
        return { ...x, pctr_wow: wow(x.ctr, p?.ctr), erpi_wow: wow(x.rev_per_imp, p?.rev_per_imp) };
      });
      const hasSignificantSrc = srcTopWithWow.some((x) =>
        (x.pctr_wow != null && Math.abs(x.pctr_wow) >= 0.08) || (x.erpi_wow != null && Math.abs(x.erpi_wow) >= 0.08)
      );

      const wrap1 = $('conclusionSection1Wrap');
      const wrap2 = $('conclusionSection2Wrap');
      if (wrap1) wrap1.style.display = hasSignificantWow ? '' : 'none';
      if (wrap2) wrap2.style.display = hasSignificantSrc ? '' : 'none';
      if (hasSignificantWow) buildTable($('latestWeekWowTable'), wowCols, wowRows);
      if (hasSignificantSrc) {
        const srcColsTop5 = [
          { label: '宣发来源', value: (r) => r.src },
          { label: '运营宣推曝光', className: 'num', value: (r) => fmtInt(r.imp) },
          { label: 'p-CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
          { label: 'p-CTR环比', className: 'num', value: (r) => fmtWow(r.pctr_wow) },
          { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.rev_per_imp, 4) },
          { label: '每曝光收入环比', className: 'num', value: (r) => fmtWow(r.erpi_wow) },
        ];
        buildTable($('latestWeekSourceTable'), srcColsTop5, srcTopWithWow);
      }

      const summaryLines = [];
      summaryLines.push(`<strong>【口径】</strong>以下结论仅针对运营宣推曝光及效率进行分析。`);
      summaryLines.push(`<strong>当前周：</strong>${weekStart} ~ ${weekEnd}${days < 7 ? '（非完整周）' : ''}`);
      if (prevD) {
        const effUp = d.rev_per_imp > prevD.rev_per_imp * 1.02;
        const effDown = d.rev_per_imp < prevD.rev_per_imp * 0.98;
        const shareUp = opsShareGlobal > prevOpsShareGlobal * 1.02;
        const shareDown = opsShareGlobal < prevOpsShareGlobal * 0.98;
        let oneLiner = '最新周大盘';
        if (effUp && shareUp) oneLiner += '效率与曝光占比均向好';
        else if (effUp && shareDown) oneLiner += '效率提升但曝光占比下降';
        else if (effDown && shareUp) oneLiner += '曝光占比上升但效率下降';
        else if (effDown && shareDown) oneLiner += '效率与曝光占比均下降';
        else oneLiner += '效率与曝光占比变化较平稳';
        summaryLines.push(`<strong>【结论】</strong>${oneLiner}。`);
      } else {
        summaryLines.push(`<strong>【结论】</strong>最新周数据已展示（当前仅一周，暂无环比）。`);
      }
      let adviceText = '';
      if (high.length && low.length) {
        adviceText = `优先加量 ${high.map((x) => x.pos).join('、')}；优先整改/挪量 ${low.map((x) => x.pos).join('、')}。`;
      } else if (high.length) {
        adviceText = `优先加量 ${high.map((x) => x.pos).join('、')}；暂无明显低效资源位。`;
      } else if (low.length) {
        adviceText = `暂无高效资源位，建议从曝光占比高的资源位中择优提升；优先整改/挪量 ${low.map((x) => x.pos).join('、')}。`;
      } else {
        adviceText = '资源位效率分布较均衡，建议维持当前配置。';
      }
      summaryLines.push(`<strong>调配建议：</strong>${adviceText} · 结构性挪量：低效→高效转移；效率底线：加量时每曝光收入降幅≤5%~10%。`);

      const wrap = $('conclusionBlock');
      if (wrap) wrap.style.display = '';
      $('latestWeekSummary').innerHTML = summaryLines.join('<br/>');

    }
  } catch (err) {
    state.latestWeekRange = null;
    $('latestWeekSummary').innerHTML = `最新周模块渲染异常：${err?.message || err}`;
    const kwEl = $('latestWeekKpis'); if (kwEl) kwEl.innerHTML = '';
    $('latestWeekWowTable').innerHTML = '';
    $('latestWeekTable').innerHTML = '';
    $('latestWeekSourceTable').innerHTML = '';
    $('latestWeekTopProjectsTable').innerHTML = '';
    if ($('latestPosQuadrantHint')) $('latestPosQuadrantHint').textContent = '图表渲染异常。';
    const c2 = state.charts.latestPosQuadrant;
    if (c2) { c2.destroy(); state.charts.latestPosQuadrant = null; }
  }

  // 1) Ops big盘趋势
  {
    const timeGran = $('opsTimeGran')?.value || 'day';
    const opsRows = filterRowsByDate(dailyRows, $('opsStartDate')?.value, $('opsEndDate')?.value);
    const byDate = groupBy(opsRows, (r) => String(r[COLS.date] ?? '').trim());
    const dates = Array.from(byDate.keys()).filter(Boolean).sort();
    const opsCols = pickScopeCols('ops');

    const dailySeries = dates.map((d) => {
      const rows = byDate.get(d);
      const gImp = rows.reduce((s, r) => s + num(r[COLS.g_imp]), 0);
      const a = computeFunnelAgg(rows, opsCols);
      const derv = computeDerived(a);
      const opsShare = safeDiv(a.imp, gImp);
      return { date: d, gImp, ...a, ...derv, opsShare };
    });

    const total = dailySeries.reduce((m, x) => ({
      gImp: m.gImp + x.gImp,
      imp: m.imp + x.imp,
      clk: m.clk + x.clk,
      read: m.read + x.read,
      payu: m.payu + x.payu,
      rev: m.rev + x.rev,
    }), { gImp: 0, imp: 0, clk: 0, read: 0, payu: 0, rev: 0 });
    const totalDer = computeDerived(total);
    const totalShare = safeDiv(total.imp, total.gImp);

    // 最新周环比：按自然周聚合，取最近两周
    const byWeekOps = new Map();
    for (const x of dailySeries) {
      const wk = toWeekStartLocal(x.date);
      if (!byWeekOps.has(wk)) byWeekOps.set(wk, { gImp: 0, imp: 0, clk: 0, read: 0, payu: 0, rev: 0 });
      const b = byWeekOps.get(wk);
      b.gImp += x.gImp;
      b.imp += x.imp;
      b.clk += x.clk;
      b.read += x.read;
      b.payu += x.payu;
      b.rev += x.rev;
    }
    const weekStarts = Array.from(byWeekOps.keys()).sort();
    const latestWeekAgg = weekStarts.length >= 1 ? byWeekOps.get(weekStarts[weekStarts.length - 1]) : null;
    const prevWeekAgg = weekStarts.length >= 2 ? byWeekOps.get(weekStarts[weekStarts.length - 2]) : null;
    const latestDer = latestWeekAgg ? computeDerived(latestWeekAgg) : null;
    const prevDer = prevWeekAgg ? computeDerived(prevWeekAgg) : null;
    const latestShare = latestWeekAgg ? safeDiv(latestWeekAgg.imp, latestWeekAgg.gImp) : null;
    const prevShare = prevWeekAgg ? safeDiv(prevWeekAgg.imp, prevWeekAgg.gImp) : null;

    const wowShare = latestShare != null && prevShare ? fmtWow(wow(latestShare, prevShare)) : null;
    const wowRevPerImp = latestDer && prevDer ? fmtWow(wow(latestDer.rev_per_imp, prevDer.rev_per_imp)) : null;
    const wowCtr = latestDer && prevDer ? fmtWow(wow(latestDer.ctr, prevDer.ctr)) : null;
    const wowReadRate = latestDer && prevDer ? fmtWow(wow(latestDer.read_rate, prevDer.read_rate)) : null;
    const wowPayRate = latestDer && prevDer ? fmtWow(wow(latestDer.pay_rate, prevDer.pay_rate)) : null;

    const wowSuffix = (v) => (v != null ? ` <span class="muted">（周环比${v}）</span>` : '');

    $('opsKpis').innerHTML = [
      kpi('运营宣推曝光占比（全期）', fmtRate(totalShare, 2) + wowSuffix(wowShare)),
      kpi('运营宣推每曝光收入（全期）', `<span class="num">${fmtMoney(totalDer.rev_per_imp, 4)}</span>${wowSuffix(wowRevPerImp)}`),
      kpi('运营宣推收入（全期）', `<span class="num">${fmtMoney(total.rev, 0)}</span>`),
      kpi('运营宣推CTR（全期）', fmtRate(totalDer.ctr, 2) + wowSuffix(wowCtr)),
      kpi('运营宣推阅读率（全期）', fmtRate(totalDer.read_rate, 2) + wowSuffix(wowReadRate)),
      kpi('运营宣推付费率（全期）', fmtRate(totalDer.pay_rate, 2) + wowSuffix(wowPayRate)),
    ].join('');

    const series = aggregateSeries(dailySeries, timeGran);
    const chartEl = $('opsTrendChart');
    const labels = series.map((x) => x.date);
    const metricKey = effMetric;
    const metricVals = series.map((x) => x[metricKey] ?? 0);
    const shareVals = series.map((x) => x.opsShare ?? 0);

    const hintEl = $('chartHint');
    if (typeof Chart !== 'undefined' && chartEl) {
      const ctx = chartEl.getContext('2d');
      if (state.charts.opsTrend) state.charts.opsTrend.destroy();
      state.charts.opsTrend = new Chart(ctx, {
        type: 'line',
        data: {
          labels,
          datasets: [
            {
              label: `效率：${effLabel(metricKey)}`,
              data: metricVals,
              borderColor: '#4f7cff',
              backgroundColor: 'rgba(79,124,255,.18)',
              tension: 0.25,
              pointRadius: 0,
              borderWidth: 2,
              yAxisID: 'y',
            },
            {
              label: '运营宣推曝光占比',
              data: shareVals,
              borderColor: 'rgba(15,23,42,.35)',
              backgroundColor: 'rgba(15,23,42,.08)',
              tension: 0.25,
              pointRadius: 0,
              borderWidth: 2,
              yAxisID: 'y1',
            },
          ],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { display: true, labels: { color: 'rgba(15,23,42,.72)' } },
            tooltip: {
              callbacks: {
                label: (ctx2) => {
                  const v = ctx2.parsed.y;
                  const isShare = ctx2.dataset.yAxisID === 'y1';
                  const isRate = metricKey === 'ctr' || metricKey === 'read_rate' || metricKey === 'pay_rate';
                  if (isShare) return `${ctx2.dataset.label}: ${(v * 100).toFixed(2)}%`;
                  if (metricKey === 'rev_per_imp') return `${ctx2.dataset.label}: ${Number(v).toFixed(4)}`;
                  if (isRate) return `${ctx2.dataset.label}: ${(v * 100).toFixed(2)}%`;
                  return `${ctx2.dataset.label}: ${v}`;
                },
              },
            },
          },
          scales: {
            x: {
              ticks: { color: 'rgba(15,23,42,.55)', maxRotation: 0, autoSkip: true, maxTicksLimit: 10 },
              grid: { color: 'rgba(15,23,42,.06)' },
            },
            y: {
              ticks: {
                color: 'rgba(15,23,42,.55)',
                callback: (v) => {
                  if (metricKey === 'rev_per_imp') return Number(v).toFixed(4);
                  if (metricKey === 'ctr' || metricKey === 'read_rate' || metricKey === 'pay_rate') return `${Number(v * 100).toFixed(0)}%`;
                  return v;
                },
              },
              grid: { color: 'rgba(15,23,42,.06)' },
            },
            y1: {
              position: 'right',
              ticks: { color: 'rgba(15,23,42,.45)', callback: (v) => `${Number(v * 100).toFixed(0)}%` },
              grid: { drawOnChartArea: false },
            },
          },
        },
      });
      if (hintEl) hintEl.textContent = '图表：蓝线=效率指标，灰线=运营宣推曝光占比';
    } else if (hintEl) {
      hintEl.textContent = '图表库未加载（可能网络受限），仍可查看下方表格。';
    }

    // 效率-规模平衡提示：仅针对最新周（最新周 vs 上周）
    const balanceHintEl = $('opsBalanceHint');
    if (balanceHintEl && weekStarts.length >= 2 && latestDer && prevDer) {
      const avgEffFirst = prevDer[effMetric] ?? 0;
      const avgEffSecond = latestDer[effMetric] ?? 0;
      const avgShareFirst = prevShare ?? 0;
      const avgShareSecond = latestShare ?? 0;
      const effUp = avgEffSecond > avgEffFirst * 1.02;
      const effDown = avgEffSecond < avgEffFirst * 0.98;
      const shareUp = avgShareSecond > avgShareFirst * 1.02;
      const shareDown = avgShareSecond < avgShareFirst * 0.98;
      let balanceText = '';
      if (effUp && shareDown) {
        balanceText = '<strong>效率-规模平衡提示：</strong>效率提升但曝光占比下降，两条曲线在拉大。建议在高效资源位/来源适度加量，避免为冲规模牺牲效率。';
      } else if (effDown && shareUp) {
        balanceText = '<strong>效率-规模平衡提示：</strong>曝光占比上升但效率下降，两条曲线在拉大。建议结构性挪量：从低效资源位向高效资源位调配，加量时设效率底线（如每曝光收入下降不超过5%）。';
      } else if (effUp && shareUp) {
        balanceText = '<strong>效率-规模平衡提示：</strong>效率与规模均向好，保持当前策略。加量时优先高效资源位，维持效率底线。';
      } else if (effDown && shareDown) {
        balanceText = '<strong>效率-规模平衡提示：</strong>效率与曝光占比均下降，需关注。建议先做结构性调配，从低效挪向高效，再考虑加量。';
      } else {
        balanceText = '<strong>效率-规模平衡提示：</strong>效率与规模变化较平稳。加量时优先高效资源位，建议设效率底线（每曝光收入下降不超过5%~10%）。';
      }
      balanceHintEl.innerHTML = balanceText;
    } else if (balanceHintEl) {
      balanceHintEl.innerHTML = '';
    }

    // 运营周报一句话汇总：仅针对最新周输出（最新周 vs 上周）
    const summaryEl = $('opsWeeklySummary');
    if (summaryEl && weekStarts.length >= 1 && opsRows.length > 0) {
      const excludePlacement = new Set(['其它', '其他', '(空)']);
      const opsRowsForSummary = opsRows.filter((r) => !excludePlacement.has(String(r[COLS.placement] ?? '').trim() || '(空)'));

      const prevWeekStart = weekStarts.length >= 2 ? weekStarts[weekStarts.length - 2] : null;
      const latestWeekStart = weekStarts.length >= 1 ? weekStarts[weekStarts.length - 1] : null;
      const firstKeys = prevWeekStart ? new Set([prevWeekStart]) : new Set();
      const secondKeys = latestWeekStart ? new Set([latestWeekStart]) : new Set();

      const avgEffFirst = prevDer ? prevDer[effMetric] ?? 0 : 0;
      const avgEffSecond = latestDer ? latestDer[effMetric] ?? 0 : 0;
      const avgShareFirst = prevShare ?? 0;
      const avgShareSecond = latestShare ?? 0;
      const effUp = avgEffSecond > avgEffFirst * 1.05;
      const effDown = avgEffSecond < avgEffFirst * 0.95;
      const shareUp = avgShareSecond > avgShareFirst * 1.05;
      const shareDown = avgShareSecond < avgShareFirst * 0.95;

      const getBucketKey = (d) => toWeekStartLocal(d);
      const allFirstRows = opsRowsForSummary.filter((r) => firstKeys.has(getBucketKey(String(r[COLS.date] ?? '').trim())));
      const allSecondRows = opsRowsForSummary.filter((r) => secondKeys.has(getBucketKey(String(r[COLS.date] ?? '').trim())));
      const totalFirstImp = allFirstRows.reduce((s, r) => s + num(r[COLS.ops_imp]), 0);
      const totalSecondImp = allSecondRows.reduce((s, r) => s + num(r[COLS.ops_imp]), 0);
      const opsCols = pickScopeCols('ops');
      const bySource = groupBy(opsRowsForSummary, (r) => String(r[COLS.source] ?? '').trim() || '(空)');
      const sourceChanges = [];
      for (const [src, rows] of bySource) {
        const firstRows = rows.filter((r) => firstKeys.has(getBucketKey(String(r[COLS.date] ?? '').trim())));
        const secondRows = rows.filter((r) => secondKeys.has(getBucketKey(String(r[COLS.date] ?? '').trim())));
        if (!prevWeekStart || firstRows.length < 1 || secondRows.length < 1) continue;
        const a1 = computeFunnelAgg(firstRows, opsCols);
        const a2 = computeFunnelAgg(secondRows, opsCols);
        const d1 = computeDerived(a1);
        const d2 = computeDerived(a2);
        const effChange = safeDiv(d2[effMetric] - d1[effMetric], d1[effMetric] || 1);
        const share1 = safeDiv(a1.imp, totalFirstImp);
        const share2 = safeDiv(a2.imp, totalSecondImp);
        const shareChange = share1 > 0 ? safeDiv(share2 - share1, share1) : 0;
        const totalImp = a1.imp + a2.imp;
        sourceChanges.push({ src, effChange, shareChange, eff1: d1[effMetric], eff2: d2[effMetric], totalImp });
      }
      sourceChanges.sort((a, b) => b.totalImp - a.totalImp);

      const parts = [];
      let bigTrend = '';
      if (prevWeekStart && latestWeekStart) {
        if (effUp && shareUp) bigTrend = '最新周大盘效率与曝光占比均向好';
        else if (effUp && shareDown) bigTrend = '最新周大盘效率提升但曝光占比下降';
        else if (effDown && shareUp) bigTrend = '最新周大盘曝光占比上升但效率下降';
        else if (effDown && shareDown) bigTrend = '最新周大盘效率与曝光占比均下降';
        else bigTrend = '最新周大盘效率与曝光占比变化较平稳';
      } else {
        bigTrend = '最新周大盘数据已展示（当前仅一周，暂无环比）';
      }
      parts.push(bigTrend);

      const highlights = [];
      const anomalies = [];
      for (const x of sourceChanges) {
        if (x.effChange >= 0.08 && x.totalImp > 50000) highlights.push(`${x.src}效率提升明显`);
        else if (x.effChange <= -0.08 && x.totalImp > 50000) anomalies.push(`${x.src}效率下降明显`);
        else if (x.shareChange >= 0.15 && x.totalImp > 50000) highlights.push(`${x.src}曝光占比提升`);
        else if (x.shareChange <= -0.15 && x.totalImp > 50000) anomalies.push(`${x.src}曝光占比下降需关注`);
      }
      if (highlights.length) parts.push(`其中${highlights.slice(0, 2).join('、')}`);
      if (anomalies.length) parts.push(`${anomalies.slice(0, 2).join('、')}`);

      summaryEl.innerHTML = `<strong>【结论】</strong>${parts.join('；')}。（仅针对最新周）`;
    } else if (summaryEl) {
      summaryEl.innerHTML = '';
    }

    const dataRows = series.slice(-60).reverse();
    const wowWrap = (r, i, data, getRaw, format) => {
      const raw = format(getRaw(r));
      if (i !== 0 || !data?.[1]) return raw;
      const cv = getRaw(r);
      const pv = getRaw(data[1]);
      if (pv == null || (typeof pv === 'number' && !Number.isFinite(pv))) return raw;
      if (cv > pv) return `<span class="bad">${raw}</span>`;
      if (cv < pv) return `<span class="good">${raw}</span>`;
      return raw;
    };
    const cols = [
      { label: timeGran === 'day' ? '日期' : '周起始(周一)', value: (r) => `<span class="num">${r.date}</span>` },
      ...(timeGran === 'day' ? [] : [{ label: '覆盖天数', className: 'num', value: (r) => `<span class="num">${r.days ?? ''}</span>` }]),
      { label: '曝光占比', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.opsShare, (v) => fmtRate(v, 2)) },
      { label: '运营宣推曝光', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.imp, (v) => fmtInt(v)) },
      { label: '运营宣推点击', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.clk, (v) => fmtInt(v)) },
      { label: '运营宣推阅读', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.read, (v) => fmtInt(v)) },
      { label: '运营宣推付费用户', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.payu, (v) => fmtInt(v)) },
      { label: '运营宣推收入', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.rev, (v) => fmtMoney(v, 2)) },
      { label: '每曝光收入', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.rev_per_imp, (v) => fmtMoney(v, 4)) },
      { label: 'CTR', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.ctr, (v) => fmtRate(v, 2)) },
      { label: '阅读率', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.read_rate, (v) => fmtRate(v, 2)) },
      { label: '付费率', className: 'num', value: (r, i, data) => wowWrap(r, i, data, (x) => x?.pay_rate, (v) => fmtRate(v, 2)) },
    ];
    buildTable($('opsDailyTable'), cols, dataRows);
  }

  // 2) 各业务宣发效率（趋势） – 固定运营宣推口径
  {
    const timeGran = $('srcTimeGran')?.value || 'day';
    const srcRows = filterRowsByDate(dailyRows, $('srcStartDate')?.value, $('srcEndDate')?.value);

    const opsScopeCols = pickScopeCols('ops');
    const bySrc = groupBy(srcRows, (r) => String(r[COLS.source] ?? '').trim() || '(空)');
    const data = Array.from(bySrc.entries()).map(([src, rows]) => {
      const a = computeFunnelAgg(rows, opsScopeCols);
      const d = computeDerived(a);
      return {
        src,
        ...a,
        ...d,
        erpi_ops: d.rev_per_imp,
      };
    });

    const totalOpsImp = data.reduce((s, r) => s + r.imp, 0);
    data.forEach((r) => {
      r.ops_share = safeDiv(r.imp, totalOpsImp);
    });

    data.sort((a, b) => (b.erpi_ops - a.erpi_ops));

    const pillsEl = $('sourcePills');
    if (pillsEl && pillsEl.childElementCount === 0) {
      const sourcesByImp = [...data].sort((a, b) => b.imp - a.imp).map((x) => x.src);
      const defaultPick = sourcesByImp.slice(0, 6);
      state.selectedSources = new Set(defaultPick);

      pillsEl.innerHTML = '';
      for (const src of sourcesByImp) {
        const label = document.createElement('label');
        label.className = `pill ${state.selectedSources.has(src) ? 'pill--on' : ''}`;
        label.innerHTML = `<input type="checkbox" ${state.selectedSources.has(src) ? 'checked' : ''} />${src}`;
        label.addEventListener('click', (e) => {
          e.preventDefault();
          if (state.selectedSources.has(src)) state.selectedSources.delete(src);
          else state.selectedSources.add(src);
          label.classList.toggle('pill--on');
          render();
        });
        pillsEl.appendChild(label);
      }
    }

    const chartEl = $('sourceTrendChart');
    const hintEl = $('sourceChartHint');
    const palette = ['#4f7cff', '#40c79a', '#ff9f43', '#a855f7', '#ef4444', '#06b6d4', '#84cc16', '#f97316'];

    const byDate = groupBy(srcRows, (r) => String(r[COLS.date] ?? '').trim());
    const dates = Array.from(byDate.keys()).filter(Boolean).sort();
    const dailyBuckets = dates.map((d) => ({ date: d, rows: byDate.get(d) }));

    const aggregateBuckets = (buckets, gran) => {
      if (gran === 'day') return buckets.map((b) => ({ key: b.date, rows: b.rows, days: 1 }));
      const byWeek = new Map();
      for (const b of buckets) {
        const wk = toWeekStartLocal(b.date);
        if (!byWeek.has(wk)) byWeek.set(wk, []);
        byWeek.get(wk).push(b);
      }
      const weeks = Array.from(byWeek.keys()).sort();
      return weeks.map((wk) => {
        const arr = byWeek.get(wk);
        return { key: wk, days: arr.length, rows: arr.flatMap((x) => x.rows) };
      });
    };

    const buckets = aggregateBuckets(dailyBuckets, timeGran);
    const labels = buckets.map((b) => b.key);

    const selected = Array.from(state.selectedSources || []);
    if (typeof Chart !== 'undefined' && chartEl) {
      const ctx = chartEl.getContext('2d');
      if (state.charts.sourceTrend) state.charts.sourceTrend.destroy();

      const datasets = [];
      selected.forEach((src, i) => {
        const color = palette[i % palette.length];

        const effVals = buckets.map((b) => {
          const r = b.rows.filter((x) => (String(x[COLS.source] ?? '').trim() || '(空)') === src);
          const a = computeFunnelAgg(r, opsScopeCols);
          const d = computeDerived(a);
          return d[effMetric] ?? 0;
        });
        datasets.push({
          label: `${src} · ${effLabel(effMetric)}`,
          data: effVals,
          borderColor: color,
          backgroundColor: 'transparent',
          tension: 0.25,
          pointRadius: 0,
          borderWidth: 2,
          yAxisID: 'y',
        });

        const ctrVals = buckets.map((b) => {
          const r = b.rows.filter((x) => (String(x[COLS.source] ?? '').trim() || '(空)') === src);
          const a = computeFunnelAgg(r, opsScopeCols);
          return safeDiv(a.clk, a.imp);
        });
        datasets.push({
          label: `${src} · p-CTR`,
          data: ctrVals,
          borderColor: color,
          borderDash: [4, 4],
          backgroundColor: 'transparent',
          tension: 0.25,
          pointRadius: 0,
          borderWidth: 1.5,
          yAxisID: 'y1',
        });
      });

      state.charts.sourceTrend = new Chart(ctx, {
        type: 'line',
        data: { labels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { display: true, labels: { color: 'rgba(15,23,42,.72)' } },
            tooltip: {
              callbacks: {
                label: (ctx2) => {
                  const v = ctx2.parsed.y;
                  if (ctx2.dataset.yAxisID === 'y1') {
                    return `${ctx2.dataset.label}: ${(v * 100).toFixed(2)}%`;
                  }
                  if (effMetric === 'rev_per_imp') return `${ctx2.dataset.label}: ${Number(v).toFixed(4)}`;
                  if (effMetric === 'ctr' || effMetric === 'read_rate' || effMetric === 'pay_rate') return `${ctx2.dataset.label}: ${(v * 100).toFixed(2)}%`;
                  return `${ctx2.dataset.label}: ${v}`;
                },
              },
            },
          },
          scales: {
            x: {
              ticks: { color: 'rgba(15,23,42,.55)', maxRotation: 0, autoSkip: true, maxTicksLimit: 10 },
              grid: { color: 'rgba(15,23,42,.06)' },
            },
            y: {
              ticks: {
                color: 'rgba(15,23,42,.55)',
                callback: (v) => {
                  if (effMetric === 'rev_per_imp') return Number(v).toFixed(4);
                  if (effMetric === 'ctr' || effMetric === 'read_rate' || effMetric === 'pay_rate') return `${Number(v * 100).toFixed(0)}%`;
                  return v;
                },
              },
              grid: { color: 'rgba(15,23,42,.06)' },
            },
            y1: {
              position: 'right',
              ticks: {
                color: 'rgba(15,23,42,.45)',
                callback: (v) => `${Number(v * 100).toFixed(0)}%`,
              },
              grid: { drawOnChartArea: false },
            },
          },
        },
      });
      if (hintEl) {
        const granLabel = timeGran === 'day' ? '日' : (timeGran === 'week_avg' ? '周（日均）' : '周（汇总）');
        hintEl.textContent = `趋势：${granLabel} · 口径=运营宣推 · 每来源两条线：实线=${effLabel(effMetric)}，虚线=p-CTR · 已选来源=${selected.length}`;
      }
    } else if (hintEl) {
      hintEl.textContent = '图表库未加载（可能网络受限），仍可查看下方表格。';
    }

    const cols = [
      { label: '宣发来源', value: (r) => r.src },
      { label: '运营宣推曝光', className: 'num', value: (r) => fmtInt(r.imp) },
      { label: '运营宣推曝光占比', className: 'num', value: (r) => fmtRate(r.ops_share, 2) },
      { label: '运营宣推收入', className: 'num', value: (r) => fmtMoney(r.rev, 2) },
      { label: '每曝光收入(运营宣推)', className: 'num', value: (r) => fmtMoney(r.erpi_ops, 4) },
      { label: 'p-CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
      { label: '阅读率', className: 'num', value: (r) => fmtRate(r.read_rate, 2) },
      { label: '付费率', className: 'num', value: (r) => fmtRate(r.pay_rate, 2) },
    ];
    buildTable($('bySourceTable'), cols, data.slice(0, 20));
  }

  // 3) 各资源位效率趋势
  {
    const timeGran = $('posTimeGran')?.value || 'day';
    const posRowsData = filterRowsByDate(dailyRows, $('posStartDate')?.value, $('posEndDate')?.value);

    const byPos = groupBy(posRowsData, (r) => String(r[COLS.placement] ?? '').trim() || '(空)');
    const data = Array.from(byPos.entries()).map(([pos, rows]) => {
      const a = computeFunnelAgg(rows, scopeCols);
      const d = computeDerived(a);
      return { pos, ...a, ...d };
    }).sort((a, b) => (b[effMetric] - a[effMetric]));

    const pillsEl = $('placementPills');
    if (pillsEl && pillsEl.childElementCount === 0) {
      const placementsByImp = [...data].sort((a, b) => b.imp - a.imp).map((x) => x.pos);
      const defaultPick = placementsByImp.slice(0, 6);
      state.selectedPlacements = new Set(defaultPick);

      pillsEl.innerHTML = '';
      for (const pos of placementsByImp) {
        const label = document.createElement('label');
        label.className = `pill ${state.selectedPlacements.has(pos) ? 'pill--on' : ''}`;
        label.innerHTML = `<input type="checkbox" ${state.selectedPlacements.has(pos) ? 'checked' : ''} />${pos}`;
        label.addEventListener('click', (e) => {
          e.preventDefault();
          if (state.selectedPlacements.has(pos)) state.selectedPlacements.delete(pos);
          else state.selectedPlacements.add(pos);
          label.classList.toggle('pill--on');
          render();
        });
        pillsEl.appendChild(label);
      }
    }

    const chartEl = $('placementTrendChart');
    const hintEl = $('placementChartHint');
    const palette = ['#4f7cff', '#40c79a', '#ff9f43', '#a855f7', '#ef4444', '#06b6d4', '#84cc16', '#f97316'];

    const byDate = groupBy(posRowsData, (r) => String(r[COLS.date] ?? '').trim());
    const dates = Array.from(byDate.keys()).filter(Boolean).sort();
    const dailyBuckets = dates.map((d) => ({ date: d, rows: byDate.get(d) }));

    const aggregateBuckets = (buckets, gran) => {
      if (gran === 'day') return buckets.map((b) => ({ key: b.date, rows: b.rows, days: 1 }));
      const byWeek = new Map();
      for (const b of buckets) {
        const wk = toWeekStartLocal(b.date);
        if (!byWeek.has(wk)) byWeek.set(wk, []);
        byWeek.get(wk).push(b);
      }
      const weeks = Array.from(byWeek.keys()).sort();
      return weeks.map((wk) => {
        const arr = byWeek.get(wk);
        return { key: wk, days: arr.length, rows: arr.flatMap((x) => x.rows) };
      });
    };

    const buckets = aggregateBuckets(dailyBuckets, timeGran);
    const labels = buckets.map((b) => b.key);
    const selected = Array.from(state.selectedPlacements || []);

    if (typeof Chart !== 'undefined' && chartEl) {
      const ctx = chartEl.getContext('2d');
      if (state.charts.placementTrend) state.charts.placementTrend.destroy();

      const datasets = selected.map((pos, i) => {
        const vals = buckets.map((b) => {
          const r = b.rows.filter((x) => (String(x[COLS.placement] ?? '').trim() || '(空)') === pos);
          const a = computeFunnelAgg(r, scopeCols);
          const d = computeDerived(a);
          return d[effMetric] ?? 0;
        });
        return {
          label: pos,
          data: vals,
          borderColor: palette[i % palette.length],
          backgroundColor: 'transparent',
          tension: 0.25,
          pointRadius: 0,
          borderWidth: 2,
        };
      });

      state.charts.placementTrend = new Chart(ctx, {
        type: 'line',
        data: { labels, datasets },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          plugins: {
            legend: { display: true, labels: { color: 'rgba(15,23,42,.72)' } },
            tooltip: {
              callbacks: {
                label: (ctx2) => {
                  const v = ctx2.parsed.y;
                  if (effMetric === 'rev_per_imp') return `${ctx2.dataset.label}: ${Number(v).toFixed(4)}`;
                  if (effMetric === 'ctr' || effMetric === 'read_rate' || effMetric === 'pay_rate') return `${ctx2.dataset.label}: ${(v * 100).toFixed(2)}%`;
                  return `${ctx2.dataset.label}: ${v}`;
                },
              },
            },
          },
          scales: {
            x: {
              ticks: { color: 'rgba(15,23,42,.55)', maxRotation: 0, autoSkip: true, maxTicksLimit: 10 },
              grid: { color: 'rgba(15,23,42,.06)' },
            },
            y: {
              ticks: {
                color: 'rgba(15,23,42,.55)',
                callback: (v) => {
                  if (effMetric === 'rev_per_imp') return Number(v).toFixed(4);
                  if (effMetric === 'ctr' || effMetric === 'read_rate' || effMetric === 'pay_rate') return `${Number(v * 100).toFixed(0)}%`;
                  return v;
                },
              },
              grid: { color: 'rgba(15,23,42,.06)' },
            },
          },
        },
      });
      if (hintEl) {
        const granLabel = timeGran === 'day' ? '日' : (timeGran === 'week_avg' ? '周（日均）' : '周（汇总）');
        hintEl.textContent = `趋势：${granLabel} · 指标=${effLabel(effMetric)} · 已选资源位=${selected.length}`;
      }
    } else if (hintEl) {
      hintEl.textContent = '图表库未加载（可能网络受限），仍可查看下方表格。';
    }

    const cols = [
      { label: '资源位', value: (r) => r.pos },
      { label: '曝光', className: 'num', value: (r) => fmtInt(r.imp) },
      { label: '收入', className: 'num', value: (r) => fmtMoney(r.rev, 2) },
      { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.rev_per_imp, 4) },
      { label: 'CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
      { label: '阅读率', className: 'num', value: (r) => fmtRate(r.read_rate, 2) },
      { label: '付费率', className: 'num', value: (r) => fmtRate(r.pay_rate, 2) },
    ];
    buildTable($('byPlacementTable'), cols, data);
  }

  // 4) 每天/每周综合效率最高的项目明细 Top 10
  {
    const topTimeMode = $('topTimeMode')?.value || 'day';
    state.topProjectsByWeek = topTimeMode === 'week';
    const topRowsData = filterRowsByDate(dailyRows, $('topStartDate')?.value, $('topEndDate')?.value);

    const topRows = [];
    const dateColLabel = state.topProjectsByWeek ? '周起始' : '日期';

    if (state.topProjectsByWeek) {
      const byWeek = new Map();
      for (const r of topRowsData) {
        const d = String(r[COLS.date] ?? '').trim();
        if (!d) continue;
        const wk = toWeekStartLocal(d);
        if (!byWeek.has(wk)) byWeek.set(wk, []);
        byWeek.get(wk).push(r);
      }
      const weeks = Array.from(byWeek.keys()).sort();
      for (const weekStart of weeks) {
        const weekRows = byWeek.get(weekStart);
        const byTopic = groupBy(weekRows, (r) => String(r[COLS.topicId] ?? '').trim() || String(r[COLS.topicName] ?? '').trim() || '(未知)');
        const topics = Array.from(byTopic.entries()).map(([tid, rows]) => {
          const a = computeFunnelAgg(rows, scopeCols);
          const d = computeDerived(a);
          const name = String(rows[0]?.[COLS.topicName] ?? '').trim();
          const src = String(rows[0]?.[COLS.source] ?? '').trim();
          const pos = String(rows[0]?.[COLS.placement] ?? '').trim() || '(空)';
          return { date: weekStart, tid, name, src, pos, ...a, ...d };
        }).filter((t) => {
          if (t.imp < minImp) return false;
          if (state.topPlacementFilter && state.topPlacementFilter !== '(全部)' && t.pos !== state.topPlacementFilter) return false;
          if (state.topSourceFilter && state.topSourceFilter !== '(全部)' && t.src !== state.topSourceFilter) return false;
          return true;
        });
        topics.sort((a, b) => (b[effMetric] - a[effMetric]));
        const displayTopics = state.topProjectsShowAll ? topics : topics.slice(0, 10);
        for (let i = 0; i < displayTopics.length; i += 1) {
          topRows.push({ rank: i + 1, ...displayTopics[i] });
        }
      }
    } else {
      const byDate = groupBy(topRowsData, (r) => String(r[COLS.date] ?? '').trim());
      const dates = Array.from(byDate.keys()).filter(Boolean).sort();
      for (const date of dates) {
        const day = byDate.get(date);
        const byTopic = groupBy(day, (r) => String(r[COLS.topicId] ?? '').trim() || String(r[COLS.topicName] ?? '').trim() || '(未知)');
        const topics = Array.from(byTopic.entries()).map(([tid, rows]) => {
          const a = computeFunnelAgg(rows, scopeCols);
          const d = computeDerived(a);
          const name = String(rows[0]?.[COLS.topicName] ?? '').trim();
          const src = String(rows[0]?.[COLS.source] ?? '').trim();
          const pos = String(rows[0]?.[COLS.placement] ?? '').trim() || '(空)';
          return { date, tid, name, src, pos, ...a, ...d };
        }).filter((t) => {
          if (t.imp < minImp) return false;
          if (state.topPlacementFilter && state.topPlacementFilter !== '(全部)' && t.pos !== state.topPlacementFilter) return false;
          if (state.topSourceFilter && state.topSourceFilter !== '(全部)' && t.src !== state.topSourceFilter) return false;
          return true;
        });
        topics.sort((a, b) => (b[effMetric] - a[effMetric]));
        const displayTopics = state.topProjectsShowAll ? topics : topics.slice(0, 10);
        for (let i = 0; i < displayTopics.length; i += 1) {
          topRows.push({ rank: i + 1, ...displayTopics[i] });
        }
      }
    }

    const pillsEl = $('topPlacementPills');
    const allPlacements = Array.from(new Set(topRowsData.map((r) => String(r[COLS.placement] ?? '').trim() || '(空)'))).sort();
    if (pillsEl && pillsEl.childElementCount === 0) {
      state.topPlacementFilter = '(全部)';
      const mkPill = (label, filterKey, pillsContainer) => {
        const el = document.createElement('label');
        el.className = `pill ${state[filterKey] === label ? 'pill--on' : ''}`;
        el.textContent = label;
        el.addEventListener('click', (e) => {
          e.preventDefault();
          state[filterKey] = label;
          Array.from(pillsContainer.children).forEach((c) => c.classList.remove('pill--on'));
          el.classList.add('pill--on');
          render();
        });
        return el;
      };
      pillsEl.innerHTML = '';
      pillsEl.appendChild(mkPill('(全部)', 'topPlacementFilter', pillsEl));
      allPlacements.forEach((p) => pillsEl.appendChild(mkPill(p, 'topPlacementFilter', pillsEl)));
    }

    // 宣发来源筛选 pill
    const srcPillsEl = $('topSourcePills');
    const allSources = Array.from(new Set(topRowsData.map((r) => String(r[COLS.source] ?? '').trim() || '(空)'))).sort();
    if (srcPillsEl && srcPillsEl.childElementCount === 0) {
      state.topSourceFilter = '(全部)';
      const mkPill = (label, filterKey, pillsContainer) => {
        const el = document.createElement('label');
        el.className = `pill ${state[filterKey] === label ? 'pill--on' : ''}`;
        el.textContent = label;
        el.addEventListener('click', (e) => {
          e.preventDefault();
          state[filterKey] = label;
          Array.from(pillsContainer.children).forEach((c) => c.classList.remove('pill--on'));
          el.classList.add('pill--on');
          render();
        });
        return el;
      };
      srcPillsEl.innerHTML = '';
      srcPillsEl.appendChild(mkPill('(全部)', 'topSourceFilter', srcPillsEl));
      allSources.forEach((s) => srcPillsEl.appendChild(mkPill(s, 'topSourceFilter', srcPillsEl)));
    }

    const topProjectsToggleBtn = $('topProjectsToggleBtn');
    if (topProjectsToggleBtn) {
      topProjectsToggleBtn.textContent = state.topProjectsShowAll ? '返回 Top 10' : '查看全部项目';
    }

    const cols = [
      { label: dateColLabel, value: (r) => `<span class="num">${r.date}</span>` },
      { label: 'Top', className: 'num', value: (r) => `<span class="num">#${r.rank}</span>` },
      { label: '专题ID', className: 'num', value: (r) => `<span class="num">${r.tid}</span>` },
      { label: '专题名称', value: (r) => r.name || '<span class="muted">（空）</span>' },
      { label: '宣发来源', value: (r) => r.src || '<span class="muted">（空）</span>' },
      { label: '资源位', value: (r) => r.pos || '<span class="muted">（空）</span>' },
      { label: '曝光', className: 'num', value: (r) => fmtInt(r.imp) },
      { label: '收入', className: 'num', value: (r) => fmtMoney(r.rev, 2) },
      { label: effLabel(effMetric), className: 'num', value: (r) => {
        if (effMetric === 'rev_per_imp') return fmtMoney(r.rev_per_imp, 4);
        if (effMetric === 'ctr' || effMetric === 'read_rate' || effMetric === 'pay_rate') return fmtRate(r[effMetric], 2);
        return fmtMoney(r[effMetric], 4);
      } },
      { label: 'CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
      { label: '阅读率', className: 'num', value: (r) => fmtRate(r.read_rate, 2) },
      { label: '付费率', className: 'num', value: (r) => fmtRate(r.pay_rate, 2) },
    ];
    buildTable($('dailyTopProjectsTable'), cols, topRows);
  }

  // 5) 资源位 × 宣发来源效率矩阵（热力）
  {
    const matrixRows = filterRowsByDate(dailyRows, $('matrixStartDate')?.value, $('matrixEndDate')?.value);
    const topPosN = Number($('matrixTopPos')?.value || 8);
    const topSrcN = Number($('matrixTopSrc')?.value || 8);

    const byPos = groupBy(matrixRows, (r) => String(r[COLS.placement] ?? '').trim() || '(空)');
    const bySrc = groupBy(matrixRows, (r) => String(r[COLS.source] ?? '').trim() || '(空)');

    const posRank = Array.from(byPos.entries()).map(([pos, rows]) => ({
      pos,
      imp: computeFunnelAgg(rows, scopeCols).imp,
    })).sort((a, b) => b.imp - a.imp).slice(0, topPosN).map((x) => x.pos);

    const srcRank = Array.from(bySrc.entries()).map(([src, rows]) => ({
      src,
      imp: computeFunnelAgg(rows, scopeCols).imp,
    })).sort((a, b) => b.imp - a.imp).slice(0, topSrcN).map((x) => x.src);

    const cell = new Map();
    posRank.forEach((pos) => {
      srcRank.forEach((src) => {
        const rows = matrixRows.filter((r) =>
          (String(r[COLS.placement] ?? '').trim() || '(空)') === pos &&
          (String(r[COLS.source] ?? '').trim() || '(空)') === src
        );
        const a = computeFunnelAgg(rows, scopeCols);
        const d = computeDerived(a);
        cell.set(`${pos}\u0001${src}`, { ...a, ...d });
      });
    });

    const vals = [];
    cell.forEach((v) => vals.push(v[effMetric] || 0));
    const minV = vals.length ? Math.min(...vals) : 0;
    const maxV = vals.length ? Math.max(...vals) : 0;
    const norm = (x) => {
      if (maxV === minV) return 0;
      return (x - minV) / (maxV - minV);
    };

    const metricFormatter = (v) => {
      if (effMetric === 'rev_per_imp') return fmtMoney(v, 4);
      if (effMetric === 'ctr' || effMetric === 'read_rate' || effMetric === 'pay_rate') return fmtRate(v, 2);
      return String(v);
    };

    const head = ['<th>资源位 \\ 来源</th>', ...srcRank.map((s) => `<th>${s}</th>`)].join('');
    const body = posRank.map((pos) => {
      const tds = srcRank.map((src) => {
        const v = cell.get(`${pos}\u0001${src}`) || { imp: 0, rev: 0 };
        const mv = v[effMetric] || 0;
        const alpha = 0.08 + 0.42 * norm(mv);
        return `<td style="background: rgba(79,124,255,${alpha.toFixed(3)});">
          <div class="cellBar">
            <span class="cellVal">${metricFormatter(mv)}</span>
            <span class="cellMeta">曝光 ${fmtInt(v.imp || 0)}</span>
          </div>
        </td>`;
      }).join('');
      return `<tr><td><strong>${pos}</strong></td>${tds}</tr>`;
    }).join('');

    const el = $('matrixTable');
    el.innerHTML = `<table class="matrixTable"><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;

    const hint = $('matrixHint');
    if (hint) {
      hint.textContent = `口径=${scope === 'ops' ? '运营宣推' : '全局'}；指标=${effLabel(effMetric)}；颜色越深效率越高。`;
    }
  }
}

async function loadRows(rows, fileNameLabel = null) {
  state.rows = mergeRowsDedup(rows);
  state.lastFileName = fileNameLabel;
  const srcPills = $('sourcePills');
  if (srcPills) srcPills.innerHTML = '';
  const posPills = $('placementPills');
  if (posPills) posPills.innerHTML = '';
  const topPosPills = $('topPlacementPills');
  if (topPosPills) topPosPills.innerHTML = '';
  const topSrcPills = $('topSourcePills');
  if (topSrcPills) topSrcPills.innerHTML = '';
  [
    'latestStartDate', 'latestEndDate',
    'opsStartDate', 'opsEndDate',
    'srcStartDate', 'srcEndDate',
    'posStartDate', 'posEndDate',
    'topStartDate', 'topEndDate',
    'matrixStartDate', 'matrixEndDate',
  ].forEach((id) => {
    const el = $(id);
    if (el) el.value = '';
  });
  render();
}

async function parseCsvText(text, fileName = null) {
  try {
    $('statusHint').textContent = '正在解析 CSV...';
    const rows = await parseCsvRows(text);
    await loadRows(rows, fileName);
  } catch (err) {
    $('statusHint').textContent = `解析失败：${err?.message || err}`;
  }
}

async function resolveOpsSubdirHandle(rootHandle) {
  for (const name of OPS_SUBDIR_CANDIDATES) {
    try {
      const h = await rootHandle.getDirectoryHandle(name, { create: false });
      return { handle: h, name };
    } catch (_) {}
  }
  const list = OPS_SUBDIR_CANDIDATES.map((x) => `「${x}」`).join(' 或 ');
  throw new Error(`未找到运营宣推子文件夹（需要在绑定目录下创建 ${list} 文件夹）。`);
}

async function loadAllCsvFromBoundFolder() {
  if (!state.boundDirHandle) {
    $('statusHint').textContent = '请先点击“绑定数据文件夹”。';
    return;
  }

  try {
    const perm = await state.boundDirHandle.requestPermission({ mode: 'read' });
    if (perm !== 'granted') {
      $('statusHint').textContent = '未获得文件夹读取权限，无法更新。';
      return;
    }
    const { handle: opsDir, name: subName } = await resolveOpsSubdirHandle(state.boundDirHandle);
    $('statusHint').textContent = `正在扫描并合并「${subName}」文件夹中的 CSV...`;

    const allRows = [];
    let fileCount = 0;
    for await (const entry of opsDir.values()) {
      if (entry.kind !== 'file') continue;
      if (!entry.name.toLowerCase().endsWith('.csv')) continue;
      const file = await entry.getFile();
      const text = await file.text();
      const rows = await parseCsvRows(text);
      // Avoid call stack overflow on huge CSV (don't use push(...rows))
      for (const r of rows) allRows.push(r);
      fileCount += 1;
    }

    if (!fileCount) {
      $('statusHint').textContent = '运营宣推子文件夹下未找到 CSV 文件。';
      return;
    }
    await loadRows(allRows, `文件夹模式/${subName}（${fileCount}个CSV）`);
    $('statusHint').textContent = `更新完成：已合并「${subName}」下 ${fileCount} 个 CSV，去重后 ${state.rows.length.toLocaleString('zh-CN')} 行。`;
  } catch (err) {
    $('statusHint').textContent = `文件夹更新失败：${err?.message || err}`;
  }
}

async function restoreBoundFolderHandle() {
  if (!window.showDirectoryPicker) return;
  try {
    const savedHandle = await dbGet(DB_KEY_DIR_HANDLE);
    if (!savedHandle) return;
    const perm = await savedHandle.queryPermission({ mode: 'read' });
    if (perm === 'granted') {
      state.boundDirHandle = savedHandle;
      await loadAllCsvFromBoundFolder();
    } else {
      state.boundDirHandle = savedHandle;
      const hint = $('statusHint');
      if (hint) hint.textContent = `已恢复绑定文件夹：${savedHandle.name}。请点“更新最新数据”或重新授权。`;
    }
  } catch (err) {
    // Ignore restore failures; user can re-bind manually.
  }
}

async function copyTableAsImage(btn) {
  const wrap = btn?.closest('.tableWrap');
  const table = wrap?.querySelector('table');
  if (!table || typeof html2canvas !== 'function') return;
  try {
    const canvas = await html2canvas(table, {
      useCORS: true,
      scale: 2,
      backgroundColor: '#ffffff',
      logging: false,
    });
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (blob && navigator.clipboard?.write) {
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
      const orig = btn.textContent;
      btn.textContent = '已复制';
      setTimeout(() => { btn.textContent = orig; }, 1500);
    } else {
      const url = canvas.toDataURL('image/png');
      const a = document.createElement('a');
      a.href = url;
      a.download = 'table.png';
      a.click();
      btn.textContent = '已下载';
      setTimeout(() => { btn.textContent = '复制为图片'; }, 1500);
    }
  } catch (err) {
    btn.textContent = '复制失败';
    setTimeout(() => { btn.textContent = '复制为图片'; }, 1500);
  }
}

async function copyChartAsImage(btn) {
  const wrap = btn?.closest('.chartWrap');
  const canvas = wrap?.querySelector('canvas');
  if (!canvas) return;
  try {
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
    if (blob && navigator.clipboard?.write) {
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
      const orig = btn.textContent;
      btn.textContent = '已复制';
      setTimeout(() => { btn.textContent = orig; }, 1500);
    } else {
      const url = canvas.toDataURL('image/png');
      const a = document.createElement('a');
      a.href = url;
      a.download = 'chart.png';
      a.click();
      btn.textContent = '已下载';
      setTimeout(() => { btn.textContent = '复制为图片'; }, 1500);
    }
  } catch (err) {
    const url = canvas.toDataURL('image/png');
    const a = document.createElement('a');
    a.href = url;
    a.download = 'chart.png';
    a.click();
    btn.textContent = '已下载';
    setTimeout(() => { btn.textContent = '复制为图片'; }, 1500);
  }
}

function setup() {
  window.addEventListener('error', (e) => {
    const hint = $('statusHint');
    if (hint) hint.textContent = `页面脚本异常：${e.message || 'unknown error'}`;
  });

  document.addEventListener('click', (e) => {
    if (e.target?.classList?.contains('copyChartBtn')) {
      copyChartAsImage(e.target);
    }
    if (e.target?.classList?.contains('copyTableBtn')) {
      copyTableAsImage(e.target);
    }
  });

  const fileInput = $('fileInput');
  const bindFolderBtn = $('bindFolderBtn');
  const updateFolderBtn = $('updateFolderBtn');

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const text = await file.text();
    await parseCsvText(text, file.name);
  });

  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', async (e) => {
    e.preventDefault();
    const file = e.dataTransfer?.files?.[0];
    if (!file) return;
    const text = await file.text();
    await parseCsvText(text, file.name);
  });

  bindFolderBtn?.addEventListener('click', async () => {
    if (!window.showDirectoryPicker) {
      $('statusHint').textContent = '当前浏览器不支持“绑定文件夹”。请使用 Chrome/Edge 最新版，或继续手动选择CSV。';
      return;
    }
    try {
      const dir = await window.showDirectoryPicker({ mode: 'read' });
      state.boundDirHandle = dir;
      try {
        await dbSet(DB_KEY_DIR_HANDLE, dir);
      } catch (_) {
        // Ignore persistence failure, manual bind still works.
      }
      await loadAllCsvFromBoundFolder();
    } catch (err) {
      if (err?.name === 'AbortError') return;
      $('statusHint').textContent = `绑定文件夹失败：${err?.message || err}`;
    }
  });

  updateFolderBtn?.addEventListener('click', async () => {
    await loadAllCsvFromBoundFolder();
  });

  $('topProjectsToggleBtn')?.addEventListener('click', () => {
    state.topProjectsShowAll = !state.topProjectsShowAll;
    render();
  });

  $('topTimeMode')?.addEventListener('change', () => render());

  $('syncMatrixToLatestWeek')?.addEventListener('click', () => {
    if (state.latestWeekRange) {
      const s = $('matrixStartDate');
      const e = $('matrixEndDate');
      if (s) s.value = state.latestWeekRange.start;
      if (e) e.value = state.latestWeekRange.end;
      render();
    }
  });

  $('minImp').addEventListener('input', () => {
    $('minImpLabel').textContent = String($('minImp').value);
  });
  $('minImp').addEventListener('change', () => render());
  $('metricScope').addEventListener('change', () => render());
  $('effMetric').addEventListener('change', () => render());
  $('opsTimeGran')?.addEventListener('change', () => render());
  $('srcTimeGran')?.addEventListener('change', () => render());
  $('posTimeGran')?.addEventListener('change', () => render());
  $('matrixTopPos')?.addEventListener('change', () => render());
  $('matrixTopSrc')?.addEventListener('change', () => render());
  [
    'latestStartDate', 'latestEndDate',
    'opsStartDate', 'opsEndDate',
    'srcStartDate', 'srcEndDate',
    'posStartDate', 'posEndDate',
    'topStartDate', 'topEndDate',
    'matrixStartDate', 'matrixEndDate',
  ].forEach((id) => {
    $(id)?.addEventListener('change', () => render());
  });

  $('minImpLabel').textContent = String($('minImp').value);
  restoreBoundFolderHandle();

  $('exportReportBtn')?.addEventListener('click', async () => {
    if (!state.rows.length) {
      const h = $('statusHint');
      if (h) h.textContent = '请先加载 CSV 数据后再导出报告。';
      return;
    }
    const btn = $('exportReportBtn');
    const orig = btn?.textContent;
    if (btn) btn.textContent = '生成中…';
    try {
      const { html, dataUrl, shareUrl } = await generateReportHtml();
      const name = `运营宣推报告_${new Date().toISOString().slice(0, 10)}.html`;
      const blob = new Blob([html], { type: 'text/html;charset=utf-8' });

      const download = () => {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = name;
        a.click();
        URL.revokeObjectURL(a.href);
      };

      download();

      let urlToCopy = shareUrl;
      // 0x0.st 为第三方托管，会将报告上传至外部服务器。若需短链接，可取消下方注释并知悉数据会离开本地。
      // if (shareUrl && shareUrl.length < 60000) {
      //   try {
      //     const formData = new FormData();
      //     formData.append('file', blob, name);
      //     const res = await fetch('https://0x0.st', {
      //       method: 'POST',
      //       body: formData,
      //       headers: { 'User-Agent': 'OpsDashboard/1.0' },
      //     });
      //     if (res.ok) {
      //       const shortUrl = (await res.text()).trim();
      //       if (shortUrl && shortUrl.startsWith('http')) urlToCopy = shortUrl;
      //     }
      //   } catch (_) {}
      // }

      if (urlToCopy) {
        try {
          await navigator.clipboard.writeText(urlToCopy);
          if (btn) btn.textContent = urlToCopy.length < 100 ? '已复制短链接' : '已复制分享链接，粘贴到浏览器即可查看';
        } catch {
          if (btn) btn.textContent = '已下载，复制失败请手动分享文件';
        }
      } else {
        if (btn) btn.textContent = '已下载，可将文件上传至网盘生成分享链接';
      }
      setTimeout(() => { if (btn) btn.textContent = orig; }, 3000);
    } catch (err) {
      if (btn) btn.textContent = '导出失败';
      setTimeout(() => { if (btn) btn.textContent = orig; }, 2000);
    }
  });
}

async function generateReportHtml() {
  const reportCss = `
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,"PingFang SC","Microsoft YaHei",sans-serif;color:rgba(15,23,42,.92);background:#f6f8ff;padding:16px}
.report{max-width:1000px;margin:0 auto}
.report h1{font-size:18px;margin:0 0 4px}
.report .meta{font-size:12px;color:rgba(15,23,42,.6);margin-bottom:20px}
.report section{margin-bottom:24px;border:1px solid rgba(15,23,42,.1);border-radius:12px;background:#fff;padding:14px}
.report .sectionTitle{font-size:14px;font-weight:600;margin-bottom:10px;color:rgba(15,23,42,.9)}
.report .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0}
.report .kpi{border:1px solid rgba(15,23,42,.1);border-radius:10px;padding:10px;background:rgba(246,248,255,.8)}
.report .kpi__label{font-size:11px;color:rgba(15,23,42,.6)}
.report .kpi__value{margin-top:6px;font-family:ui-monospace,monospace;font-size:14px}
.report .adviceBox{border:1px solid rgba(15,23,42,.1);background:rgba(246,248,255,.8);border-radius:10px;padding:10px 12px;font-size:12px;line-height:1.7;margin:10px 0}
.report table{width:100%;border-collapse:collapse;font-size:12px;margin:10px 0}
.report th,.report td{padding:8px 10px;border:1px solid rgba(15,23,42,.1);text-align:left}
.report th{background:rgba(246,248,255,.95)}
.report .chartImg{max-width:100%;height:auto;border-radius:10px;border:1px solid rgba(15,23,42,.1);margin:10px 0}
.report .subTitle{font-size:12px;color:rgba(15,23,42,.6);margin:12px 0 6px}
`;

  const getHtml = (id) => { const el = document.getElementById(id); return el ? el.innerHTML : ''; };
  const getTableHtml = (id) => {
    const t = document.querySelector(`#${id}`);
    if (!t || !t.tagName || t.tagName.toLowerCase() !== 'table') return '';
    const clone = t.cloneNode(true);
    clone.querySelectorAll('.copyTableBtn').forEach((b) => b.remove());
    return clone.outerHTML;
  };
  const getMatrixHtml = () => {
    const wrap = document.getElementById('matrixTable');
    if (!wrap) return '';
    const table = wrap.querySelector('table');
    return table ? table.outerHTML : wrap.innerHTML;
  };

  const chartIds = ['opsTrendChart', 'sourceTrendChart', 'placementTrendChart', 'latestPosQuadrantChart'];
  const chartDataUrls = {};
  for (const id of chartIds) {
    const c = document.getElementById(id);
    if (c && c.tagName && c.tagName.toLowerCase() === 'canvas') {
      try {
        chartDataUrls[id] = c.toDataURL('image/png');
      } catch (_) {}
    }
  }

  let matrixImg = '';
  const matrixEl = document.getElementById('matrixTable');
  if (matrixEl && typeof html2canvas === 'function') {
    try {
      const table = matrixEl.querySelector('table');
      if (table) {
        const canvas = await html2canvas(table, { useCORS: true, scale: 2, backgroundColor: '#fff', logging: false });
        matrixImg = canvas.toDataURL('image/png');
      }
    } catch (_) {}
  }
  if (!matrixImg && matrixEl) matrixImg = ''; else if (!matrixImg) matrixImg = '';

  const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>运营宣推数据报告</title>
  <style>${reportCss}</style>
</head>
<body>
  <div class="report">
    <h1>运营宣推数据报告</h1>
    <div class="meta">导出时间：${new Date().toLocaleString('zh-CN')} · 本地解析，未上传原始数据。分享方式：将报告链接粘贴到浏览器地址栏即可查看。</div>

    <section>
      <div class="sectionTitle">结论（最新周总览与调配建议）</div>
      <div class="adviceBox">${getHtml('latestWeekSummary') || '（无数据）'}</div>
      ${getTableHtml('latestWeekWowTable') ? `<div class="subTitle">1. 当周运营宣推资源效率环比上周</div>${getTableHtml('latestWeekWowTable')}` : ''}
      ${getTableHtml('latestWeekSourceTable') ? `<div class="subTitle">2. 各宣发来源 Top5 环比</div>${getTableHtml('latestWeekSourceTable')}` : ''}
      <div class="subTitle">3. 各资源位使用情况</div>
      ${getTableHtml('latestWeekTable') || '—'}
      <div class="subTitle">4. 综合效率最高投放项目 Top5</div>
      ${getTableHtml('latestWeekTopProjectsTable') || '—'}
      <div class="subTitle">当周 KPI</div>
      <div class="kpis">${getHtml('latestWeekKpis') || getHtml('opsKpis') || '<div class="kpi"><div class="kpi__label">—</div><div class="kpi__value">—</div></div>'}</div>
      <div class="subTitle">资源位 × 宣发来源效率矩阵</div>
      ${matrixImg ? `<img src="${matrixImg}" alt="热力图" class="chartImg"/>` : (getMatrixHtml() ? getMatrixHtml() : '')}
      <div class="subTitle">宣发来源周度明细</div>
      ${getTableHtml('sourceWeeklyTable') || '—'}
    </section>

    <section>
      <div class="sectionTitle">运营宣推大盘趋势</div>
      <div class="kpis">${getHtml('opsKpis') || '—'}</div>
      <div class="adviceBox">${getHtml('opsBalanceHint') || ''}</div>
      ${chartDataUrls.opsTrendChart ? `<img src="${chartDataUrls.opsTrendChart}" alt="大盘趋势" class="chartImg"/>` : ''}
    </section>

    <section>
      <div class="sectionTitle">各业务线宣发效率趋势</div>
      ${chartDataUrls.sourceTrendChart ? `<img src="${chartDataUrls.sourceTrendChart}" alt="来源趋势" class="chartImg"/>` : ''}
    </section>

    <section>
      <div class="sectionTitle">各资源位效率趋势</div>
      ${chartDataUrls.placementTrendChart ? `<img src="${chartDataUrls.placementTrendChart}" alt="资源位趋势" class="chartImg"/>` : ''}
    </section>

    <section>
      <div class="sectionTitle">资源位效率-规模象限</div>
      ${chartDataUrls.latestPosQuadrantChart ? `<img src="${chartDataUrls.latestPosQuadrantChart}" alt="象限图" class="chartImg"/>` : ''}
    </section>

    <section>
      <div class="sectionTitle">项目效率明细</div>
      ${getTableHtml('dailyTopProjectsTable') || '—'}
    </section>
  </div>
</body>
</html>`;

  const dataUrl = `data:text/html;charset=utf-8;base64,${btoa(unescape(encodeURIComponent(html)))}`;

  // 精简版（无图表图片）用于 hash URL，保证链接可分享
  const liteHtml = html.replace(/<img[^>]+src="data:image[^"]*"[^>]*\/?>/gi, '<p class="subTitle">[图表见完整报告]</p>');
  let shareUrl = null;
  if (typeof pako !== 'undefined') {
    const bytes = pako.gzip(liteHtml, { level: 9 });
    let binary = '';
    for (let i = 0; i < bytes.length; i += 8192) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
    }
    const b64 = btoa(binary).replace(/\+/g, '-').replace(/\//g, '_');
    const base = (location.origin + location.pathname).replace(/(index\.html)?\/?$/, '/').split('?')[0];
    shareUrl = base + '#r=' + b64;
  }
  return { html, dataUrl, shareUrl };
}

if (!window.__reportView) {
  setup();
  if (window.__SNAPSHOT_DATA && window.__SNAPSHOT_DATA.opsAppRows) {
    state.rows = mergeRowsDedup(window.__SNAPSHOT_DATA.opsAppRows);
    state.lastFileName = '快照数据';
    render();
    const h = $('statusHint');
    if (h) h.textContent = `快照模式：已加载 ${state.rows.length} 行运营宣推数据。`;
  }
}

