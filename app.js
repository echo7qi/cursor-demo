/* global Papa, Chart */

const $ = (id) => document.getElementById(id);

const state = {
  rows: [],
  lastFileName: null,
  charts: {
    opsTrend: null,
    sourceTrend: null,
    placementTrend: null,
    latestSourceWow: null,
    latestPosQuadrant: null,
  },
  selectedSources: new Set(),
  selectedPlacements: new Set(),
  topPlacementFilter: '(全部)',
  topSourceFilter: '(全部)',
  boundDirHandle: null,
};

const DB_NAME = 'ops-dashboard-local-db';
const DB_STORE = 'kv';
const DB_KEY_DIR_HANDLE = 'boundDirHandle';

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
  for (const row of data) {
    const tr = document.createElement('tr');
    for (const c of columns) {
      const td = document.createElement('td');
      const v = typeof c.value === 'function' ? c.value(row) : row[c.value];
      td.innerHTML = v;
      if (c.className) td.className = c.className;
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

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

  // 0) 最新周资源效率 + 调配建议（固定运营宣推口径）
  try {
    const latestRows = filterRowsByDate(dailyRows, $('latestStartDate')?.value, $('latestEndDate')?.value);
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
      $('latestWeekSummary').innerHTML = latestNoDataInRange ? '当前筛选日期范围内没有数据，请调整开始/结束日期。' : '未找到可计算的周数据。';
      $('latestWeekKpis').innerHTML = '';
      $('latestWeekWowTable').innerHTML = '';
      $('latestWeekTable').innerHTML = '';
      $('latestWeekSourceTable').innerHTML = '';
      $('latestWeekTopProjectsTable').innerHTML = '';
      if ($('latestSourceWowFallback')) $('latestSourceWowFallback').innerHTML = '';
      if ($('latestPosQuadrantFallback')) $('latestPosQuadrantFallback').innerHTML = '';
      if ($('latestSourceWowHint')) $('latestSourceWowHint').textContent = '';
      if ($('latestPosQuadrantHint')) $('latestPosQuadrantHint').textContent = '';
      const c1 = state.charts.latestSourceWow;
      if (c1) { c1.destroy(); state.charts.latestSourceWow = null; }
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

      $('latestWeekKpis').innerHTML = [
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

      // by placement
      const byPos = groupBy(weekRows, (r) => String(r[COLS.placement] ?? '').trim() || '(空)');
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

      const weekCols = [
        { label: '资源位', value: (r) => r.pos },
        { label: '运营宣推曝光占比', className: 'num', value: (r) => fmtRate(r.shareOps, 2) },
        { label: '运营宣推曝光', className: 'num', value: (r) => fmtInt(r.imp) },
        { label: '运营宣推收入', className: 'num', value: (r) => fmtMoney(r.rev, 2) },
        { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.rev_per_imp, 4) },
        { label: 'p-CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
        { label: '阅读率', className: 'num', value: (r) => fmtRate(r.read_rate, 2) },
        { label: '付费率', className: 'num', value: (r) => fmtRate(r.pay_rate, 2) },
        { label: '建议标签', value: (r) => r.tag },
      ];
      buildTable($('latestWeekTable'), weekCols, posRows.slice(0, 20));

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
      buildTable($('latestWeekSourceTable'), srcCols, srcAllWow);

      // 来源环比图（Top5曝光）
      const srcChartEl = $('latestSourceWowChart');
      const srcHintEl = $('latestSourceWowHint');
      const srcFallbackEl = $('latestSourceWowFallback');
      const top5 = srcAllWow.slice(0, 5);
      if (srcFallbackEl) {
        srcFallbackEl.innerHTML = `<div class="fallbackList">${
          top5.map((x) =>
            `<div class="fallbackItem"><span>${x.src}</span><span>每曝光收入环比 ${fmtWow(x.erpi_wow)} ｜ p-CTR环比 ${fmtWow(x.pctr_wow)}</span></div>`
          ).join('')
        }</div>`;
      }

      if (typeof Chart !== 'undefined' && srcChartEl) {
        const labels = top5.map((x) => x.src);
        const erpiWow = top5.map((x) => x.erpi_wow ?? 0);
        const pctrWow = top5.map((x) => x.pctr_wow ?? 0);
        const ctx = srcChartEl.getContext('2d');
        if (state.charts.latestSourceWow) state.charts.latestSourceWow.destroy();
        state.charts.latestSourceWow = new Chart(ctx, {
          type: 'bar',
          data: {
            labels,
            datasets: [
              { label: '每曝光收入环比', data: erpiWow, backgroundColor: 'rgba(79,124,255,.65)' },
              { label: 'p-CTR环比', data: pctrWow, backgroundColor: 'rgba(64,199,154,.65)' },
            ],
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              y: { ticks: { callback: (v) => `${Number(v * 100).toFixed(0)}%` }, grid: { color: 'rgba(15,23,42,.06)' } },
              x: { grid: { display: false } },
            },
            plugins: { tooltip: { callbacks: { label: (ctx2) => `${ctx2.dataset.label}: ${fmtWow(ctx2.parsed.y)}` } } },
          },
        });
        if (srcHintEl) srcHintEl.textContent = '图：Top5来源的每曝光收入环比与p-CTR环比（柱状对比）';
        if (srcFallbackEl) srcFallbackEl.style.display = 'none';
      } else {
        if (srcHintEl) srcHintEl.textContent = '图表库未加载，已显示列表版环比结果。';
        if (srcFallbackEl) srcFallbackEl.style.display = '';
      }

      // resource usage summary top5 by exposure
      const posUsageTop = [...posRows].sort((x, y) => y.imp - x.imp).slice(0, 5);
      const prevPosMap = new Map();
      if (prevWeek) {
        const byPosPrev = groupBy(prevWeek.rows, (r) => String(r[COLS.placement] ?? '').trim() || '(空)');
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

      // top5 projects in current week (ops scope)
      const byTopic = groupBy(weekRows, (r) => String(r[COLS.topicId] ?? '').trim() || String(r[COLS.topicName] ?? '').trim() || '(未知)');
      const topicTop = Array.from(byTopic.entries()).map(([tid, rows]) => {
        const ta = computeFunnelAgg(rows, opsCols);
        const td = computeDerived(ta);
        const name = String(rows[0]?.[COLS.topicName] ?? '').trim() || '(空)';
        const src = String(rows[0]?.[COLS.source] ?? '').trim() || '(空)';
        const pos = String(rows[0]?.[COLS.placement] ?? '').trim() || '(空)';
        return { tid, name, src, pos, ...ta, ...td };
      }).filter((x) => x.imp >= Math.max(10000, Math.floor(minImp / 2)))
        .sort((x, y) => y.rev_per_imp - x.rev_per_imp)
        .slice(0, 5);

      const topicCols = [
        { label: '专题ID', className: 'num', value: (r) => r.tid },
        { label: '专题名称', value: (r) => r.name },
        { label: '宣发来源', value: (r) => r.src },
        { label: '资源位', value: (r) => r.pos },
        { label: '运营宣推曝光', className: 'num', value: (r) => fmtInt(r.imp) },
        { label: 'p-CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
        { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.rev_per_imp, 4) },
      ];
      buildTable($('latestWeekTopProjectsTable'), topicCols, topicTop);

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
        if (quadHintEl) quadHintEl.textContent = '横轴=资源位运营宣推曝光占比，纵轴=每曝光收入，气泡大小≈曝光量';
        if (quadFallbackEl) quadFallbackEl.style.display = 'none';
      } else {
        if (quadHintEl) quadHintEl.textContent = '图表库未加载，已显示列表版资源位效率。';
        if (quadFallbackEl) quadFallbackEl.style.display = '';
      }

      const high = posRows.filter((r) => r.tag.includes('高效')).slice(0, 3);
      const low = posRows.filter((r) => r.tag.includes('低效')).slice(0, 3);

      const summaryLines = [];
      summaryLines.push(`<strong>当前周：</strong>${weekStart} ~ ${weekEnd}${days < 7 ? '（非完整周）' : ''}`);
      summaryLines.push(
        `<strong>1. 当周运营宣推资源效率环比上周：</strong>` +
        `每曝光收入 ${fmtWow(wow(d.rev_per_imp, prevD?.rev_per_imp))}，` +
        `p-CTR ${fmtWow(wow(d.ctr, prevD?.ctr))}，` +
        `阅读率 ${fmtWow(wow(d.read_rate, prevD?.read_rate))}，` +
        `付费率 ${fmtWow(wow(d.pay_rate, prevD?.pay_rate))}，` +
        `运营宣推曝光占比 ${fmtWow(wow(opsShareGlobal, prevOpsShareGlobal))}`
      );
      summaryLines.push(
        `<strong>2. 各宣发来源（Top5曝光）p-CTR与每曝光收入环比：</strong><ul>` +
        srcTop.map((x) => {
          const p = srcPrevMap.get(x.src);
          return `<li>${x.src}：p-CTR ${fmtRate(x.ctr, 2)}（环比 ${fmtWow(wow(x.ctr, p?.ctr))}）；每曝光收入 ${fmtMoney(x.rev_per_imp, 4)}（环比 ${fmtWow(wow(x.rev_per_imp, p?.rev_per_imp))}）</li>`;
        }).join('') +
        `</ul>`
      );
      summaryLines.push(
        `<strong>3. 各资源使用情况（按曝光占比Top5）：</strong><ul>` +
        posUsageTop.map((x) => `<li>${x.pos}：占比 ${fmtRate(x.shareOps, 2)}，每曝光收入 ${fmtMoney(x.rev_per_imp, 4)}，标签：${x.tag}</li>`).join('') +
        `</ul>`
      );
      summaryLines.push(
        `<strong>4. 综合效率最高投放项目 Top5（当周）：</strong><ul>` +
        topicTop.map((x) => `<li>${x.name}（${x.tid}）| ${x.src} / ${x.pos} | 每曝光收入 ${fmtMoney(x.rev_per_imp, 4)} | p-CTR ${fmtRate(x.ctr, 2)}</li>`).join('') +
        `</ul>`
      );
      summaryLines.push(
        `<strong>调配建议：</strong>` +
        `${high.length ? `优先加量 ${high.map((x) => x.pos).join('、')}` : '优先维持高效资源位'}；` +
        `${low.length ? `优先整改/挪量 ${low.map((x) => x.pos).join('、')}` : '低效大盘不明显'}。` +
        `建议先做 5%~10% 小步调配，观察 2-3 天再放大。`
      );

      $('latestWeekSummary').innerHTML = summaryLines.join('<br/>');

    }
  } catch (err) {
    $('latestWeekSummary').innerHTML = `最新周模块渲染异常：${err?.message || err}`;
    $('latestWeekKpis').innerHTML = '';
    $('latestWeekWowTable').innerHTML = '';
    $('latestWeekTable').innerHTML = '';
    $('latestWeekSourceTable').innerHTML = '';
    $('latestWeekTopProjectsTable').innerHTML = '';
    if ($('latestSourceWowHint')) $('latestSourceWowHint').textContent = '图表渲染异常。';
    if ($('latestPosQuadrantHint')) $('latestPosQuadrantHint').textContent = '图表渲染异常。';
    const c1 = state.charts.latestSourceWow;
    if (c1) { c1.destroy(); state.charts.latestSourceWow = null; }
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

    const wowSuffix = (v) => (v != null ? ` <span class="muted">（周环比${v}）</span>` : '');

    $('opsKpis').innerHTML = [
      kpi('运营宣推曝光占比（全期）', fmtRate(totalShare, 2) + wowSuffix(wowShare)),
      kpi('运营宣推每曝光收入（全期）', `<span class="num">${fmtMoney(totalDer.rev_per_imp, 4)}</span>${wowSuffix(wowRevPerImp)}`),
      kpi('运营宣推CTR（全期）', fmtRate(totalDer.ctr, 2) + wowSuffix(wowCtr)),
      kpi('运营宣推阅读率（全期）', fmtRate(totalDer.read_rate, 2) + wowSuffix(wowReadRate)),
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

    const cols = [
      { label: timeGran === 'day' ? '日期' : '周起始(周一)', value: (r) => `<span class="num">${r.date}</span>` },
      ...(timeGran === 'day' ? [] : [{ label: '覆盖天数', className: 'num', value: (r) => `<span class="num">${r.days ?? ''}</span>` }]),
      { label: '曝光占比', className: 'num', value: (r) => fmtRate(r.opsShare, 2) },
      { label: '运营宣推曝光', className: 'num', value: (r) => fmtInt(r.imp) },
      { label: '运营宣推点击', className: 'num', value: (r) => fmtInt(r.clk) },
      { label: '运营宣推阅读', className: 'num', value: (r) => fmtInt(r.read) },
      { label: '运营宣推付费用户', className: 'num', value: (r) => fmtInt(r.payu) },
      { label: '运营宣推收入', className: 'num', value: (r) => fmtMoney(r.rev, 2) },
      { label: '每曝光收入', className: 'num', value: (r) => fmtMoney(r.rev_per_imp, 4) },
      { label: 'CTR', className: 'num', value: (r) => fmtRate(r.ctr, 2) },
      { label: '阅读率', className: 'num', value: (r) => fmtRate(r.read_rate, 2) },
      { label: '付费率', className: 'num', value: (r) => fmtRate(r.pay_rate, 2) },
    ];
    buildTable($('opsDailyTable'), cols, series.slice(-60));
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

        const erpiVals = buckets.map((b) => {
          const r = b.rows.filter((x) => (String(x[COLS.source] ?? '').trim() || '(空)') === src);
          const a = computeFunnelAgg(r, opsScopeCols);
          const d = computeDerived(a);
          return d.rev_per_imp ?? 0;
        });
        datasets.push({
          label: `${src} · 每曝光收入`,
          data: erpiVals,
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
                  return `${ctx2.dataset.label}: ${Number(v).toFixed(4)}`;
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
                callback: (v) => Number(v).toFixed(4),
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
        hintEl.textContent = `趋势：${granLabel} · 口径=运营宣推 · 每来源两条线：实线=每曝光收入，虚线=p-CTR（运营宣推点击/运营宣推曝光） · 已选来源=${selected.length}`;
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

  // 4) 每天综合效率最高的项目明细（最近 7 天 Top 10）
  {
    const topRowsData = filterRowsByDate(dailyRows, $('topStartDate')?.value, $('topEndDate')?.value);
    const byDate = groupBy(topRowsData, (r) => String(r[COLS.date] ?? '').trim());
    const dates = Array.from(byDate.keys()).filter(Boolean).sort();

    // 资源位筛选 pill
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

    const topRows = [];
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
      const top10 = topics.slice(0, 10);
      for (let i = 0; i < top10.length; i += 1) {
        topRows.push({ rank: i + 1, ...top10[i] });
      }
    }

    const cols = [
      { label: '日期', value: (r) => `<span class="num">${r.date}</span>` },
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
    $('statusHint').textContent = '正在扫描并合并文件夹中的 CSV...';

    const allRows = [];
    let fileCount = 0;
    for await (const entry of state.boundDirHandle.values()) {
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
      $('statusHint').textContent = '绑定文件夹下未找到 CSV 文件。';
      return;
    }
    await loadRows(allRows, `文件夹模式（${fileCount}个CSV）`);
    $('statusHint').textContent = `更新完成：已合并 ${fileCount} 个 CSV，去重后 ${state.rows.length.toLocaleString('zh-CN')} 行。`;
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

function setup() {
  window.addEventListener('error', (e) => {
    const hint = $('statusHint');
    if (hint) hint.textContent = `页面脚本异常：${e.message || 'unknown error'}`;
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
}

setup();

