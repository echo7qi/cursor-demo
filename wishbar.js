/* global Papa, Chart */

const $ = (id) => document.getElementById(id);

const DB_NAME = 'ops-dashboard-local-db';
const DB_STORE = 'kv';
const DB_KEY_DIR_HANDLE = 'boundDirHandle';

const COLS = {
  date: '日期',
  period: '统计周期',
  l1: '一级来源',
  l2: '二级来源',
  uv: '祈愿bar访问uv',
};

const OPS_ALIASES = new Set(['广告_资源投放', '广告资源投放', 'v2_资源投放', 'v2资源投放']);

const WISH_SUBDIR_CANDIDATES = ['祈愿', '祈愿bar来源'];

function num(v) {
  if (v == null) return 0;
  const s = String(v).trim();
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function normSource(s) {
  const v = (s == null ? '' : String(s)).trim();
  if (!v || v === '--' || v.toLowerCase() === 'nan') return '(空)';
  if (OPS_ALIASES.has(v)) return '运营资源投放';
  return v;
}

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data || []),
      error: (err) => reject(err),
    });
  });
}

function openDb() {
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

async function dbGet(key) {
  const db = await openDb();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readonly');
      const store = tx.objectStore(DB_STORE);
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

async function dbSet(key, value) {
  const db = await openDb();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, 'readwrite');
      const store = tx.objectStore(DB_STORE);
      const req = store.put(value, key);
      req.onsuccess = () => resolve();
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

async function getBoundDirHandle() {
  try {
    const h = await dbGet(DB_KEY_DIR_HANDLE);
    return h || null;
  } catch (_) {
    return null;
  }
}

async function setBoundDirHandle(handle) {
  await dbSet(DB_KEY_DIR_HANDLE, handle);
}

async function pickAndBindFolder() {
  if (!window.showDirectoryPicker) {
    throw new Error('当前浏览器不支持“绑定文件夹”（需要 File System Access API）。');
  }
  // 有些浏览器会拦截 iframe 内的目录选择器
  if (window.top && window.top !== window.self) {
    throw new Error('当前页面在汇总页(iframe)内，浏览器可能拦截目录选择器。请在新窗口打开祈愿bar来源页再绑定。');
  }
  const handle = await window.showDirectoryPicker({ mode: 'read' });
  await setBoundDirHandle(handle);
  return handle;
}

async function resolveWishSubdir(rootHandle) {
  for (const name of WISH_SUBDIR_CANDIDATES) {
    try {
      const h = await rootHandle.getDirectoryHandle(name, { create: false });
      return { handle: h, name };
    } catch (_) {}
  }
  const list = WISH_SUBDIR_CANDIDATES.map((x) => `「${x}」`).join(' 或 ');
  throw new Error(`未找到祈愿数据子文件夹（需要在绑定目录下创建 ${list} 文件夹）。`);
}

async function parseCsvText(text) {
  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      complete: (res) => resolve(res.data || []),
      error: (err) => reject(err),
    });
  });
}

function fmtPct(v) {
  if (!Number.isFinite(v)) return '0.00%';
  return `${v.toFixed(2)}%`;
}

function buildWeeklyShares(rows) {
  const weekL1Uv = new Map(); // week -> Map(l1 -> uv)
  const weekL2Uv = new Map(); // week -> Map(l2 -> uv)

  for (const r of rows) {
    if (!r || r[COLS.period] == null) continue;
    const period = String(r[COLS.period]).trim();
    if (period !== '周') continue;

    const week = String(r[COLS.date] ?? '').trim();
    if (!week) continue;

    const l1 = normSource(r[COLS.l1]);
    const l2 = normSource(r[COLS.l2]);
    const uv = num(r[COLS.uv]);
    if (uv === 0) continue;

    if (!weekL1Uv.has(week)) weekL1Uv.set(week, new Map());
    if (!weekL2Uv.has(week)) weekL2Uv.set(week, new Map());
    const m1 = weekL1Uv.get(week);
    const m2 = weekL2Uv.get(week);
    m1.set(l1, (m1.get(l1) || 0) + uv);
    m2.set(l2, (m2.get(l2) || 0) + uv);
  }

  const weeks = Array.from(new Set([...weekL1Uv.keys(), ...weekL2Uv.keys()])).sort(); // yyyy-mm-dd
  const l1Set = new Set();
  const l2Set = new Set();
  for (const w of weeks) {
    for (const s of (weekL1Uv.get(w) || new Map()).keys()) l1Set.add(s);
    for (const s of (weekL2Uv.get(w) || new Map()).keys()) l2Set.add(s);
  }
  const l1All = Array.from(l1Set);
  const l2All = Array.from(l2Set);

  // compute totals and shares (same total for both dims)
  const totalsByWeek = new Map(); // week -> total_uv
  for (const w of weeks) {
    const m = weekL1Uv.get(w) || new Map();
    let total = 0;
    for (const v of m.values()) total += v;
    totalsByWeek.set(w, total);
  }

  const l1SharesByWeek = new Map(); // week -> Map(l1 -> share_pct)
  const l2SharesByWeek = new Map(); // week -> Map(l2 -> share_pct)
  for (const w of weeks) {
    const total = totalsByWeek.get(w) || 0;
    const m1 = weekL1Uv.get(w) || new Map();
    const m2 = weekL2Uv.get(w) || new Map();
    const s1 = new Map();
    const s2 = new Map();
    for (const s of l1All) s1.set(s, total > 0 ? ((m1.get(s) || 0) / total) * 100 : 0);
    for (const s of l2All) s2.set(s, total > 0 ? ((m2.get(s) || 0) / total) * 100 : 0);
    l1SharesByWeek.set(w, s1);
    l2SharesByWeek.set(w, s2);
  }

  return {
    weeks,
    l1All,
    l2All,
    l1SharesByWeek,
    l2SharesByWeek,
    l1UvByWeek: weekL1Uv,
    l2UvByWeek: weekL2Uv,
    totalsByWeek,
  };
}

function pickTopKByLatest(sharesByWeek, latestWeek, sourcesAll, topK) {
  const m = sharesByWeek.get(latestWeek) || new Map();
  const arr = sourcesAll
    .map((s) => ({ s, v: m.get(s) || 0 }))
    .sort((a, b) => b.v - a.v);

  if (topK >= 999) return { sources: arr.map((x) => x.s), foldOther: false };
  const sources = arr.slice(0, topK).map((x) => x.s);
  return { sources, foldOther: true };
}

function renderTable({ weeks, shownSources, foldOther, sharesByWeek, freezeWeek }, tableEl) {
  const cols = [...shownSources, ...(foldOther ? ['其他'] : [])];

  const sticky = freezeWeek ? ' class="stickyCol"' : '';
  let html = `<thead><tr><th${sticky}>周</th>`;
  for (const c of cols) html += `<th>${c}</th>`;
  html += '</tr></thead><tbody>';

  for (const w of weeks) {
    const m = sharesByWeek.get(w);
    let other = 0;
    html += `<tr><td${sticky} class="mono${freezeWeek ? ' stickyCol' : ''}">${w}</td>`;
    for (const s of shownSources) html += `<td>${fmtPct(m.get(s) || 0)}</td>`;
    if (foldOther) {
      for (const [s, v] of m.entries()) if (!shownSources.includes(s)) other += v;
      html += `<td>${fmtPct(other)}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody>';

  tableEl.innerHTML = html;
}

function palette(n) {
  const base = [
    '#4f7cff',
    '#40c79a',
    '#f59e0b',
    '#ef4444',
    '#a855f7',
    '#06b6d4',
    '#22c55e',
    '#e11d48',
    '#0ea5e9',
    '#84cc16',
    '#f97316',
    '#64748b',
  ];
  const out = [];
  for (let i = 0; i < n; i++) out.push(base[i % base.length]);
  return out;
}

function buildDatasets({ weeks, shownSources, foldOther, sharesByWeek }) {
  const cols = [...shownSources, ...(foldOther ? ['其他'] : [])];
  const colors = palette(cols.length);
  const datasets = cols.map((name, idx) => {
    const data = weeks.map((w) => {
      const m = sharesByWeek.get(w);
      if (name === '其他') {
        let other = 0;
        for (const [s, v] of m.entries()) if (!shownSources.includes(s)) other += v;
        return other;
      }
      return m.get(name) || 0;
    });
    return {
      label: name,
      data,
      borderColor: colors[idx],
      backgroundColor: colors[idx] + '22',
      pointRadius: 2,
      pointHoverRadius: 4,
      tension: 0.25,
      fill: false,
      borderWidth: name === '(空)' ? 2.2 : 1.8,
      borderDash: name === '(空)' ? [6, 4] : undefined,
    };
  });
  return { labels: weeks, datasets };
}

function makeChartConfig(type, data) {
  let maxVal = 0;
  if (data && Array.isArray(data.datasets)) {
    for (const ds of data.datasets) {
      for (const v of ds.data) {
        if (Number.isFinite(v) && v > maxVal) maxVal = v;
      }
    }
  }
  // 留一点顶部空间，但不超过 100%
  const yMax = Math.min(100, Math.ceil((maxVal + 5) / 5) * 5 || 100);

  const common = {
    data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      layout: {
        padding: {
          top: 10,
          bottom: 10,
        },
      },
      plugins: {
        legend: {
          position: 'bottom',
          align: 'center',
          labels: { usePointStyle: true, boxWidth: 10, padding: 12 },
          padding: 50, // 折线与图例之间约 50px 间距
        },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${fmtPct(ctx.parsed.y)}`,
          },
        },
      },
      scales: {
        y: {
          min: 0,
          max: yMax,
          ticks: { callback: (v) => `${v}%` },
          grid: { color: 'rgba(15,23,42,.08)' },
        },
        x: {
          grid: { display: false },
          ticks: {
            maxRotation: 45,
            minRotation: 30,
            autoSkip: true,
            maxTicksLimit: 10,
          },
        },
      },
    },
  };

  return {
    type: 'line',
    ...common,
    options: {
      ...common.options,
      elements: { line: { borderWidth: 1.5, fill: false }, point: { radius: 2 } },
    },
  };
}

let chart = null;
let volumeChart = null;
let parsed = null;
let trendSelectedL1 = new Set();
let trendSelectedL2 = new Set();
let trendInitialized = false;
let barSelectedL1 = new Set();
let barSelectedL2 = new Set();
let barInitialized = false;

function setStatus(msg) {
  $('wishbarStatusHint').textContent = msg;
}

function buildSourceOrderByLatest(parsedData) {
  const latestWeek = parsedData.weeks[parsedData.weeks.length - 1];
  const m = parsedData.l1SharesByWeek.get(latestWeek) || new Map();
  return parsedData.l1All
    .map((s) => ({ s, v: m.get(s) || 0 }))
    .sort((a, b) => b.v - a.v)
    .map((x) => x.s);
}

function buildL2OrderByLatest(parsedData) {
  const latestWeek = parsedData.weeks[parsedData.weeks.length - 1];
  const m = parsedData.l2SharesByWeek.get(latestWeek) || new Map();
  return parsedData.l2All
    .map((s) => ({ s, v: m.get(s) || 0 }))
    .sort((a, b) => b.v - a.v)
    .map((x) => x.s);
}

function summarizeSelection(set, maxItems = 3) {
  const arr = Array.from(set);
  if (arr.length === 0) return '未选择';
  const head = arr.slice(0, maxItems).join('、');
  return arr.length > maxItems ? `${head} 等 ${arr.length} 项` : `${head}（${arr.length}）`;
}

function ensureFilterPopover() {
  let pop = document.getElementById('filterPopover');
  if (pop) return pop;

  pop = document.createElement('div');
  pop.id = 'filterPopover';
  pop.className = 'popover';
  pop.innerHTML = `
    <div class="popover__head">
      <div class="popover__title" id="filterPopTitle">筛选</div>
      <button class="btn btn--ghost btn--sm" id="filterPopClose" type="button">关闭</button>
    </div>
    <div class="popover__body">
      <div class="filterSearch">
        <span style="color:rgba(15,23,42,.55)">🔎</span>
        <input id="filterPopSearch" type="text" placeholder="搜索…" />
      </div>
      <div class="filterRow" style="margin-top:10px">
        <div class="filterRow__left">
          <button class="btn btn--ghost btn--sm" id="filterPopSelectAll" type="button">全选</button>
          <button class="btn btn--ghost btn--sm" id="filterPopInvert" type="button">反选</button>
          <button class="btn btn--ghost btn--sm" id="filterPopClear" type="button">清空</button>
        </div>
        <div class="filterRow__right" id="filterPopCount">已选：0</div>
      </div>
      <div class="filterList" id="filterPopList"></div>
    </div>
    <div class="popover__footer">
      <button class="btn btn--ghost" id="filterPopCancel" type="button">取消</button>
      <button class="btn" id="filterPopOk" type="button">确定</button>
    </div>
  `;
  document.body.appendChild(pop);
  return pop;
}

function openFilterPopover({ anchorEl, title, items, selectedSet, onApply }) {
  const pop = ensureFilterPopover();
  const titleEl = document.getElementById('filterPopTitle');
  const closeBtn = document.getElementById('filterPopClose');
  const cancelBtn = document.getElementById('filterPopCancel');
  const okBtn = document.getElementById('filterPopOk');
  const searchEl = document.getElementById('filterPopSearch');
  const listEl = document.getElementById('filterPopList');
  const countEl = document.getElementById('filterPopCount');
  const btnAll = document.getElementById('filterPopSelectAll');
  const btnInv = document.getElementById('filterPopInvert');
  const btnClr = document.getElementById('filterPopClear');

  const temp = new Set(Array.from(selectedSet));
  let q = '';

  function visibleItems() {
    const qq = q.trim().toLowerCase();
    return qq ? items.filter((x) => String(x).toLowerCase().includes(qq)) : items;
  }
  function updateCount() {
    countEl.textContent = `已选：${temp.size}`;
  }
  function renderList() {
    const vis = visibleItems();
    listEl.innerHTML = vis
      .map((s) => {
        const checked = temp.has(s) ? 'checked' : '';
        const safe = String(s).replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return `<label class="filterItem"><input type="checkbox" value="${safe}" ${checked} />${safe}</label>`;
      })
      .join('');
    listEl.querySelectorAll('input[type="checkbox"]').forEach((el) => {
      el.addEventListener('change', () => {
        const v = el.value;
        if (el.checked) temp.add(v);
        else temp.delete(v);
        updateCount();
      });
    });
  }

  function close() {
    pop.style.display = 'none';
    document.removeEventListener('mousedown', onOutsideDown, true);
    document.removeEventListener('keydown', onKey, true);
  }
  function onKey(e) {
    if (e.key === 'Escape') close();
  }
  function onOutsideDown(e) {
    if (!pop.contains(e.target) && e.target !== anchorEl) close();
  }

  titleEl.textContent = title;
  searchEl.value = '';
  q = '';
  updateCount();
  renderList();

  closeBtn.onclick = close;
  cancelBtn.onclick = close;
  okBtn.onclick = () => {
    onApply(new Set(temp));
    close();
  };
  btnClr.onclick = () => {
    temp.clear();
    updateCount();
    renderList();
  };
  btnAll.onclick = () => {
    for (const s of visibleItems()) temp.add(s);
    updateCount();
    renderList();
  };
  btnInv.onclick = () => {
    const vis = visibleItems();
    for (const s of vis) {
      if (temp.has(s)) temp.delete(s);
      else temp.add(s);
    }
    updateCount();
    renderList();
  };
  searchEl.oninput = () => {
    q = searchEl.value || '';
    renderList();
  };

  // Position below anchor
  const rect = anchorEl.getBoundingClientRect();
  const gap = 8;
  const vw = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
  const vh = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
  const width = Math.min(720, vw - 24);
  let left = Math.max(12, rect.left);
  left = Math.min(left, vw - width - 12);
  let top = rect.bottom + gap;
  // If too low, place above
  const estH = 520;
  if (top + estH > vh - 12) top = Math.max(12, rect.top - gap - estH);

  pop.style.width = `${width}px`;
  pop.style.left = `${left}px`;
  pop.style.top = `${top}px`;
  pop.style.display = 'block';

  setTimeout(() => searchEl.focus(), 0);
  document.addEventListener('mousedown', onOutsideDown, true);
  document.addEventListener('keydown', onKey, true);
}

function renderTrendFilterSummaries() {
  const s1 = $('wishbarTrendL1Summary');
  const s2 = $('wishbarTrendL2Summary');
  if (s1) s1.textContent = summarizeSelection(trendSelectedL1);
  if (s2) s2.textContent = summarizeSelection(trendSelectedL2);
}

// 模块2日期筛选：单个“截止日期”（见 wishbarTrendEndOnly）

function initTrendDefaults(parsedData, topSources, foldOther) {
  if (trendInitialized) return;
  trendSelectedL1 = new Set([...topSources, ...(foldOther ? ['其他'] : [])]);
  trendSelectedL2 = new Set();
  trendInitialized = true;
  renderTrendFilterSummaries();
}

function renderBarFilterSummaries() {
  const s1 = $('wishbarBarL1Summary');
  const s2 = $('wishbarBarL2Summary');
  if (s1) s1.textContent = summarizeSelection(barSelectedL1);
  if (s2) s2.textContent = summarizeSelection(barSelectedL2);
}

function initBarDefaults(parsedData, topSources, foldOther) {
  if (barInitialized) return;
  barSelectedL1 = new Set([...topSources, ...(foldOther ? ['其他'] : [])]);
  barSelectedL2 = new Set();
  barInitialized = true;
  renderBarFilterSummaries();
}

function render() {
  if (!parsed || parsed.weeks.length === 0) {
    setStatus('未找到可用数据（需要统计周期=周，且包含日期/一级来源/祈愿bar访问uv）。');
    $('wishbarShareTable').innerHTML = '';
    if (chart) {
      chart.destroy();
      chart = null;
    }
    if (volumeChart) {
      volumeChart.destroy();
      volumeChart = null;
    }
    return;
  }

  const latestWeek = parsed.weeks[parsed.weeks.length - 1];
  const pickedWeek = $('wishbarWeekPick')?.value ? String($('wishbarWeekPick').value) : latestWeek;
  const weekForBar = parsed.weeks.includes(pickedWeek) ? pickedWeek : latestWeek;
  const topK = 8; // 默认Top8（不再暴露TopK筛选控件）
  const { sources: topSources, foldOther } = pickTopKByLatest(
    parsed.l1SharesByWeek,
    latestWeek,
    parsed.l1All,
    topK
  );

  initBarDefaults(parsed, topSources, foldOther);
  initTrendDefaults(parsed, topSources, foldOther);
  renderBarFilterSummaries();
  renderTrendFilterSummaries();

  // 模块1：最新周柱状（一级+二级并集）
  const barL1 = Array.from(barSelectedL1);
  const barL2 = Array.from(barSelectedL2);
  const barShowOther = foldOther && barL1.includes('其他');
  const barL1Clean = barL1.filter((s) => s !== '其他');
  const barL2Clean = barL2;

  if (barL1Clean.length === 0 && barL2Clean.length === 0 && !barShowOther) {
    if (chart) {
      chart.destroy();
      chart = null;
    }
    if (volumeChart) {
      volumeChart.destroy();
      volumeChart = null;
    }
    const hint = `最新周：${latestWeek}；未选择任何来源，请在一级/二级来源筛选里选择。`;
    $('wishbarChartHint').textContent = hint;
    const volHintEl = $('wishbarVolumeHint');
    if (volHintEl) volHintEl.textContent = '请先选择来源后展示趋势折线。';
    setStatus(`已载入 ${parsed.weeks.length} 个周。${hint}`);
    return;
  }

  const l1Map = parsed.l1SharesByWeek.get(weekForBar) || new Map();
  const l2Map = parsed.l2SharesByWeek.get(weekForBar) || new Map();
  const barSeries = [];
  // 柱状图标签也保持干净：默认只显示名称，同名时再区分一级/二级
  const barNameCounts = new Map();
  for (const n of barL1Clean) barNameCounts.set(n, (barNameCounts.get(n) || 0) + 1);
  for (const n of barL2Clean) barNameCounts.set(n, (barNameCounts.get(n) || 0) + 1);

  for (const name of barL1Clean) {
    const label = (barNameCounts.get(name) || 0) > 1 ? `${name}（一级）` : name;
    barSeries.push({ label, value: l1Map.get(name) || 0 });
  }
  if (barShowOther) {
    let other = 0;
    for (const [s, v] of l1Map.entries()) if (!barL1Clean.includes(s)) other += v;
    barSeries.push({ label: '其他（一级）', value: other });
  }
  for (const name of barL2Clean) {
    const label = (barNameCounts.get(name) || 0) > 1 ? `${name}（二级）` : name;
    barSeries.push({ label, value: l2Map.get(name) || 0 });
  }
  barSeries.sort((a, b) => b.value - a.value);
  const sortedLabels = barSeries.map((x) => x.label);
  const sortedValues = barSeries.map((x) => x.value);

  const colors = palette(sortedLabels.length);
  const canvas = $('wishbarShareChart');
  canvas.parentElement.style.height = '420px';

  if (chart) chart.destroy();
  chart = new Chart(canvas.getContext('2d'), {
    type: 'bar',
    data: {
      labels: sortedLabels,
      datasets: [
        {
          label: '最新周占比',
          data: sortedValues,
          backgroundColor: colors.map((c) => c + '44'),
          borderColor: colors,
          borderWidth: 1.2,
          borderRadius: 6,
          barThickness: 18,
        },
      ],
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'nearest', intersect: true },
      layout: { padding: { top: 10, bottom: 10 } },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.label}: ${fmtPct(ctx.parsed.x)}`,
          },
        },
      },
      scales: {
        x: {
          min: 0,
          max: Math.min(100, Math.ceil((Math.max(...sortedValues) + 5) / 5) * 5 || 100),
          ticks: { callback: (v) => `${v}%` },
          grid: { color: 'rgba(15,23,42,.08)' },
        },
        y: {
          grid: { display: false },
        },
      },
    },
  });

  // 明细表：按周展示“模块1选择的一级来源”（+可选其他）的占比
  const tableEl = $('wishbarShareTable');
  if (tableEl) {
    const cols = [...barL1Clean];
    const showOtherCol = barShowOther;
    const weeks = parsed.weeks;
    let html = '<thead><tr><th>周</th>';
    for (const c of cols) html += `<th>${c}</th>`;
    if (showOtherCol) html += '<th>其他</th>';
    html += '</tr></thead><tbody>';
    for (const w of weeks) {
      const m = parsed.l1SharesByWeek.get(w) || new Map();
      html += `<tr><td class="mono">${w}</td>`;
      for (const c of cols) html += `<td>${fmtPct(m.get(c) || 0)}</td>`;
      if (showOtherCol) {
        let other = 0;
        for (const [k, v] of m.entries()) if (!cols.includes(k)) other += v;
        html += `<td>${fmtPct(other)}</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody>';
    tableEl.innerHTML = html;
  }

  const hint = `最新周：${latestWeek}；柱状展示：${sortedLabels.length} 个来源。`;
  $('wishbarChartHint').textContent = hint;
  setStatus(`已载入 ${parsed.weeks.length} 个周。当前日期：${weekForBar}。${hint}`);

  // 二维并集趋势折线图（周）
  // 日期筛选：仅作用于趋势图（周）
  const trendEndOnly = $('wishbarTrendEndOnly')?.value ? String($('wishbarTrendEndOnly').value) : '';
  let trendWeeks = parsed.weeks;
  if (trendEndOnly) trendWeeks = trendWeeks.filter((w) => w <= trendEndOnly);

  if (!trendWeeks.length) {
    const volHintEl = $('wishbarVolumeHint');
    if (volHintEl) volHintEl.textContent = '当前日期筛选下无可用周数据，请调整开始/结束日期。';
    if (volumeChart) {
      volumeChart.destroy();
      volumeChart = null;
    }
    return;
  }

  const trendSeries = [];
  const trendL1 = Array.from(trendSelectedL1);
  const trendL2 = Array.from(trendSelectedL2);
  const trendShowOther = foldOther && trendL1.includes('其他');
  const trendL1Clean = trendL1.filter((s) => s !== '其他');
  const trendL2Clean = trendL2;
  const seriesKeys = [
    ...trendL1Clean.map((s) => ({ type: 'l1', name: s })),
    ...(trendShowOther ? [{ type: 'l1', name: '其他' }] : []),
    ...trendL2Clean.map((s) => ({ type: 'l2', name: s })),
  ];
  const nameCounts = new Map();
  for (const k of seriesKeys) nameCounts.set(k.name, (nameCounts.get(k.name) || 0) + 1);

  const seriesColors = palette(Math.max(12, seriesKeys.length));
  for (let i = 0; i < seriesKeys.length; i += 1) {
    const { type, name } = seriesKeys[i];
    const color = seriesColors[i % seriesColors.length];
    const values = trendWeeks.map((w) => {
      if (type === 'l1') {
        const m = parsed.l1SharesByWeek.get(w) || new Map();
        if (name === '其他') {
          let other = 0;
          for (const [k, v] of m.entries()) if (!trendL1Clean.includes(k)) other += v;
          return other;
        }
        return m.get(name) || 0;
      }
      const m = parsed.l2SharesByWeek.get(w) || new Map();
      return m.get(name) || 0;
    });

    // Legend label: keep clean; only disambiguate when names collide across dims
    const label = (nameCounts.get(name) || 0) > 1 ? `${name}（${type === 'l1' ? '一级' : '二级'}）` : name;
    trendSeries.push({
      label,
      data: values,
      borderColor: color,
      backgroundColor: color + '22',
      pointRadius: 0,
      pointHoverRadius: 3,
      tension: 0.25,
      fill: false,
      borderWidth: type === 'l2' ? 1.4 : 1.8,
      borderDash: type === 'l2' ? [4, 3] : undefined,
    });
  }

  const volCanvas = $('wishbarVolumeChart');
  volCanvas.parentElement.style.height = '360px';
  if (volumeChart) volumeChart.destroy();
  volumeChart = new Chart(volCanvas.getContext('2d'), {
    type: 'line',
    data: { labels: trendWeeks, datasets: trendSeries },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      layout: { padding: { top: 10, bottom: 10 } },
      plugins: {
        legend: {
          position: 'bottom',
          align: 'center',
          labels: { usePointStyle: true, boxWidth: 10, padding: 10 },
        },
        tooltip: {
          callbacks: { label: (ctx) => `${ctx.dataset.label}: ${fmtPct(ctx.parsed.y)}` },
        },
      },
      scales: {
        y: {
          min: 0,
          ticks: { callback: (v) => `${v}%` },
          grid: { color: 'rgba(15,23,42,.08)' },
        },
        x: {
          grid: { display: false },
          ticks: { maxRotation: 45, minRotation: 30, autoSkip: true, maxTicksLimit: 10 },
        },
      },
    },
  });

  const volHintEl = $('wishbarVolumeHint');
  if (volHintEl) {
    const numL1 = trendL1Clean.length + (trendShowOther ? 1 : 0);
    const numL2 = trendL2Clean.length;
    const rangeLabel = `${trendWeeks[0]} ~ ${trendWeeks[trendWeeks.length - 1]}`;
    volHintEl.textContent = `已选系列：一级 ${numL1} 条，二级 ${numL2} 条（并集展示；二级用虚线）。日期范围：${rangeLabel}`;
  }
}

// 旧版：基于顶部 TopK 控件。已移除该控件，保留趋势筛选默认逻辑在 initTrendDefaults。

async function onFile(file) {
  setStatus('解析中…');
  try {
    const rows = await parseCsvFile(file);
    parsed = buildWeeklyShares(rows);
    const w = $('wishbarWeekPick');
    if (w) w.value = parsed.weeks[parsed.weeks.length - 1] || '';
    const te = $('wishbarTrendEndOnly');
    if (te && !te.value) te.value = parsed.weeks[parsed.weeks.length - 1] || '';
    render();
  } catch (e) {
    parsed = null;
    setStatus(`解析失败：${e && e.message ? e.message : String(e)}`);
  }
}

async function loadLatestFromBoundFolder() {
  setStatus('读取绑定文件夹中…');
  try {
    const root = await getBoundDirHandle();
    if (!root) throw new Error('尚未绑定数据文件夹。');

    const perm = await root.queryPermission?.({ mode: 'read' });
    if (perm !== 'granted') {
      const req = await root.requestPermission?.({ mode: 'read' });
      if (req !== 'granted') throw new Error('未获得读取权限。');
    }

    const { handle: wishDir, name: subName } = await resolveWishSubdir(root);
    setStatus(`正在扫描「${subName}」下所有 CSV 并合并…`);
    const allRows = [];
    let fileCount = 0;
    // eslint-disable-next-line no-restricted-syntax
    for await (const entry of wishDir.values()) {
      if (!entry || entry.kind !== 'file') continue;
      if (!entry.name.toLowerCase().endsWith('.csv')) continue;
      const file = await entry.getFile();
      const text = await file.text();
      const rows = await parseCsvText(text);
      for (const r of rows) allRows.push(r);
      fileCount += 1;
    }
    if (!fileCount) throw new Error('祈愿子文件夹下未找到任何 CSV 文件。');
    parsed = buildWeeklyShares(allRows);
    const w = $('wishbarWeekPick');
    if (w) w.value = parsed.weeks[parsed.weeks.length - 1] || '';
    const te = $('wishbarTrendEndOnly');
    if (te) te.value = parsed.weeks[parsed.weeks.length - 1] || '';
    render();
    setStatus(`更新完成：已合并「${subName}」下 ${fileCount} 个 CSV。`);
  } catch (e) {
    setStatus(`读取失败：${e && e.message ? e.message : String(e)}`);
  }
}

function init() {
  const input = $('wishbarFileInput');
  input.addEventListener('change', (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) onFile(f);
  });

  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) onFile(f);
  });

  $('wishbarWeekPick')?.addEventListener('change', render);

  $('wishbarTrendEndOnly')?.addEventListener('change', render);

  $('wishbarBarL1Btn')?.addEventListener('click', () => {
    if (!parsed) return;
    const latestWeek = parsed.weeks[parsed.weeks.length - 1];
    const { sources: topSources, foldOther } = pickTopKByLatest(parsed.l1SharesByWeek, latestWeek, parsed.l1All, 8);
    initBarDefaults(parsed, topSources, foldOther);
    const l1Order = buildSourceOrderByLatest(parsed);
    const items = [...l1Order, ...(foldOther ? ['其他'] : [])];
    openFilterPopover({
      anchorEl: $('wishbarBarL1Btn'),
      title: '一级来源筛选（模块1）',
      items,
      selectedSet: barSelectedL1,
      onApply: (next) => {
        barSelectedL1 = next;
        renderBarFilterSummaries();
        render();
      },
    });
  });
  $('wishbarBarL2Btn')?.addEventListener('click', () => {
    if (!parsed) return;
    const items = buildL2OrderByLatest(parsed);
    openFilterPopover({
      anchorEl: $('wishbarBarL2Btn'),
      title: '二级来源筛选（模块1）',
      items,
      selectedSet: barSelectedL2,
      onApply: (next) => {
        barSelectedL2 = next;
        renderBarFilterSummaries();
        render();
      },
    });
  });
  $('wishbarTrendL1Btn')?.addEventListener('click', () => {
    if (!parsed) return;
    const latestWeek = parsed.weeks[parsed.weeks.length - 1];
    const { sources: topSources, foldOther } = pickTopKByLatest(parsed.l1SharesByWeek, latestWeek, parsed.l1All, 8);
    initTrendDefaults(parsed, topSources, foldOther);
    const l1Order = buildSourceOrderByLatest(parsed);
    const items = [...l1Order, ...(foldOther ? ['其他'] : [])];
    openFilterPopover({
      anchorEl: $('wishbarTrendL1Btn'),
      title: '一级来源筛选',
      items,
      selectedSet: trendSelectedL1,
      onApply: (next) => {
        trendSelectedL1 = next;
        renderTrendFilterSummaries();
        render();
      },
    });
  });
  $('wishbarTrendL2Btn')?.addEventListener('click', () => {
    if (!parsed) return;
    const items = buildL2OrderByLatest(parsed);
    openFilterPopover({
      anchorEl: $('wishbarTrendL2Btn'),
      title: '二级来源筛选',
      items,
      selectedSet: trendSelectedL2,
      onApply: (next) => {
        trendSelectedL2 = next;
        renderTrendFilterSummaries();
        render();
      },
    });
  });
  $('wishbarBindFolderBtn')?.addEventListener('click', async () => {
    setStatus('等待选择文件夹…');
    try {
      await pickAndBindFolder();
      setStatus('已绑定。点击「读取祈愿文件夹全部CSV并更新」或继续手动选择文件。');
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      setStatus(`绑定失败：${msg}`);
      // 若在 iframe 内，给出一键打开新窗口的引导
      if (msg.includes('iframe')) {
        try {
          window.open('./wishbar.html', '_blank', 'noopener,noreferrer');
        } catch (_) {}
      }
    }
  });
  $('wishbarUpdateBtn')?.addEventListener('click', loadLatestFromBoundFolder);
  $('wishbarClearBtn').addEventListener('click', () => {
    input.value = '';
    parsed = null;
    trendSelectedL1 = new Set();
    trendSelectedL2 = new Set();
    trendInitialized = false;
    render();
    setStatus('已清空。请选择/拖入 CSV 文件开始分析。');
  });

  // 如果已经绑定过文件夹，给一个提示
  getBoundDirHandle().then((h) => {
    if (h) setStatus('检测到已绑定数据文件夹：可点击「读取祈愿文件夹全部CSV并更新」。');
  });
}

init();

