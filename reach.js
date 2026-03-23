/* global Papa, Chart */

const $ = (id) => document.getElementById(id);

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const DB_NAME = 'ops-dashboard-local-db';
const DB_STORE = 'kv';
const DB_KEY_DIR_HANDLE = 'boundDirHandle';

const SUBDIR_CANDIDATES = [
  '触达',
  '祈愿触达',
  '目标用户触达',
  '祈愿目标用户触达',
  '祈愿目标用户触达看板',
];

const OPS_SUBDIR_CANDIDATES = ['运营宣推', '运营宣推大盘', '运营宣推上'];

const OPS_COLS = {
  gran: '时间维度',
  date: '日期',
  topicName: '专题名称',
  source: '宣发来源',
  placement: '资源位',
  imp: '全局曝光uv-运营宣推',
  clk: '全局点击uv-运营宣推',
  read: '运营宣推贡献-阅读uv',
  rev: '运营宣推贡献-付费解锁收入',
};

let rows = [];
let opsRows = [];
/** 汇总页快照导出：let 不会挂 window */
if (typeof window !== 'undefined') {
  window.__SNAPSHOT_REACH__ = {};
  Object.defineProperty(window.__SNAPSHOT_REACH__, 'rows', {
    configurable: true,
    enumerable: true,
    get() {
      return rows;
    },
  });
  Object.defineProperty(window.__SNAPSHOT_REACH__, 'opsRows', {
    configurable: true,
    enumerable: true,
    get() {
      return opsRows;
    },
  });
}
let chart = null;
let columns = [];
let selectedCategories = new Set(); // empty = all

// Fixed schema (based on user's header screenshot). If a column is missing, we try fuzzy match.
const COL = {
  category: '品类',
  project: '项目名称',
  launchDate: '上线日期',
  onlineDay: '上线天数',
  targetType: '目标用户类型',
  targetSize: '目标用户数',
  reached: '目标用户累计触达',
};

// Reuse the same inline popover style as wishbar
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

  const rect = anchorEl.getBoundingClientRect();
  const gap = 8;
  const vw = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
  const vh = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
  const width = Math.min(720, vw - 24);
  let left = Math.max(12, rect.left);
  left = Math.min(left, vw - width - 12);
  let top = rect.bottom + gap;
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

function summarizeSelection(set, maxItems = 3) {
  const arr = Array.from(set);
  if (arr.length === 0) return '未选择';
  const head = arr.slice(0, maxItems).join('、');
  return arr.length > maxItems ? `${head} 等 ${arr.length} 项` : `${head}（${arr.length}）`;
}

function ensureDateRangePopover() {
  let pop = document.getElementById('dateRangePopover');
  if (pop) return pop;
  pop = document.createElement('div');
  pop.id = 'dateRangePopover';
  pop.className = 'popover';
  pop.innerHTML = `
    <div class="popover__head">
      <div class="popover__title">上线日期范围</div>
      <button class="btn btn--ghost btn--sm" id="datePopClose" type="button">关闭</button>
    </div>
    <div class="popover__body" id="datePopBody"></div>
    <div class="popover__footer">
      <button class="btn btn--ghost" id="datePopClear" type="button">清空</button>
      <button class="btn btn--ghost" id="datePopCancel" type="button">取消</button>
      <button class="btn" id="datePopOk" type="button">确定</button>
    </div>
  `;
  document.body.appendChild(pop);
  return pop;
}

function openLaunchRangePopover(anchorEl, onApply) {
  const pop = ensureDateRangePopover();
  const closeBtn = document.getElementById('datePopClose');
  const cancelBtn = document.getElementById('datePopCancel');
  const okBtn = document.getElementById('datePopOk');
  const clearBtn = document.getElementById('datePopClear');
  const bodyEl = document.getElementById('datePopBody');
  const start0 = $('reachLaunchStart')?.value || '';
  const end0 = $('reachLaunchEnd')?.value || '';

  bodyEl.innerHTML = `
    <div class="filterRow" style="margin-bottom:10px">
      <div class="control" style="min-width:220px">
        <div class="control__label">开始日期</div>
        <input id="tmpLaunchStart" type="date" value="${start0}" />
      </div>
      <div class="control" style="min-width:220px">
        <div class="control__label">结束日期</div>
        <input id="tmpLaunchEnd" type="date" value="${end0}" />
      </div>
    </div>
    <div class="muted" style="font-size:12px">留空表示不筛选。</div>
  `;

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

  closeBtn.onclick = close;
  cancelBtn.onclick = close;
  clearBtn.onclick = () => {
    const sEl = document.getElementById('tmpLaunchStart');
    const eEl = document.getElementById('tmpLaunchEnd');
    if (sEl) sEl.value = '';
    if (eEl) eEl.value = '';
  };
  okBtn.onclick = () => {
    const s = document.getElementById('tmpLaunchStart')?.value || '';
    const e = document.getElementById('tmpLaunchEnd')?.value || '';
    onApply({ start: s, end: e });
    close();
  };

  const rect = anchorEl.getBoundingClientRect();
  const gap = 8;
  const vw = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
  const vh = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
  const width = Math.min(520, vw - 24);
  let left = Math.max(12, rect.left);
  left = Math.min(left, vw - width - 12);
  let top = rect.bottom + gap;
  const estH = 260;
  if (top + estH > vh - 12) top = Math.max(12, rect.top - gap - estH);
  pop.style.width = `${width}px`;
  pop.style.left = `${left}px`;
  pop.style.top = `${top}px`;
  pop.style.display = 'block';
  setTimeout(() => document.getElementById('tmpLaunchStart')?.focus(), 0);
  document.addEventListener('mousedown', onOutsideDown, true);
  document.addEventListener('keydown', onKey, true);
}

function num(v) {
  if (v == null) return 0;
  const s = String(v).trim();
  if (!s) return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function setStatus(msg) {
  const el = $('reachStatus');
  if (el) el.textContent = msg;
}

function parseDate(s) {
  const v = (s == null ? '' : String(s)).trim();
  if (!v) return null;
  const d = new Date(v);
  // Support yyyy-mm-dd
  if (!Number.isNaN(d.getTime())) return d;
  return null;
}

function daysBetween(a, b) {
  return Math.floor((b.getTime() - a.getTime()) / 86400000);
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

async function pickAndBindFolder() {
  if (!window.showDirectoryPicker) {
    throw new Error('当前浏览器不支持“绑定文件夹”（需要 Chrome/Edge 的 File System Access API）。');
  }
  if (window.top && window.top !== window.self) {
    throw new Error('当前页面在汇总页(iframe)内，浏览器可能拦截目录选择器。请新窗口打开本页绑定。');
  }
  const handle = await window.showDirectoryPicker({ mode: 'read' });
  await dbSet(DB_KEY_DIR_HANDLE, handle);
  return handle;
}

async function resolveSubdir(rootHandle) {
  for (const name of SUBDIR_CANDIDATES) {
    try {
      const h = await rootHandle.getDirectoryHandle(name, { create: false });
      return { handle: h, name };
    } catch (_) {}
  }
  throw new Error(`未找到触达数据子文件夹：请在绑定目录下创建 ${SUBDIR_CANDIDATES.map((x) => `「${x}」`).join(' 或 ')}。`);
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

async function loadAllFromBoundFolder() {
  setStatus('读取绑定文件夹中…');
  const root = await getBoundDirHandle();
  if (!root) throw new Error('尚未绑定数据文件夹。');
  const perm = await root.queryPermission?.({ mode: 'read' });
  if (perm !== 'granted') {
    const req = await root.requestPermission?.({ mode: 'read' });
    if (req !== 'granted') throw new Error('未获得读取权限。');
  }
  const { handle: dir, name } = await resolveSubdir(root);
  setStatus(`正在扫描「${name}」下所有 CSV 并合并…`);
  const all = [];
  let fileCount = 0;
  // eslint-disable-next-line no-restricted-syntax
  for await (const entry of dir.values()) {
    if (!entry || entry.kind !== 'file') continue;
    if (!entry.name.toLowerCase().endsWith('.csv')) continue;
    const file = await entry.getFile();
    const text = await file.text();
    const r = await parseCsvText(text);
    for (const row of r) all.push(row);
    fileCount += 1;
  }
  if (!fileCount) throw new Error('触达子文件夹下未找到任何 CSV 文件。');
  rows = all;
  setStatus(`更新完成：已合并 ${fileCount} 个 CSV（${rows.length} 行）。请检查字段映射后查看图表。`);
  hydrateColumns();
  await loadOpsDataSilently();
  render();
}

async function resolveOpsSubdir(rootHandle) {
  for (const name of OPS_SUBDIR_CANDIDATES) {
    try {
      const h = await rootHandle.getDirectoryHandle(name, { create: false });
      return { handle: h, name };
    } catch (_) {}
  }
  return null;
}

async function loadOpsDataSilently() {
  try {
    const root = await getBoundDirHandle();
    if (!root) return;
    const perm = await root.queryPermission?.({ mode: 'read' });
    if (perm !== 'granted') return;
    const sub = await resolveOpsSubdir(root);
    if (!sub) {
      opsRows = [];
      return;
    }
    const all = [];
    for await (const entry of sub.handle.values()) {
      if (!entry || entry.kind !== 'file') continue;
      if (!entry.name.toLowerCase().endsWith('.csv')) continue;
      const file = await entry.getFile();
      const text = await file.text();
      const r = await parseCsvText(text);
      for (const row of r) all.push(row);
    }
    opsRows = all;
  } catch (_) {
    opsRows = [];
  }
}

function opsRowIsDailyRow(row) {
  const g = String(row[OPS_COLS.gran] ?? row[`\ufeff${OPS_COLS.gran}`] ?? '').trim();
  if (!g) return true;
  return g === '日';
}

function buildOpsWishMap() {
  if (!opsRows.length) return new Map();
  const daily = opsRows.filter(opsRowIsDailyRow);
  const wishOnly = daily.filter((r) => String(r[OPS_COLS.topicName] ?? '').trim().includes('祈愿'));
  if (!wishOnly.length) return new Map();

  const allDates = Array.from(new Set(wishOnly.map((r) => String(r[OPS_COLS.date] ?? '').trim()).filter(Boolean))).sort();
  if (!allDates.length) return new Map();

  const maxDate = allDates[allDates.length - 1];
  const weekStart = allDates[Math.max(0, allDates.length - 7)];
  const dayCount = allDates.filter((d) => d >= weekStart && d <= maxDate).length || 1;

  const weekRows = wishOnly.filter((r) => {
    const d = String(r[OPS_COLS.date] ?? '').trim();
    return d >= weekStart && d <= maxDate;
  });

  const byTopic = new Map();
  for (const r of weekRows) {
    const name = String(r[OPS_COLS.topicName] ?? '').trim() || '(空)';
    const cur = byTopic.get(name) || { name, imp: 0, clk: 0, rev: 0, dayCount };
    cur.imp += num(r[OPS_COLS.imp]);
    cur.clk += num(r[OPS_COLS.clk]);
    cur.rev += num(r[OPS_COLS.rev]);
    byTopic.set(name, cur);
  }
  return byTopic;
}

function buildOpsCtrBenchmark() {
  if (!opsRows.length) return null;
  const daily = opsRows.filter(opsRowIsDailyRow);
  const wishOnly = daily.filter((r) => String(r[OPS_COLS.topicName] ?? '').trim().includes('祈愿'));
  if (!wishOnly.length) return null;

  const allDates = Array.from(new Set(wishOnly.map((r) => String(r[OPS_COLS.date] ?? '').trim()).filter(Boolean))).sort();
  if (!allDates.length) return null;

  const maxDate = allDates[allDates.length - 1];
  const d30Start = allDates[Math.max(0, allDates.length - 30)];

  const recent = wishOnly.filter((r) => {
    const d = String(r[OPS_COLS.date] ?? '').trim();
    return d >= d30Start && d <= maxDate;
  });

  const byTopic = new Map();
  for (const r of recent) {
    const name = String(r[OPS_COLS.topicName] ?? '').trim() || '(空)';
    const cur = byTopic.get(name) || { imp: 0, clk: 0 };
    cur.imp += num(r[OPS_COLS.imp]);
    cur.clk += num(r[OPS_COLS.clk]);
    byTopic.set(name, cur);
  }

  const MIN_IMP = 1000;
  const ctrs = [];
  for (const [, v] of byTopic) {
    if (v.imp < MIN_IMP) continue;
    ctrs.push(v.clk / v.imp);
  }
  if (ctrs.length < 2) return null;

  const clean = iqrFilter(ctrs);
  const sorted = clean.slice().sort((a, b) => a - b);
  return {
    n: ctrs.length,
    nClean: clean.length,
    median: median(clean),
    sorted,
    dateRange: `${d30Start} ~ ${maxDate}`,
  };
}

function judgeCtrLevel(ctr, bench) {
  if (!bench || ctr == null) return null;
  const pct = percentileRank(ctr, bench.sorted);
  let label, color;
  if (pct >= 75) { label = '优秀'; color = '#16a34a'; }
  else if (pct >= 50) { label = '良好'; color = '#059669'; }
  else if (pct >= 25) { label = '一般'; color = '#d97706'; }
  else { label = '偏低'; color = '#dc2626'; }
  const fmtR = (v) => `${(v * 100).toFixed(2)}%`;
  const detail = `近30天祈愿P50=${fmtR(bench.median)}（${bench.nClean}样本去异常）· P${Math.round(pct)}`;
  return { label, color, detail };
}

function matchOpsProject(reachName, opsMap) {
  if (opsMap.has(reachName)) return opsMap.get(reachName);
  for (const [opsName, data] of opsMap) {
    if (reachName.includes(opsName) || opsName.includes(reachName)) return data;
  }
  const reachTokens = reachName.replace(/[-_.·,，]/g, ' ').split(/\s+/).filter((t) => t.length >= 2);
  for (const [opsName, data] of opsMap) {
    const matched = reachTokens.filter((t) => opsName.includes(t)).length;
    if (matched >= 2 || (reachTokens.length === 1 && matched === 1)) return data;
  }
  return null;
}

function hydrateColumns() {
  columns = rows.length ? Object.keys(rows[0]) : [];
  initEvalDayOptions();
  initTargetTypeOptionsFromData();
  applyReachQueryDefaults();
  initBucketPresets();
  initCategoryOptions();
}

function initEvalDayOptions() {
  const sel = $('reachEvalDay');
  if (!sel) return;
  const opts = [
    { v: '30', t: 'D30（上线30日内）' },
    { v: '14', t: 'D14' },
    { v: '7', t: 'D7' },
    { v: '0', t: 'D0' },
  ];
  sel.innerHTML = opts.map((o) => `<option value="${o.v}">${o.t}</option>`).join('');
  sel.value = '30';
}

function initTargetTypeOptionsFromData() {
  const sel = $('reachTargetType');
  if (!sel) return;
  const col = resolveCol(COL.targetType);
  const vals = col ? Array.from(new Set(rows.map((r) => String(r[col] ?? '').trim()).filter(Boolean))).sort() : [];
  const fallback = ['整体', '阅读', '单点'];
  const options = vals.length ? vals : fallback;
  sel.innerHTML = ['(全部)', ...options].map((t) => `<option value="${t}">${t}</option>`).join('');
  // 数据口径默认：有「整体」则选整体，否则选「(全部)」避免筛出空集
  sel.value = options.includes('整体') ? '整体' : '(全部)';
}

/** URL 覆盖活动标识：?activity=整体 或 ?targetType=整体（须在选项中存在） */
function applyReachQueryDefaults() {
  const sel = $('reachTargetType');
  if (!sel || !sel.options.length) return;
  const params = new URLSearchParams(window.location.search);
  const want = (params.get('activity') || params.get('targetType') || '').trim();
  if (!want) return;
  const ok = Array.from(sel.options).some((o) => o.value === want);
  if (ok) sel.value = want;
}

function initBucketPresets() {
  const sel = $('reachBucketPreset');
  if (!sel) return;
  const presets = [
    { v: '(全部)', t: '不分档（全部规模）' },
    { v: '0,10000,50000,100000,300000,1000000,inf', t: '0-1万-5万-10万-30万-100万-100万+' },
    { v: '0,5000,20000,50000,100000,300000,inf', t: '0-5千-2万-5万-10万-30万-30万+' },
    { v: '0,20000,100000,300000,1000000,inf', t: '0-2万-10万-30万-100万-100万+' },
  ];
  sel.innerHTML = presets.map((p) => `<option value="${p.v}">${p.t}</option>`).join('');
  sel.value = '(全部)';
}

function parseBucketsFromPreset() {
  const raw = ($('reachBucketPreset')?.value || '').trim();
  if (!raw || raw === '(全部)') return null;
  const parts = raw.split(',').map((x) => x.trim()).filter(Boolean);
  const vals = parts
    .map((p) => (p.toLowerCase() === 'inf' ? Infinity : Number(p)))
    .filter((n) => Number.isFinite(n) || n === Infinity);
  vals.sort((a, b) => a - b);
  if (vals.length === 0 || vals[0] !== 0) vals.unshift(0);
  if (vals[vals.length - 1] !== Infinity) vals.push(Infinity);
  return vals;
}

function bucketLabel(lo, hi) {
  const fmt = (n) => {
    if (n === Infinity) return '∞';
    if (n >= 1e6) return `${Math.round(n / 1e6)}00万+`.replace('1000万+', '1000万+');
    if (n >= 1e4) return `${Math.round(n / 1e4)}万`;
    return String(n);
  };
  if (hi === Infinity) return `${fmt(lo)}+`;
  return `${fmt(lo)}-${fmt(hi)}`;
}

function initCategoryOptions() {
  const btn = $('reachCatBtn');
  const sum = $('reachCatSummary');
  if (!btn || !sum) return;
  // defer; populated when opening popover
  sum.textContent = selectedCategories.size ? summarizeSelection(selectedCategories, 3) : '全部';
}

function resolveCol(preferName) {
  if (columns.includes(preferName)) return preferName;
  // fuzzy contains match
  const hits = columns.filter((c) => c.includes(preferName));
  if (hits.length) return hits[0];
  // keyword fallback
  const kw = {
    [COL.category]: ['二级品类', '品类'],
    [COL.project]: ['项目名称', '项目名', '专题名称'],
    [COL.launchDate]: ['上线日期', '上线'],
    [COL.onlineDay]: ['上线天数', '天数'],
    [COL.targetType]: ['目标用户类型', '二级目标用户类型'],
    [COL.targetSize]: ['目标用户数', '目标用户规模'],
    [COL.reached]: ['目标用户累计触达', '累计触达', '触达用户'],
  }[preferName] || [];
  for (const k of kw) {
    const c = columns.find((x) => x.includes(k));
    if (c) return c;
  }
  return '';
}

function compute() {
  const colCategory = resolveCol(COL.category);
  const colProject = resolveCol(COL.project);
  const colLaunch = resolveCol(COL.launchDate);
  const colDay = resolveCol(COL.onlineDay);
  const colTargetType = resolveCol(COL.targetType);
  const colTargetSize = resolveCol(COL.targetSize);
  const colReached = resolveCol(COL.reached);

  if (![colCategory, colProject, colLaunch, colDay, colTargetType, colTargetSize, colReached].every((x) => x)) {
    return { ok: false, reason: 'CSV缺少必要字段，请确认包含：品类、项目名称、上线日期、上线天数、目标用户类型、目标用户数、目标用户累计触达。' };
  }

  const buckets = parseBucketsFromPreset(); // null means no bucket filter
  const evalDay = Math.max(0, Math.min(30, Number($('reachEvalDay')?.value || 30)));
  const targetTypeWanted = String($('reachTargetType')?.value || '整体').trim();
  const q = String($('reachProjectQ')?.value || '').trim().toLowerCase();
  const launchStart = parseDate($('reachLaunchStart')?.value);
  const launchEnd = parseDate($('reachLaunchEnd')?.value);
  const topN = Number($('reachTopN')?.value || 12);

  // per project timeseries (0..evalDay)
  const proj = new Map(); // key -> {category, project, launchDateStr, size, byDay: Map(day->reachedMax)}
  for (const r of rows) {
    const targetType = String(r[colTargetType] ?? '').trim();
    if (targetTypeWanted && targetTypeWanted !== '(全部)' && targetType !== targetTypeWanted) continue;
    const launch = parseDate(r[colLaunch]);
    if (!launch) continue;
    if (launchStart && launch < launchStart) continue;
    if (launchEnd && launch > launchEnd) continue;

    const day = num(r[colDay]);
    if (day < 0 || day > evalDay) continue;

    const category = String(r[colCategory] ?? '').trim() || '(空)';
    const project = String(r[colProject] ?? '').trim() || '(空)';
    if (selectedCategories.size && !selectedCategories.has(category)) continue;
    if (q && !project.toLowerCase().includes(q)) continue;

    const size = num(r[colTargetSize]);
    const reached = num(r[colReached]);
    const key = `${category}||${project}`;
    const cur = proj.get(key) || {
      category,
      project,
      launchDateStr: String(r[colLaunch] ?? '').trim(),
      size: 0,
      byDay: new Map(),
    };
    cur.size = Math.max(cur.size || 0, size);
    const prev = cur.byDay.get(day) || 0;
    cur.byDay.set(day, Math.max(prev, reached));
    proj.set(key, cur);
  }

  let items = Array.from(proj.values()).filter((x) => x.size > 0);

  // bucket filter (optional)
  if (buckets) {
    const defs = [];
    for (let i = 0; i < buckets.length - 1; i += 1) defs.push([buckets[i], buckets[i + 1]]);
    const pick = (n) => {
      for (const [lo, hi] of defs) if (n >= lo && n < hi) return bucketLabel(lo, hi);
      return bucketLabel(defs[defs.length - 1][0], defs[defs.length - 1][1]);
    };
    const allowedBucket = new Set(defs.map((d) => bucketLabel(d[0], d[1])));
    // keep all buckets; filter is actually all, so no-op. (Preset dropdown already chooses defs)
    items = items.filter((x) => allowedBucket.has(pick(x.size)));
  }

  // compute D{evalDay} reachRate for ranking and details
  const withD = items.map((x) => {
    let maxReached = 0;
    for (let d = 0; d <= evalDay; d += 1) {
      const v = x.byDay.get(d) || 0;
      if (v > maxReached) maxReached = v;
    }
    return { ...x, reachRateAtDay: maxReached / x.size };
  });

  // rank: by target size desc then reach rate desc
  withD.sort((a, b) => (b.size - a.size) || (b.reachRateAtDay - a.reachRateAtDay));
  const picked = topN >= 999 ? withD : withD.slice(0, topN);

  // build chart series (carry-forward)
  const days = Array.from({ length: evalDay + 1 }, (_, i) => i);
  const series = picked.map((x) => {
    let last = 0;
    const data = days.map((d) => {
      const v = x.byDay.get(d);
      if (Number.isFinite(v)) last = Math.max(last, v);
      return (last / x.size) * 100;
    });
    return { ...x, days, seriesPct: data };
  });

  // details grouped by category
  const byCat = new Map();
  for (const it of withD) {
    if (!byCat.has(it.category)) byCat.set(it.category, []);
    byCat.get(it.category).push(it);
  }
  for (const list of byCat.values()) list.sort((a, b) => (b.size - a.size) || (b.reachRateAtDay - a.reachRateAtDay));

  return { ok: true, evalDay, days, series, byCat };
}

function extractBaseProject(projectName) {
  const parts = projectName.split('.');
  return parts[0].trim();
}

function iqrFilter(arr) {
  if (arr.length < 4) return arr.slice();
  const sorted = arr.slice().sort((a, b) => a - b);
  const q1 = sorted[Math.floor(sorted.length * 0.25)];
  const q3 = sorted[Math.floor(sorted.length * 0.75)];
  const iqr = q3 - q1;
  const lo = q1 - 1.5 * iqr;
  const hi = q3 + 1.5 * iqr;
  return sorted.filter((v) => v >= lo && v <= hi);
}

function median(arr) {
  if (!arr.length) return 0;
  const s = arr.slice().sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

function percentileRank(value, sortedArr) {
  if (!sortedArr.length) return 50;
  let below = 0;
  for (const v of sortedArr) { if (v < value) below++; else break; }
  return (below / sortedArr.length) * 100;
}

function buildHistoricalBenchmarks(colCategory, targetTypeWanted) {
  const colProject = resolveCol(COL.project);
  const colLaunch = resolveCol(COL.launchDate);
  const colDay = resolveCol(COL.onlineDay);
  const colTargetType = resolveCol(COL.targetType);
  const colTargetSize = resolveCol(COL.targetSize);
  const colReached = resolveCol(COL.reached);

  const allProj = new Map();
  for (const r of rows) {
    const tt = String(r[colTargetType] ?? '').trim();
    if (targetTypeWanted && targetTypeWanted !== '(全部)' && tt !== targetTypeWanted) continue;
    const project = String(r[colProject] ?? '').trim();
    const launch = parseDate(String(r[colLaunch] ?? '').trim());
    if (!launch || !project) continue;
    const day = num(r[colDay]);
    const size = num(r[colTargetSize]);
    const reached = num(r[colReached]);
    const category = String(r[colCategory] ?? '').trim() || '(空)';
    const cur = allProj.get(project) || { project, category, size: 0, byDay: new Map() };
    cur.size = Math.max(cur.size, size);
    const prev = cur.byDay.get(day) || 0;
    cur.byDay.set(day, Math.max(prev, reached));
    allProj.set(project, cur);
  }

  function rateAtDay(proj, d) {
    if (proj.size <= 0) return null;
    let maxR = 0;
    let hasData = false;
    for (let i = 0; i <= d; i++) {
      const v = proj.byDay.get(i);
      if (v != null) { hasData = true; if (v > maxR) maxR = v; }
    }
    return hasData ? maxR / proj.size : null;
  }

  function getCategoryBenchmark(category, atDay, excludeProject) {
    const rates = [];
    for (const p of allProj.values()) {
      if (p.category !== category) continue;
      if (excludeProject && p.project === excludeProject) continue;
      const r = rateAtDay(p, atDay);
      if (r != null) rates.push(r);
    }
    const clean = iqrFilter(rates);
    const sorted = clean.slice().sort((a, b) => a - b);
    return {
      n: rates.length,
      nClean: clean.length,
      median: median(clean),
      mean: clean.length ? clean.reduce((s, v) => s + v, 0) / clean.length : null,
      sorted,
    };
  }

  function getProjectHistory(baseProject, excludeName, atDay) {
    const hist = [];
    for (const p of allProj.values()) {
      if (p.project === excludeName) continue;
      if (extractBaseProject(p.project) !== baseProject) continue;
      const r = rateAtDay(p, atDay);
      if (r != null) hist.push({ project: p.project, rate: r });
    }
    return hist;
  }

  return { getCategoryBenchmark, getProjectHistory };
}

function judgeReachLevel(currentRate, catBench, projHistory, maxDay) {
  const pct = (v) => `${(v * 100).toFixed(2)}%`;
  const hasHistory = projHistory.length > 0;
  const histMedian = hasHistory ? median(projHistory.map((h) => h.rate)) : null;

  let catPctRank = null;
  let catLabel = '';
  if (catBench.sorted.length >= 3) {
    catPctRank = percentileRank(currentRate, catBench.sorted);
    catLabel = `品类P50=${pct(catBench.median)}（去异常值后${catBench.nClean}个样本）· 当前排P${Math.round(catPctRank)}`;
  } else if (catBench.mean != null) {
    catLabel = `品类中位数=${pct(catBench.median)}（${catBench.n}个样本，不足去异常值）`;
  }

  let histLabel = '';
  if (hasHistory) {
    const histRates = projHistory.map((h) => `${pct(h.rate)}`).join('、');
    histLabel = `同项目往期${projHistory.length}期中位数=${pct(histMedian)}（${histRates}）`;
  } else {
    histLabel = '无同项目往期特典';
  }

  let score = 0;
  let factors = 0;

  if (catPctRank != null) {
    if (catPctRank >= 75) score += 2;
    else if (catPctRank >= 50) score += 1;
    else if (catPctRank >= 25) score -= 1;
    else score -= 2;
    factors++;
  } else if (catBench.median > 0) {
    const ratio = currentRate / catBench.median;
    if (ratio >= 1.15) score += 1.5;
    else if (ratio >= 0.85) score += 0;
    else score -= 1.5;
    factors++;
  }

  if (hasHistory && histMedian > 0) {
    const ratio = currentRate / histMedian;
    if (ratio >= 1.15) score += 2;
    else if (ratio >= 0.95) score += 0.5;
    else if (ratio >= 0.85) score -= 0.5;
    else if (ratio >= 0.7) score -= 1.5;
    else score -= 2;
    factors++;
  }

  const avg = factors > 0 ? score / factors : 0;

  let label, color, suggestion;
  if (avg >= 1.2) {
    label = '优秀';
    color = '#16a34a';
    suggestion = '触达势头强劲，可考虑延长或加码推广周期';
  } else if (avg >= 0.3) {
    label = '良好';
    color = '#059669';
    suggestion = '表现正常偏上，维持当前资源投入';
  } else if (avg >= -0.3) {
    label = '一般';
    color = '#d97706';
    suggestion = '触达中规中矩，可尝试优化素材或调整资源位';
  } else if (avg >= -1.2) {
    label = '偏低';
    color = '#ea580c';
    suggestion = '低于预期，建议排查素材吸引力或投放定向';
  } else {
    label = '待提升';
    color = '#dc2626';
    suggestion = '显著低于同类，建议复盘策略并考虑追加曝光';
  }

  if (maxDay <= 2) {
    suggestion = `(D${maxDay}数据较早) ` + suggestion;
  }

  const detail = [catLabel, histLabel].filter(Boolean).join('；');

  return { label, color, suggestion, detail };
}

// ─── Deep analysis per project ───

function analyzeTargetTypes(projectName) {
  const colProject = resolveCol(COL.project);
  const colDay = resolveCol(COL.onlineDay);
  const colTargetType = resolveCol(COL.targetType);
  const colTargetSize = resolveCol(COL.targetSize);
  const colReached = resolveCol(COL.reached);
  if (!colProject || !colTargetType) return [];

  const byType = new Map();
  for (const r of rows) {
    const proj = String(r[colProject] ?? '').trim();
    if (proj !== projectName) continue;
    const tt = String(r[colTargetType] ?? '').trim() || '(空)';
    const day = num(r[colDay]);
    const size = num(r[colTargetSize]);
    const reached = num(r[colReached]);
    const cur = byType.get(tt) || { type: tt, size: 0, maxDay: -1, reachedAtMax: 0 };
    cur.size = Math.max(cur.size, size);
    if (day > cur.maxDay) { cur.maxDay = day; cur.reachedAtMax = reached; }
    else if (day === cur.maxDay) cur.reachedAtMax = Math.max(cur.reachedAtMax, reached);
    byType.set(tt, cur);
  }
  return Array.from(byType.values())
    .filter((x) => x.size > 0)
    .map((x) => ({ ...x, rate: x.reachedAtMax / x.size }))
    .sort((a, b) => a.rate - b.rate);
}

function analyzePositions(projectName, sourceFilter) {
  if (!opsRows.length) return [];
  const daily = opsRows.filter(opsRowIsDailyRow);
  let projRows = daily.filter((r) => {
    const name = String(r[OPS_COLS.topicName] ?? '').trim();
    return name === projectName || name.includes(projectName) || projectName.includes(name);
  });
  if (sourceFilter) {
    projRows = projRows.filter((r) => (String(r[OPS_COLS.source] ?? '').trim() || '(空)') === sourceFilter);
  }
  if (!projRows.length) {
    const base = extractBaseProject(projectName);
    const fuzzy = daily.filter((r) => {
      const n = String(r[OPS_COLS.topicName] ?? '').trim();
      return n.includes(base) && n.includes('祈愿');
    });
    if (!fuzzy.length) return [];
    let fuzzyFiltered = fuzzy;
    if (sourceFilter) fuzzyFiltered = fuzzy.filter((r) => (String(r[OPS_COLS.source] ?? '').trim() || '(空)') === sourceFilter);
    return buildPositionStats(fuzzyFiltered.length ? fuzzyFiltered : fuzzy);
  }
  return buildPositionStats(projRows);
}

function buildPositionStats(projRows) {
  const allDates = Array.from(new Set(projRows.map((r) => String(r[OPS_COLS.date] ?? '').trim()).filter(Boolean))).sort();
  const maxDate = allDates[allDates.length - 1] || '';
  const weekStart = allDates[Math.max(0, allDates.length - 7)] || maxDate;

  const sorted = projRows
    .filter((r) => { const d = String(r[OPS_COLS.date] ?? '').trim(); return d >= weekStart && d <= maxDate; })
    .sort((a, b) => String(a[OPS_COLS.date] ?? '').localeCompare(String(b[OPS_COLS.date] ?? '')));

  const excludePlacement = new Set(['其它', '其他', '(空)']);
  const byPos = new Map();
  for (const r of sorted) {
    const pos = String(r[OPS_COLS.placement] ?? '').trim() || '(空)';
    if (excludePlacement.has(pos)) continue;
    const cur = byPos.get(pos) || { pos, imp: 0, clk: 0, rev: 0, dailyCtr: [] };
    const dayImp = num(r[OPS_COLS.imp]);
    const dayClk = num(r[OPS_COLS.clk]);
    cur.imp += dayImp;
    cur.clk += dayClk;
    cur.rev += num(r[OPS_COLS.rev]);
    if (dayImp > 100) cur.dailyCtr.push(dayClk / dayImp);
    byPos.set(pos, cur);
  }

  return Array.from(byPos.values())
    .filter((x) => x.imp > 0)
    .map((x) => {
      const ctr = x.clk / x.imp;
      const erpi = x.rev / x.imp;
      let trend = 'stable';
      if (x.dailyCtr.length >= 4) {
        const first = x.dailyCtr.slice(0, Math.ceil(x.dailyCtr.length / 2));
        const second = x.dailyCtr.slice(Math.ceil(x.dailyCtr.length / 2));
        const avgFirst = first.reduce((s, v) => s + v, 0) / first.length;
        const avgSecond = second.reduce((s, v) => s + v, 0) / second.length;
        if (avgFirst > 0 && avgSecond < avgFirst * 0.8) trend = 'declining';
        else if (avgFirst > 0 && avgSecond > avgFirst * 1.2) trend = 'rising';
      }
      return { ...x, ctr, erpi, trend };
    })
    .sort((a, b) => b.erpi - a.erpi);
}

function buildCrossProjectPositionMatrix(wishProjectNames, sourceFilter) {
  if (!opsRows.length) return { matrix: [], swapSuggestions: [], saturationAlerts: [] };
  const daily = opsRows.filter(opsRowIsDailyRow);
  const allDates = Array.from(new Set(daily.map((r) => String(r[OPS_COLS.date] ?? '').trim()).filter(Boolean))).sort();
  const maxDate = allDates[allDates.length - 1] || '';
  const weekStart = allDates[Math.max(0, allDates.length - 7)] || maxDate;

  let recent = daily.filter((r) => {
    const d = String(r[OPS_COLS.date] ?? '').trim();
    return d >= weekStart && d <= maxDate;
  });

  if (sourceFilter) {
    recent = recent.filter((r) => (String(r[OPS_COLS.source] ?? '').trim() || '(空)') === sourceFilter);
  }

  const wishRecent = recent.filter((r) => String(r[OPS_COLS.topicName] ?? '').trim().includes('祈愿'));

  const excludePlacement = new Set(['其它', '其他', '(空)']);
  const matrix = new Map();
  const allPositions = new Set();

  for (const r of wishRecent) {
    const pos = String(r[OPS_COLS.placement] ?? '').trim() || '(空)';
    if (excludePlacement.has(pos)) continue;
    const topic = String(r[OPS_COLS.topicName] ?? '').trim();
    allPositions.add(pos);
    const key = `${topic}||${pos}`;
    const cur = matrix.get(key) || { topic, pos, imp: 0, clk: 0, rev: 0 };
    cur.imp += num(r[OPS_COLS.imp]);
    cur.clk += num(r[OPS_COLS.clk]);
    cur.rev += num(r[OPS_COLS.rev]);
    matrix.set(key, cur);
  }

  const topicSet = new Set();
  for (const v of matrix.values()) topicSet.add(v.topic);
  const topics = Array.from(topicSet);

  const getMetric = (topic, pos) => {
    const v = matrix.get(`${topic}||${pos}`);
    if (!v || v.imp < 500) return null;
    return { ctr: v.clk / v.imp, erpi: v.rev / v.imp, imp: v.imp, rev: v.rev };
  };

  const swapSuggestions = [];
  const positions = Array.from(allPositions);

  for (let i = 0; i < topics.length; i++) {
    for (let j = i + 1; j < topics.length; j++) {
      for (let pi = 0; pi < positions.length; pi++) {
        for (let pj = pi + 1; pj < positions.length; pj++) {
          const a1 = getMetric(topics[i], positions[pi]);
          const a2 = getMetric(topics[i], positions[pj]);
          const b1 = getMetric(topics[j], positions[pi]);
          const b2 = getMetric(topics[j], positions[pj]);
          if (!a1 || !a2 || !b1 || !b2) continue;
          const currentRev = a1.rev + b2.rev;
          const swappedRev = a2.erpi * a1.imp + b1.erpi * b2.imp;
          if (currentRev <= 0) continue;
          const gainPct = (swappedRev - currentRev) / currentRev;
          if (gainPct > 0.1 && a2.erpi > a1.erpi && b1.erpi > b2.erpi) {
            swapSuggestions.push({
              projA: topics[i], projB: topics[j],
              posA: positions[pi], posB: positions[pj],
              gainPct,
              detail: `${topics[i]} 在「${positions[pj]}」ERPI(${a2.erpi.toFixed(4)}) > 「${positions[pi]}」(${a1.erpi.toFixed(4)})；${topics[j]} 反之(${b1.erpi.toFixed(4)} > ${b2.erpi.toFixed(4)})。按当前曝光量估算，调换后综合收入约提升 ${(gainPct * 100).toFixed(0)}%`,
            });
          }
        }
      }
    }
  }
  swapSuggestions.sort((a, b) => b.gainPct - a.gainPct);

  const saturationAlerts = [];
  for (const [, v] of matrix) {
    if (v.imp < 5000) continue;
    const projDailyRows = wishRecent.filter((r) =>
      String(r[OPS_COLS.topicName] ?? '').trim() === v.topic &&
      String(r[OPS_COLS.placement] ?? '').trim() === (v.pos || '(空)')
    ).sort((a, b) => String(a[OPS_COLS.date] ?? '').localeCompare(String(b[OPS_COLS.date] ?? '')));

    if (projDailyRows.length < 4) continue;
    const points = [];
    for (const r of projDailyRows) {
      const di = num(r[OPS_COLS.imp]);
      const dc = num(r[OPS_COLS.clk]);
      if (di > 200) points.push({ ctr: dc / di });
    }
    if (points.length < 4) continue;
    const firstHalf = points.slice(0, Math.ceil(points.length / 2));
    const secondHalf = points.slice(Math.ceil(points.length / 2));
    const avgFirst = firstHalf.reduce((s, p) => s + p.ctr, 0) / firstHalf.length;
    const avgSecond = secondHalf.reduce((s, p) => s + p.ctr, 0) / secondHalf.length;
    if (avgFirst > 0 && avgSecond < avgFirst * 0.75) {
      const dropPct = ((1 - avgSecond / avgFirst) * 100).toFixed(0);
      saturationAlerts.push({
        topic: v.topic, pos: v.pos, cumImp: v.imp, dropPct,
        detail: `${v.topic} 在「${v.pos}」累计曝光 ${Math.round(v.imp).toLocaleString()} 后，后半周CTR均值较前半周下降 ${dropPct}%（前${firstHalf.length}天avg=${(avgFirst * 100).toFixed(2)}% → 后${secondHalf.length}天avg=${(avgSecond * 100).toFixed(2)}%），可能接近饱和`,
      });
    }
  }
  saturationAlerts.sort((a, b) => Number(b.dropPct) - Number(a.dropPct));

  return { matrix: Array.from(matrix.values()), swapSuggestions, saturationAlerts };
}

function generateProjectSuggestions(p, targetTypes, positions) {
  const tips = [];
  const fR = (v) => `${(v * 100).toFixed(2)}%`;

  if (targetTypes.length > 1) {
    const overall = targetTypes.find((t) => t.type === '整体');
    const weakTypes = targetTypes.filter((t) => t.type !== '整体' && overall && t.rate < overall.rate * 0.8);
    const strongTypes = targetTypes.filter((t) => t.type !== '整体' && overall && t.rate > overall.rate * 1.1);
    if (weakTypes.length) {
      tips.push(`⚠️ 补量建议：「${weakTypes.map((t) => t.type).join('、')}」类用户触达率偏低（${weakTypes.map((t) => `${t.type}=${fR(t.rate)}`).join('、')}），建议针对性增加该类用户的曝光资源`);
    }
    if (strongTypes.length) {
      tips.push(`✅ 优势用户：「${strongTypes.map((t) => t.type).join('、')}」类用户触达表现突出（${strongTypes.map((t) => `${t.type}=${fR(t.rate)}`).join('、')}），可适当维持`);
    }
  }

  if (positions.length > 1) {
    const best = positions[0];
    const worst = positions[positions.length - 1];
    if (best.erpi > worst.erpi * 1.5 && worst.imp > 1000) {
      tips.push(`📊 资源位效率差异大：「${best.pos}」每曝光收入 ${best.erpi.toFixed(4)} 远优于「${worst.pos}」的 ${worst.erpi.toFixed(4)}，建议将曝光向高效资源位倾斜`);
    }
    const declining = positions.filter((x) => x.trend === 'declining');
    if (declining.length) {
      tips.push(`📉 效率递减：「${declining.map((x) => x.pos).join('、')}」近期CTR呈下降趋势，可能已接近饱和，不建议继续加量`);
    }
  }

  if (!tips.length) {
    tips.push('ℹ️ 当前各维度数据表现均衡，建议维持现有策略并持续观察');
  }
  return tips;
}

function renderLatestWeekTable() {
  const wrap = $('reachLatestWeekTable');
  const hintEl = $('latestWeekHint');
  if (!wrap) return;

  const sourceFilterEl = $('reachSourceFilter');
  if (sourceFilterEl && opsRows.length) {
    const daily = opsRows.filter(opsRowIsDailyRow);
    const sources = Array.from(new Set(daily.map((r) => String(r[OPS_COLS.source] ?? '').trim() || '(空)'))).filter(Boolean).sort();
    const curVal = sourceFilterEl.value;
    const opts = ['<option value="">全部</option>'].concat(sources.map((s) => `<option value="${escapeHtml(s)}"${s === curVal ? ' selected' : ''}>${escapeHtml(s)}</option>`));
    sourceFilterEl.innerHTML = opts.join('');
  }

  const sourceFilter = sourceFilterEl?.value || '';

  if (!rows.length) {
    wrap.innerHTML = '<div class="muted">暂无数据，请先加载 CSV。</div>';
    if (hintEl) hintEl.textContent = '';
    return;
  }

  const colProject = resolveCol(COL.project);
  const colLaunch = resolveCol(COL.launchDate);
  const colDay = resolveCol(COL.onlineDay);
  const colTargetType = resolveCol(COL.targetType);
  const colTargetSize = resolveCol(COL.targetSize);
  const colReached = resolveCol(COL.reached);
  const colCategory = resolveCol(COL.category);

  if (![colProject, colLaunch, colDay, colTargetSize, colReached].every((x) => x)) {
    wrap.innerHTML = '<div class="muted">字段不完整，无法生成明细</div>';
    return;
  }

  const targetTypeWanted = String($('reachTargetType')?.value || '整体').trim();

  const projMap = new Map();
  for (const r of rows) {
    const tt = String(r[colTargetType] ?? '').trim();
    if (targetTypeWanted && targetTypeWanted !== '(全部)' && tt !== targetTypeWanted) continue;
    const project = String(r[colProject] ?? '').trim();
    const launchStr = String(r[colLaunch] ?? '').trim();
    const launch = parseDate(launchStr);
    if (!launch || !project) continue;
    const day = num(r[colDay]);
    const size = num(r[colTargetSize]);
    const reached = num(r[colReached]);
    const category = String(r[colCategory] ?? '').trim() || '(空)';
    const key = project;
    const cur = projMap.get(key) || {
      project,
      launchStr,
      launch,
      category,
      size: 0,
      maxDay: -1,
      reachedAtMax: 0,
    };
    cur.size = Math.max(cur.size, size);
    if (day > cur.maxDay) {
      cur.maxDay = day;
      cur.reachedAtMax = reached;
    } else if (day === cur.maxDay) {
      cur.reachedAtMax = Math.max(cur.reachedAtMax, reached);
    }
    projMap.set(key, cur);
  }

  let maxLaunch = null;
  for (const p of projMap.values()) {
    if (!maxLaunch || p.launch > maxLaunch) maxLaunch = p.launch;
  }
  if (!maxLaunch) {
    wrap.innerHTML = '<div class="muted">无有效上线日期数据</div>';
    if (hintEl) hintEl.textContent = '';
    return;
  }

  const weekAgo = new Date(maxLaunch);
  weekAgo.setDate(weekAgo.getDate() - 6);

  const wishItems = Array.from(projMap.values())
    .filter((p) => p.launch >= weekAgo && p.launch <= maxLaunch && p.size > 0 && p.project.includes('祈愿'));

  if (!wishItems.length) {
    wrap.innerHTML = '<div class="muted">最新周无上线祈愿项目</div>';
    if (hintEl) hintEl.textContent = '';
    return;
  }

  const opsMap = buildOpsWishMap();
  const ctrBench = buildOpsCtrBenchmark();
  const { getCategoryBenchmark, getProjectHistory } = buildHistoricalBenchmarks(colCategory, targetTypeWanted);

  const merged = wishItems.map((p) => {
    const ops = matchOpsProject(p.project, opsMap);
    const currentRate = p.size > 0 ? p.reachedAtMax / p.size : 0;
    const base = extractBaseProject(p.project);
    const catBench = getCategoryBenchmark(p.category, p.maxDay, p.project);
    const projHist = getProjectHistory(base, p.project, p.maxDay);
    const judge = judgeReachLevel(currentRate, catBench, projHist, p.maxDay);
    const ctrVal = ops && ops.imp > 0 ? ops.clk / ops.imp : null;
    const ctrJudge = ctrVal != null ? judgeCtrLevel(ctrVal, ctrBench) : null;
    const targetTypes = analyzeTargetTypes(p.project);
    const positions = analyzePositions(p.project, sourceFilter || undefined);
    const suggestions = generateProjectSuggestions(p, targetTypes, positions);
    return {
      ...p,
      currentRate,
      opsImp: ops ? ops.imp : null,
      opsClk: ops ? ops.clk : null,
      opsRev: ops ? ops.rev : null,
      opsDayCount: ops ? ops.dayCount : null,
      judge, ctrJudge, targetTypes, positions, suggestions,
    };
  });

  merged.sort((a, b) => {
    const dt = b.launch.getTime() - a.launch.getTime();
    if (dt !== 0) return dt;
    return (b.opsImp || 0) - (a.opsImp || 0);
  });

  const crossInsights = buildCrossProjectPositionMatrix(merged.map((p) => p.project), sourceFilter || undefined);

  const fmtD = (d) => d.toISOString().slice(0, 10);
  const fmtInt = (n) => Math.round(n).toLocaleString('zh-CN');
  const fmtRate = (n) => `${(n * 100).toFixed(2)}%`;

  const hasOps = merged.some((p) => p.opsImp != null);

  if (hintEl) {
    const parts = [
      `上线日期：${fmtD(weekAgo)} ~ ${fmtD(maxLaunch)}`,
      `${merged.length} 个祈愿项目`,
      `目标用户类型：${targetTypeWanted}`,
    ];
    if (hasOps) parts.push('曝光/U-CTR口径：运营宣推（累计7日）');
    parts.push('点击行展开详情与建议');
    hintEl.textContent = parts.join(' · ');
  }

  const trs = merged
    .map((p, idx) => {
      const reachPct = (p.currentRate * 100).toFixed(2);
      const impStr = p.opsImp != null ? fmtInt(p.opsImp) : '-';
      const uctr = p.opsImp != null && p.opsImp > 0 ? fmtRate(p.opsClk / p.opsImp) : '-';
      const cj = p.ctrJudge;
      const ctrTag = cj ? ` <span class="lwt-tag" style="background:${cj.color}">${cj.label}</span>` : '';
      const erpi = p.opsImp != null && p.opsImp > 0
        ? (p.opsRev / p.opsImp).toLocaleString('zh-CN', { minimumFractionDigits: 4, maximumFractionDigits: 4 })
        : '-';
      const j = p.judge;

      // expanded detail
      let detailHtml = '<div class="lwt-detail">';

      // suggestions
      detailHtml += '<div class="lwt-section"><div class="lwt-section__title">🎯 运营建议</div><ul class="lwt-tips">';
      for (const tip of p.suggestions) detailHtml += `<li>${tip}</li>`;
      detailHtml += '</ul></div>';

      // target user type breakdown
      if (p.targetTypes.length > 0) {
        detailHtml += '<div class="lwt-section"><div class="lwt-section__title">👥 各目标用户类型触达情况</div>';
        detailHtml += '<table class="table table--compact"><thead><tr><th>用户类型</th><th>目标用户数</th><th>D' + p.maxDay + '触达率</th><th>差异</th></tr></thead><tbody>';
        const overall = p.targetTypes.find((t) => t.type === '整体');
        for (const t of p.targetTypes) {
          const diffStr = overall && t.type !== '整体'
            ? (t.rate > overall.rate * 1.05 ? `<span style="color:#16a34a">↑高于整体</span>` : t.rate < overall.rate * 0.95 ? `<span style="color:#dc2626">↓低于整体</span>` : '<span class="muted">≈持平</span>')
            : (t.type === '整体' ? '<span class="muted">基准</span>' : '-');
          detailHtml += `<tr><td>${t.type}</td><td class="mono">${fmtInt(t.size)}</td><td class="mono">${fmtRate(t.rate)}</td><td>${diffStr}</td></tr>`;
        }
        detailHtml += '</tbody></table></div>';
      }

      // resource position breakdown
      if (p.positions.length > 0) {
        detailHtml += '<div class="lwt-section"><div class="lwt-section__title">📍 各资源位表现（最近7日）</div>';
        detailHtml += '<table class="table table--compact"><thead><tr><th>资源位</th><th>曝光UV</th><th>CTR</th><th>每曝光收入</th><th>趋势</th></tr></thead><tbody>';
        for (const pos of p.positions) {
          const trendIcon = pos.trend === 'declining' ? '<span style="color:#dc2626">📉 下降</span>' : pos.trend === 'rising' ? '<span style="color:#16a34a">📈 上升</span>' : '<span class="muted">→ 稳定</span>';
          detailHtml += `<tr><td>${pos.pos}</td><td class="mono">${fmtInt(pos.imp)}</td><td class="mono">${fmtRate(pos.ctr)}</td><td class="mono">${pos.erpi.toFixed(4)}</td><td>${trendIcon}</td></tr>`;
        }
        detailHtml += '</tbody></table></div>';
      }

      // benchmark detail
      detailHtml += `<div class="lwt-section"><div class="lwt-section__title">📊 触达率评估依据</div><div class="muted" style="font-size:12px;line-height:1.6">${j.detail}<br>💡 ${j.suggestion}</div>`;
      if (cj) detailHtml += `<div class="muted" style="font-size:12px;margin-top:4px">U-CTR评估：${cj.detail}</div>`;
      detailHtml += '</div>';

      detailHtml += '</div>';

      return `<tr class="lwt-row" data-idx="${idx}" onclick="this.classList.toggle('is-open');this.nextElementSibling.classList.toggle('is-open')">
        <td><span class="lwt-expand">▸</span> ${p.project}</td>
        <td class="mono">${p.launchStr}</td>
        <td class="mono">${impStr}</td>
        <td class="mono">${uctr}${ctrTag}</td>
        <td class="mono">${erpi}</td>
        <td class="mono">${reachPct}%<span class="muted" style="margin-left:4px;font-size:11px">(D${p.maxDay})</span></td>
        <td><span class="lwt-tag" style="background:${j.color}">${j.label}</span></td>
      </tr>
      <tr class="lwt-detail-row" data-idx="${idx}"><td colspan="7">${detailHtml}</td></tr>`;
    })
    .join('');

  let insightsHtml = '';
  if (crossInsights.swapSuggestions.length || crossInsights.saturationAlerts.length) {
    insightsHtml = '<div class="lwt-insights">';
    insightsHtml += `<div class="lwt-insights__title">🔍 跨项目综合洞察${sourceFilter ? `（宣发来源：${escapeHtml(sourceFilter)}）` : ''}</div>`;
    if (crossInsights.swapSuggestions.length) {
      insightsHtml += '<div class="lwt-insight-group"><div class="lwt-insight-group__title">🔄 资源位调换建议</div><ul class="lwt-tips">';
      for (const s of crossInsights.swapSuggestions.slice(0, 5)) {
        insightsHtml += `<li>${s.detail}</li>`;
      }
      insightsHtml += '</ul></div>';
    }
    if (crossInsights.saturationAlerts.length) {
      insightsHtml += '<div class="lwt-insight-group"><div class="lwt-insight-group__title">⚠️ 曝光饱和预警</div><ul class="lwt-tips">';
      for (const a of crossInsights.saturationAlerts.slice(0, 5)) {
        insightsHtml += `<li>${a.detail}</li>`;
      }
      insightsHtml += '</ul></div>';
    }
    insightsHtml += '</div>';
  }

  wrap.innerHTML = `
    <div class="tableWrap tableWrap--wide">
      <table class="table lwt-table">
        <thead><tr>
          <th>项目名称</th>
          <th>上线时间</th>
          <th>累计曝光UV</th>
          <th>U-CTR</th>
          <th>每曝光收入</th>
          <th>触达率</th>
          <th>水平</th>
        </tr></thead>
        <tbody>${trs}</tbody>
      </table>
    </div>
    ${insightsHtml}
  `;
}

function render() {
  if (!rows.length) return;
  renderLatestWeekTable();
  const res = compute();
  if (!res.ok) {
    setStatus(res.reason);
    return;
  }

  const { evalDay, days, series, byCat } = res;

  const hint = `口径：选择目标用户类型（可选“全部”）；横轴=上线天数(D0~D${evalDay})；纵轴=触达率；触达率=累计触达/目标用户数（按天取累计并做carry-forward）。`;
  const hintEl = $('reachChartHint');
  if (hintEl) hintEl.textContent = hint;

  const ctx = $('reachChart')?.getContext('2d');
  if (ctx) {
    if (chart) chart.destroy();
    chart = new Chart(ctx, {
      type: 'line',
      data: {
        labels: days.map((d) => `D${d}`),
        datasets: series.map((p, idx) => ({
          label: p.project,
          data: p.seriesPct,
          borderColor: ['#4f7cff','#40c79a','#f59e0b','#ef4444','#a855f7','#06b6d4','#22c55e','#e11d48','#0ea5e9','#84cc16','#f97316','#64748b'][idx % 12],
          backgroundColor: 'transparent',
          borderWidth: 1.6,
          pointRadius: 0,
          tension: 0.25,
        })),
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { position: 'bottom' },
          tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${c.parsed.y.toFixed(2)}%` } },
        },
        scales: {
          x: { grid: { display: false } },
          y: { min: 0, ticks: { callback: (v) => `${v}%` }, grid: { color: 'rgba(15,23,42,.08)' } },
        },
      },
    });
  }

  // details
  const wrap = $('reachDetails');
  if (wrap) {
    const cats = Array.from(byCat.keys()).sort();
    const html = cats.map((cat) => {
      const list = byCat.get(cat) || [];
      const rowsHtml = list.slice(0, 300).map((it) => `
        <tr>
          <td>${it.project}</td>
          <td class="mono">${it.size.toLocaleString('zh-CN')}</td>
          <td class="mono">${(it.reachRateAtDay * 100).toFixed(2)}%</td>
          <td class="mono">${it.launchDateStr || '-'}</td>
        </tr>
      `).join('');
      return `
        <details class="details" style="margin-top:10px">
          <summary><strong>${cat}</strong>（${list.length}个项目）</summary>
          <div class="details__body" style="padding-top:10px">
            <div class="tableWrap tableWrap--wide">
              <table class="table">
                <thead><tr><th>项目</th><th>目标用户数</th><th>D${evalDay}触达率</th><th>上线日期</th></tr></thead>
                <tbody>${rowsHtml || ''}</tbody>
              </table>
            </div>
          </div>
        </details>
      `;
    }).join('');
    wrap.innerHTML = html || '<div class="muted">无明细数据</div>';
  }

  setStatus(`已渲染：折线 ${series.length} 条；明细品类 ${byCat.size} 个。`);
}

async function onFile(file) {
  setStatus('解析中…');
  const text = await file.text();
  rows = await parseCsvText(text);
  hydrateColumns();
  render();
  setStatus(`已加载 ${rows.length} 行。请确认字段映射后查看。`);
}

function init() {
  $('reachFileInput')?.addEventListener('change', (e) => {
    const f = e.target.files?.[0];
    if (f) onFile(f);
  });
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    const f = e.dataTransfer?.files?.[0];
    if (f) onFile(f);
  });

  $('reachBindFolderBtn')?.addEventListener('click', async () => {
    setStatus('等待选择文件夹…');
    try {
      await pickAndBindFolder();
      setStatus('已绑定。点击“读取触达文件夹全部CSV并更新”。');
    } catch (e) {
      setStatus(`绑定失败：${e?.message || e}`);
    }
  });
  $('reachUpdateBtn')?.addEventListener('click', async () => {
    try {
      await loadAllFromBoundFolder();
    } catch (e) {
      setStatus(`更新失败：${e?.message || e}`);
    }
  });
  $('reachClearBtn')?.addEventListener('click', () => {
    rows = [];
    columns = [];
    if (chart) {
      chart.destroy();
      chart = null;
    }
    $('reachDetails').innerHTML = '';
    setStatus('已清空。请选择/拖入 CSV 文件开始分析。');
  });

  [
    'reachEvalDay',
    'reachTargetType',
    'reachBucketPreset',
    'reachLaunchStart',
    'reachLaunchEnd',
    'reachProjectQ',
    'reachTopN',
    'reachSourceFilter',
  ].forEach((id) => {
    $(id)?.addEventListener('change', render);
    $(id)?.addEventListener('input', () => render());
  });

  function renderLaunchRangeSummary() {
    const s = $('reachLaunchRangeSummary');
    if (!s) return;
    const a = $('reachLaunchStart')?.value || '';
    const b = $('reachLaunchEnd')?.value || '';
    if (!a && !b) s.textContent = '全部';
    else if (a && b) s.textContent = `${a} ~ ${b}`;
    else s.textContent = a || b;
  }

  $('reachLaunchRangeBtn')?.addEventListener('click', () => {
    const anchor = $('reachLaunchRangeBtn');
    if (!anchor) return;
    openLaunchRangePopover(anchor, ({ start, end }) => {
      const a = $('reachLaunchStart'); const b = $('reachLaunchEnd');
      if (a) a.value = start;
      if (b) b.value = end;
      renderLaunchRangeSummary();
      render();
    });
  });

  renderLaunchRangeSummary();

  $('reachCatBtn')?.addEventListener('click', () => {
    if (!rows.length) return;
    const colCategory = resolveCol(COL.category);
    const items = Array.from(new Set(rows.map((r) => String(r[colCategory] ?? '').trim() || '(空)'))).sort();
    openFilterPopover({
      anchorEl: $('reachCatBtn'),
      title: '品类筛选（二级品类）',
      items,
      selectedSet: selectedCategories,
      onApply: (next) => {
        selectedCategories = next;
        const sum = $('reachCatSummary');
        if (sum) sum.textContent = selectedCategories.size ? summarizeSelection(selectedCategories, 3) : '全部';
        render();
      },
    });
  });

  if (window.__SNAPSHOT_DATA) {
    rows = window.__SNAPSHOT_DATA.rows || [];
    opsRows = window.__SNAPSHOT_DATA.opsRows || [];
    const hideIds = ['reachFileInput', 'reachBindFolderBtn', 'reachUpdateBtn', 'reachClearBtn', 'reachExportBtn'];
    hideIds.forEach((id) => { const el = $(id); if (el) { const p = el.closest('label') || el; p.style.display = 'none'; } });
    hydrateColumns();
    render();
    setStatus(`快照模式：已加载 ${rows.length} 行触达数据 + ${opsRows.length} 行运营宣推数据。`);
  } else {
    getBoundDirHandle().then(async (h) => {
      if (!h) { loadOpsDataSilently().then(() => { if (rows.length) renderLatestWeekTable(); }); return; }
      try {
        const perm = await h.queryPermission?.({ mode: 'read' });
        if (perm === 'granted') {
          await loadAllFromBoundFolder();
        } else {
          loadOpsDataSilently().then(() => { if (rows.length) renderLatestWeekTable(); });
          setStatus('检测到已绑定数据文件夹，请点击「读取触达文件夹全部CSV并更新」授权读取。');
        }
      } catch (_) {
        loadOpsDataSilently().then(() => { if (rows.length) renderLatestWeekTable(); });
      }
    });
  }

  $('reachExportBtn')?.addEventListener('click', exportSnapshot);
}

async function exportSnapshot() {
  if (!rows.length) { alert('请先加载数据再导出快照。'); return; }

  setStatus('正在生成可交互快照…');

  const cssText = await fetch('./styles.css').then((r) => r.text()).catch(() => '');
  const jsText = await fetch('./reach.js?' + Date.now()).then((r) => r.text()).catch(() => '');

  const dataPayload = JSON.stringify({ rows, opsRows });
  const now = new Date();
  const ts = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;

  const pageHtml = document.querySelector('.appShell__content main')?.innerHTML || '';
  const headerHtml = document.querySelector('.appShell__content header')?.innerHTML || '';

  const html = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>\u7948\u613f\u89e6\u8fbe\u770b\u677f\u5feb\u7167 ${ts}</title>
<style>${cssText}</style>
<style>.sidebar{display:none!important}.appShell{grid-template-columns:1fr!important}
.snap-banner{background:linear-gradient(135deg,#f0f4ff,#fdf4ff);border-bottom:1px solid rgba(79,124,255,.15);padding:10px 18px;font-size:12px;color:rgba(15,23,42,.7);display:flex;align-items:center;justify-content:space-between}
.snap-banner strong{color:rgba(15,23,42,.9)}</style>
</head>
<body>
<div class="appShell">
<div class="appShell__content">
<div class="snap-banner"><span>\u{1f4cb} <strong>\u53ef\u4ea4\u4e92\u5feb\u7167</strong>\u3000\u751f\u6210\u65f6\u95f4\uff1a${ts}\u3000\u00b7\u3000\u7b5b\u9009/\u56fe\u8868/\u5c55\u5f00\u5747\u53ef\u6b63\u5e38\u4f7f\u7528</span></div>
<header class="header">${headerHtml}</header>
<main class="container">${pageHtml}</main>
</div>
</div>
<script>window.__SNAPSHOT_DATA=${dataPayload};<\/script>
<script src="https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js"><\/script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"><\/script>
<script>${jsText}<\/script>
</body>
</html>`;

  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
  a.download = `\u7948\u613f\u89e6\u8fbe\u5feb\u7167_${dateStr}.html`;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  setStatus('快照已下载！同事打开即可交互筛选。');
}

init();

