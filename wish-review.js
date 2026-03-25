/* 祈愿单项目复盘：与运营宣推 / 触达共用「资源位数据更新」根目录绑定，进入「祈愿收入复盘」下四套子文件夹校验并列出最新 CSV。 */

const $ = (id) => document.getElementById(id);

const DB_NAME = 'ops-dashboard-local-db';
const DB_STORE = 'kv';
const DB_KEY_DIR_HANDLE = 'boundDirHandle';

/** 绑定根目录下进入「祈愿收入复盘」 */
const REVIEW_ROOT_CANDIDATES = ['祈愿收入复盘'];

/** 与生成脚本 --data-bundle 子目录名一致 */
const BUNDLE_SUBS = {
  main: ['整体数据监测'],
  work: ['作品明细表'],
  bench: ['历史品类池数据-原漫改耽美历史数据', '历史品类池数据'],
  layer: ['分层用户监测'],
};

const SUB_LABEL = {
  main: '整体数据监测',
  work: '作品明细表',
  bench: '历史品类池数据（原漫改耽美历史数据）',
  layer: '分层用户监测',
};

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
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
    return (await dbGet(DB_KEY_DIR_HANDLE)) || null;
  } catch (_) {
    return null;
  }
}

async function pickAndBindFolder() {
  if (!window.showDirectoryPicker) {
    throw new Error('当前浏览器不支持「绑定文件夹」（需 Chrome/Edge 的 File System Access API）。');
  }
  if (window.top && window.top !== window.self) {
    throw new Error('当前在 iframe 内，请新窗口打开本页再绑定。');
  }
  const handle = await window.showDirectoryPicker({ mode: 'read' });
  await dbSet(DB_KEY_DIR_HANDLE, handle);
  return handle;
}

async function resolveFirstChildDir(parent, names) {
  for (const name of names) {
    try {
      const h = await parent.getDirectoryHandle(name, { create: false });
      return { handle: h, name };
    } catch (_) {}
  }
  return null;
}

async function listCsvWithMtime(dirHandle) {
  const out = [];
  // eslint-disable-next-line no-restricted-syntax
  for await (const entry of dirHandle.values()) {
    if (!entry || entry.kind !== 'file') continue;
    if (!entry.name.toLowerCase().endsWith('.csv')) continue;
    const file = await entry.getFile();
    out.push({ name: entry.name, lastModified: file.lastModified });
  }
  return out;
}

function pickLatestForMain(list) {
  if (!list.length) return null;
  const prefer = list.filter((x) => /^①/.test(x.name) || x.name.includes('整体数据监测'));
  const pool = prefer.length ? prefer : list;
  pool.sort((a, b) => b.lastModified - a.lastModified);
  return pool[0];
}

function pickLatestForWork(list) {
  if (!list.length) return null;
  const prefer = list.filter((x) => x.name.includes('作品明细表'));
  const pool = prefer.length ? prefer : list;
  pool.sort((a, b) => b.lastModified - a.lastModified);
  return pool[0];
}

function pickLatestForBench(list) {
  if (!list.length) return null;
  list.sort((a, b) => b.lastModified - a.lastModified);
  return list[0];
}

function pickLatestForLayer(list) {
  if (!list.length) return null;
  const prefer = list.filter((x) => /^②/.test(x.name) || x.name.includes('分层用户监测'));
  const pool = prefer.length ? prefer : list;
  pool.sort((a, b) => b.lastModified - a.lastModified);
  return pool[0];
}

async function scanBundleFromRoot(rootHandle) {
  const perm = await rootHandle.queryPermission?.({ mode: 'read' });
  if (perm !== 'granted') {
    const req = await rootHandle.requestPermission?.({ mode: 'read' });
    if (req !== 'granted') throw new Error('未获得文件夹读取权限。');
  }

  const review = await resolveFirstChildDir(rootHandle, REVIEW_ROOT_CANDIDATES);
  if (!review) {
    throw new Error(
      `未找到「${REVIEW_ROOT_CANDIDATES.join('」或「')}」文件夹。请在绑定的「资源位数据更新」目录下创建「祈愿收入复盘」。`,
    );
  }

  const result = {
    reviewFolderName: review.name,
    rows: [],
  };

  async function one(key, subNames, picker) {
    const sub = await resolveFirstChildDir(review.handle, subNames);
    const label = SUB_LABEL[key];
    if (!sub) {
      result.rows.push({ key, label, ok: false, detail: `缺少子文件夹「${subNames[0]}」` });
      return;
    }
    const files = await listCsvWithMtime(sub.handle);
    const picked = picker(files);
    if (!picked) {
      result.rows.push({
        key,
        label,
        ok: false,
        detail: `「${sub.name}」下无 CSV`,
      });
      return;
    }
    result.rows.push({
      key,
      label,
      ok: true,
      sub: sub.name,
      file: picked.name,
      lastModified: picked.lastModified,
    });
  }

  await one('main', BUNDLE_SUBS.main, pickLatestForMain);
  await one('work', BUNDLE_SUBS.work, pickLatestForWork);
  await one('bench', BUNDLE_SUBS.bench, pickLatestForBench);
  await one('layer', BUNDLE_SUBS.layer, pickLatestForLayer);

  return result;
}

function formatMtime(ts) {
  if (ts == null) return '—';
  try {
    return new Date(ts).toLocaleString('zh-CN', { hour12: false });
  } catch (_) {
    return '—';
  }
}

function renderBundleTable(scan) {
  const host = $('wishReviewBundleTable');
  if (!host) return;

  if (!scan) {
    host.innerHTML = '<p class="muted">尚未扫描。</p>';
    return;
  }

  const rows = scan.rows
    .map((r) => {
      const st = r.ok
        ? `<span style="color:var(--good)">已就绪</span>`
        : `<span style="color:var(--bad)">缺失</span>`;
      const file = r.ok
        ? `${escapeHtml(r.sub)}/${escapeHtml(r.file)} <span class="muted">· ${formatMtime(r.lastModified)}</span>`
        : escapeHtml(r.detail || '');
      return `<tr><td>${escapeHtml(r.label)}</td><td>${st}</td><td style="font-size:12px">${file}</td></tr>`;
    })
    .join('');

  host.innerHTML = `
    <p class="muted" style="margin:0 0 10px">已进入「<strong>${escapeHtml(scan.reviewFolderName)}</strong>」，与各数据源子文件夹对齐（与本地 <code>生成_人鱼全期结论表.py --data-bundle</code> 约定一致）。</p>
    <div class="tableWrap">
      <table class="table">
        <thead><tr><th>数据源</th><th>状态</th><th>将使用的 CSV（同夹内最新）</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function setHint(msg) {
  const el = $('wishReviewHint');
  if (el) el.textContent = msg;
}

async function runScan() {
  setHint('扫描中…');
  renderBundleTable(null);
  try {
    const root = await getBoundDirHandle();
    if (!root) {
      setHint('尚未绑定数据文件夹。请先点击「绑定数据文件夹」（与其它看板共用同一目录，一般为「资源位数据更新」）。');
      $('wishReviewBundleTable').innerHTML = '<p class="muted">未绑定。</p>';
      return;
    }
    const scan = await scanBundleFromRoot(root);
    window.__WISH_REVIEW_BUNDLE_SCAN__ = scan;
    renderBundleTable(scan);
    const okn = scan.rows.filter((x) => x.ok).length;
    setHint(`扫描完成：${okn}/4 项数据源就绪。复盘 HTML 需在本机运行生成脚本（见下方说明）。`);
  } catch (e) {
    setHint(`扫描失败：${e?.message || e}`);
    const host = $('wishReviewBundleTable');
    if (host) host.innerHTML = `<p class="muted">${escapeHtml(e?.message || String(e))}</p>`;
  }
}

function onBindClick() {
  (async () => {
    try {
      await pickAndBindFolder();
      setHint('已绑定，正在扫描…');
      await runScan();
    } catch (e) {
      if (e?.name === 'AbortError') return;
      setHint(`绑定失败：${e?.message || e}`);
    }
  })();
}

function bindClicks(selector, handler) {
  document.querySelectorAll(selector).forEach((el) => {
    el.addEventListener('click', handler);
  });
}

function setup() {
  bindClicks('.js-wish-review-bind', onBindClick);
  bindClicks('.js-wish-review-scan', () => runScan());

  getBoundDirHandle().then((h) => {
    if (h) runScan();
    else {
      setHint('与其它看板相同：请先绑定「资源位数据更新」根目录，本页会自动进入「祈愿收入复盘」并校验四套子文件夹。');
    }
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', setup);
} else {
  setup();
}
