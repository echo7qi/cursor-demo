/* 祈愿单项目复盘：与运营宣推 / 触达共用「资源位数据更新」根目录绑定；对标池 CSV 与「整体数据监测」同夹。
 * 扫描结果写入 window.__WISH_REVIEW_BUNDLE_SCAN__（无主区 UI）。 */

const DB_NAME = 'ops-dashboard-local-db';
const DB_STORE = 'kv';
const DB_KEY_DIR_HANDLE = 'boundDirHandle';

/** 绑定根目录下进入「祈愿收入复盘」 */
const REVIEW_ROOT_CANDIDATES = ['祈愿收入复盘'];

/** 与生成脚本 --data-bundle 子目录名一致 */
const BUNDLE_SUBS = {
  main: ['整体数据监测'],
  work: ['作品明细表'],
  layer: ['分层用户监测'],
};

const SUB_LABEL = {
  main: '整体数据监测',
  work: '作品明细表',
  bench: '历史品类池（与整体数据监测同夹）',
  layer: '分层用户监测',
};

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

function pickBenchmarkFromMainDir(list) {
  if (!list.length) return null;
  const noMonitor = list.filter((x) => !x.name.includes('整体数据监测'));
  if (!noMonitor.length) return null;
  const prefer = noMonitor.filter(
    (x) =>
      /池.*历史|历史.*池|池历史数据|漫改耽美池/i.test(x.name),
  );
  const pool = prefer.length ? prefer : noMonitor;
  pool.sort((a, b) => b.lastModified - a.lastModified);
  return pool[0];
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

  const mainSub = await resolveFirstChildDir(review.handle, BUNDLE_SUBS.main);
  if (!mainSub) {
    result.rows.push({
      key: 'main',
      label: SUB_LABEL.main,
      ok: false,
      detail: `缺少子文件夹「${BUNDLE_SUBS.main[0]}」`,
    });
    result.rows.push({
      key: 'bench',
      label: SUB_LABEL.bench,
      ok: false,
      detail: '依赖「整体数据监测」文件夹',
    });
  } else {
    const mainFiles = await listCsvWithMtime(mainSub.handle);
    const mainP = pickLatestForMain(mainFiles);
    if (!mainP) {
      result.rows.push({
        key: 'main',
        label: SUB_LABEL.main,
        ok: false,
        detail: `「${mainSub.name}」下无监测表 CSV`,
      });
    } else {
      result.rows.push({
        key: 'main',
        label: SUB_LABEL.main,
        ok: true,
        sub: mainSub.name,
        file: mainP.name,
        lastModified: mainP.lastModified,
      });
    }
    const benchP = pickBenchmarkFromMainDir(mainFiles);
    if (!benchP) {
      result.rows.push({
        key: 'bench',
        label: SUB_LABEL.bench,
        ok: false,
        detail: `「${mainSub.name}」内除监测表外未找到对标池 CSV（如 *池历史数据*.csv）`,
      });
    } else {
      result.rows.push({
        key: 'bench',
        label: SUB_LABEL.bench,
        ok: true,
        sub: `${mainSub.name}（同夹）`,
        file: benchP.name,
        lastModified: benchP.lastModified,
      });
    }
  }

  await one('work', BUNDLE_SUBS.work, pickLatestForWork);
  await one('layer', BUNDLE_SUBS.layer, pickLatestForLayer);

  return result;
}

async function runScan() {
  try {
    const root = await getBoundDirHandle();
    if (!root) return;
    const scan = await scanBundleFromRoot(root);
    window.__WISH_REVIEW_BUNDLE_SCAN__ = scan;
  } catch (e) {
    console.warn('[wish-review] 扫描失败', e?.message || e);
  }
}

function onBindClick() {
  (async () => {
    try {
      await pickAndBindFolder();
      await runScan();
    } catch (e) {
      if (e?.name === 'AbortError') return;
      console.warn('[wish-review] 绑定失败', e?.message || e);
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
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', setup);
} else {
  setup();
}
