import fs from 'node:fs/promises';
import path from 'node:path';
import { chromium } from 'playwright';
import xlsx from 'xlsx';
import { QUICKCAN_CONFIG } from './quickcan.config.mjs';

function ts() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function ensureDir(p) {
  await fs.mkdir(p, { recursive: true });
}

async function clickByText(page, text) {
  const re = new RegExp(escapeRegExp(text));
  const candidates = [
    page.getByRole('button', { name: re }),
    page.getByRole('tab', { name: re }),
    page.getByText(text, { exact: false }),
    page.locator(`a:has-text("${text}")`),
    page.locator(`[role="button"]:has-text("${text}")`),
  ];
  for (const loc of candidates) {
    const first = loc.first();
    const ok = await first.isVisible().catch(() => false);
    if (!ok) continue;
    await first.click({ timeout: 10_000 }).catch(() => null);
    return true;
  }
  return false;
}

async function findClickableByNames(page, names) {
  for (const text of names) {
    const re = new RegExp(escapeRegExp(text));
    const candidates = [
      page.getByRole('button', { name: re }),
      page.getByRole('menuitem', { name: re }),
      page.getByRole('link', { name: re }),
      page.getByText(text, { exact: false }),
      page.locator(`a:has-text("${text}")`),
      page.locator(`[role="button"]:has-text("${text}")`),
    ];
    for (const loc of candidates) {
      const first = loc.first();
      const ok = await first.isVisible().catch(() => false);
      if (ok) return first;
    }
  }
  return null;
}

async function setFilter(page, label, option) {
  // 尝试点击 label 附近控件，再选 option
  const labelLoc = page.getByText(label, { exact: false }).first();
  const hasLabel = await labelLoc.isVisible().catch(() => false);
  if (!hasLabel) return false;

  const trigger = labelLoc.locator(
    'xpath=ancestor::*[self::div or self::td][1]//*[self::button or @role="combobox" or contains(@class,"select")][1]'
  );
  const triggerVisible = await trigger.first().isVisible().catch(() => false);
  if (triggerVisible) {
    await trigger.first().click().catch(() => null);
  } else {
    await labelLoc.click().catch(() => null);
  }

  await page.waitForTimeout(300);
  return clickByText(page, option);
}

async function triggerDownloadAndWait(page, board) {
  const primaryNames = ['下载CSV文件', '下载CSV', '导出CSV', '导出', '下载'];
  const formatNames = ['CSV', 'csv', '导出Excel', '导出为Excel', 'Excel', '导出为CSV'];

  const primary = await findClickableByNames(page, primaryNames);
  if (!primary) return null;

  // 用 Promise.all 防止错过瞬时 download 事件
  try {
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 8000 }),
      primary.click({ timeout: 10_000 }),
    ]);
    return download;
  } catch (_) {
    // 可能先弹出“格式菜单”，再点 CSV 才会下载
  }

  await page.waitForTimeout(300);
  const csvBtn = await findClickableByNames(page, formatNames);
  if (!csvBtn) return null;

  try {
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: QUICKCAN_CONFIG.timeoutMs }),
      csvBtn.click({ timeout: 10_000 }),
    ]);
    return download;
  } catch (_) {
    // ignore and try board-specific flow
  }

  // 运营宣推页专用：全选 -> 批量导出Excel
  if (board?.id === 'ops') {
    const selectAll = await findClickableByNames(page, ['全选']);
    if (selectAll) {
      await selectAll.click({ timeout: 10_000 }).catch(() => null);
      await page.waitForTimeout(300);
    }
    const batch = await findClickableByNames(page, ['批量导出Excel', '批量导出', '导出Excel']);
    if (batch) {
      try {
        const [download] = await Promise.all([
          page.waitForEvent('download', { timeout: QUICKCAN_CONFIG.timeoutMs }),
          batch.click({ timeout: 10_000 }),
        ]);
        return download;
      } catch (_) {
        return null;
      }
    }
  }
  return null;
}

async function downloadBoard(page, board, outputDir) {
  await page.goto(board.url, {
    waitUntil: 'domcontentloaded',
    timeout: QUICKCAN_CONFIG.timeoutMs,
  });
  await page.waitForTimeout(1200);

  for (const t of board.clickTexts || []) {
    await clickByText(page, t);
    await page.waitForTimeout(300);
  }

  for (const f of board.filters || []) {
    await setFilter(page, f.label, f.option);
    await page.waitForTimeout(500);
  }

  const download = await triggerDownloadAndWait(page, board);
  if (!download) {
    const logsDir = path.resolve('automation/logs');
    await ensureDir(logsDir);
    const shot = path.join(logsDir, `${board.id}_download_failed_${ts()}.png`);
    await page.screenshot({ path: shot, fullPage: true }).catch(() => null);
    throw new Error(`[${board.name}] 下载未触发（已保存截图：${shot}）`);
  }
  const suggested = download.suggestedFilename() || `${board.filePrefix}.csv`;
  const ext = path.extname(suggested).toLowerCase() || '.csv';

  const rawPath = path.join(outputDir, `${board.filePrefix}_${ts()}${ext || '.tmp'}`);
  await download.saveAs(rawPath);

  let stampedPath;
  let latestPath;
  if (ext === '.csv') {
    stampedPath = rawPath;
    latestPath = path.join(outputDir, `${board.filePrefix}_latest.csv`);
    await fs.copyFile(stampedPath, latestPath);
  } else if (ext === '.xlsx' || ext === '.xls') {
    const wb = xlsx.readFile(rawPath);
    const firstSheet = wb.SheetNames[0];
    if (!firstSheet) throw new Error(`[${board.name}] Excel 无可用工作表`);
    const csv = xlsx.utils.sheet_to_csv(wb.Sheets[firstSheet]);
    stampedPath = path.join(outputDir, `${board.filePrefix}_${ts()}.csv`);
    latestPath = path.join(outputDir, `${board.filePrefix}_latest.csv`);
    await fs.writeFile(stampedPath, csv, 'utf8');
    await fs.writeFile(latestPath, csv, 'utf8');
  } else {
    throw new Error(`[${board.name}] 不支持的下载格式：${suggested}`);
  }
  return { stampedPath, latestPath };
}

async function main() {
  await ensureDir(path.dirname(QUICKCAN_CONFIG.authStatePath));
  await ensureDir(QUICKCAN_CONFIG.dataRoot);

  const headless = process.env.QUICKCAN_HEADLESS === '1';
  const browser = await chromium.launch({ headless });
  const context = await browser.newContext({
    storageState: QUICKCAN_CONFIG.authStatePath,
    acceptDownloads: true,
  });
  const page = await context.newPage();

  const results = [];
  try {
    for (const board of QUICKCAN_CONFIG.boards) {
      const dir = path.join(QUICKCAN_CONFIG.dataRoot, board.subdir);
      await ensureDir(dir);
      const r = await downloadBoard(page, board, dir);
      results.push({ board: board.name, ...r });
      console.log(`[OK] ${board.name}: ${r.stampedPath}`);
    }
  } finally {
    await browser.close();
  }

  console.log('\n全部任务完成：');
  for (const r of results) {
    console.log(`- ${r.board}: ${r.stampedPath}`);
  }
}

main().catch((e) => {
  console.error('\n同步失败：', e.message || e);
  process.exit(1);
});

