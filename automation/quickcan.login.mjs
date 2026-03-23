import fs from 'node:fs/promises';
import path from 'node:path';
import readline from 'node:readline/promises';
import { chromium } from 'playwright';
import { QUICKCAN_CONFIG } from './quickcan.config.mjs';

async function ensureDir(p) {
  await fs.mkdir(path.dirname(p), { recursive: true });
}

async function main() {
  await ensureDir(QUICKCAN_CONFIG.authStatePath);

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  await page.goto(QUICKCAN_CONFIG.boards[0].url, {
    waitUntil: 'domcontentloaded',
    timeout: QUICKCAN_CONFIG.timeoutMs,
  });

  // 用户手动完成登录，回车后保存态
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  await rl.question('\n请在浏览器中完成登录，然后回车继续保存登录态...\n');
  await rl.close();

  await context.storageState({ path: QUICKCAN_CONFIG.authStatePath });
  await browser.close();
  console.log(`登录态已保存: ${QUICKCAN_CONFIG.authStatePath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

