import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { execSync } from 'node:child_process';

const home = os.homedir();
const projectDir = path.resolve('.');
const label = 'com.cursor.quickcan.sync';
const plistPath = path.join(home, 'Library/LaunchAgents', `${label}.plist`);
const nodeBin = process.execPath;
const scriptPath = path.join(projectDir, 'automation/quickcan.sync.mjs');
const logDir = path.join(projectDir, 'automation/logs');
const outLog = path.join(logDir, 'quickcan-sync.out.log');
const errLog = path.join(logDir, 'quickcan-sync.err.log');

async function main() {
  await fs.mkdir(logDir, { recursive: true });
  await fs.mkdir(path.dirname(plistPath), { recursive: true });

  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>${label}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${nodeBin}</string>
    <string>${scriptPath}</string>
  </array>
  <key>WorkingDirectory</key><string>${projectDir}</string>
  <key>RunAtLoad</key><false/>
  <key>StartCalendarInterval</key>
  <array>
    <dict><key>Weekday</key><integer>1</integer><key>Hour</key><integer>12</integer><key>Minute</key><integer>0</integer></dict>
    <dict><key>Weekday</key><integer>3</integer><key>Hour</key><integer>12</integer><key>Minute</key><integer>0</integer></dict>
    <dict><key>Weekday</key><integer>5</integer><key>Hour</key><integer>12</integer><key>Minute</key><integer>0</integer></dict>
  </array>
  <key>StandardOutPath</key><string>${outLog}</string>
  <key>StandardErrorPath</key><string>${errLog}</string>
</dict>
</plist>`;

  await fs.writeFile(plistPath, plist, 'utf8');

  try { execSync(`launchctl unload ${plistPath}`, { stdio: 'ignore' }); } catch {}
  execSync(`launchctl load ${plistPath}`, { stdio: 'inherit' });

  console.log(`\n已安装定时任务: ${label}`);
  console.log(`- 配置文件: ${plistPath}`);
  console.log(`- 输出日志: ${outLog}`);
  console.log(`- 错误日志: ${errLog}`);
  console.log('\n可手动立即执行一次：');
  console.log(`launchctl start ${label}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

