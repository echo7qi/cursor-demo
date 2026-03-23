# Quickcan 自动同步说明

目标：自动从 3 个看板下载 CSV，并写入本地数据目录，供当前网页读取。

## 1. 安装依赖

```bash
cd /Users/wangqiqi/Desktop/cursor-demo
npm install
npx playwright install chromium
```

## 2. 首次登录（仅一次）

```bash
npm run quickcan:login
```

在弹出的浏览器里完成登录后，回车保存登录态。

## 3. 手动跑一次同步

```bash
npm run quickcan:sync
```

若需要可视化调试（推荐首次调试）：

```bash
QUICKCAN_HEADLESS=0 npm run quickcan:sync
```

说明：若看板只能导出 Excel，脚本会自动把首个工作表转换为 CSV 并写入目标目录。

默认输出目录：

`/Users/wangqiqi/Desktop/dashboard-data`

会生成 3 个子目录：

- `运营宣推`
- `祈愿`
- `触达`

## 4. 安装定时任务（周一/三/五 12:00）

```bash
npm run quickcan:install-schedule
```

手动立即触发一次：

```bash
launchctl start com.cursor.quickcan.sync
```

## 5. 看板读取方式

在网页中将“绑定数据文件夹”绑定到：

`/Users/wangqiqi/Desktop/dashboard-data`

之后刷新页面即可读到自动更新的数据。

## 6. 可配置项

可在运行命令前设置环境变量：

- `QUICKCAN_DATA_ROOT`：数据输出根目录
- `QUICKCAN_AUTH_STATE`：登录态文件路径
- `QUICKCAN_HEADLESS`：`1`=无头模式；其他值=有界面模式（默认有界面，便于排错）

例如：

```bash
QUICKCAN_DATA_ROOT="/Users/xxx/my-data" npm run quickcan:sync
```

