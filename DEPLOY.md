# 运营宣推数据看板 · 稳定链接部署说明

## 方式一：GitHub Pages（推荐，免费稳定）

部署后获得**永久稳定链接**（同一仓库、同一 GitHub Actions，推送 `main` 即更新）：

| 页面 | 地址 |
|------|------|
| 运营宣推首页 | `https://echo7qi.github.io/cursor-demo/` 或 `.../index.html` |

### 步骤

1. **提交并推送代码**
   ```bash
   cd /Users/wangqiqi/Desktop/运营宣推数据看板
   git add .
   git commit -m "部署看板到 GitHub Pages"
   git push origin main
   ```

2. **在 GitHub 启用 Pages**
   - 打开：https://github.com/echo7qi/cursor-demo/settings/pages
   - **Source**：选择「Deploy from a branch」
   - **Branch**：选择 `main`，文件夹选择 `/ (root)`
   - 点击 **Save**

3. **等待 1–2 分钟**，访问：**https://echo7qi.github.io/cursor-demo/**

### 使用说明

- 看板为纯前端，数据在本地解析，**不上传任何数据**
- 访问链接后，点击「绑定数据文件夹」或「选择 CSV 文件」加载本地数据即可
- 链接长期有效，只要仓库存在即可访问

---

## 方式二：本地临时稳定链接（Serveo 隧道）

需要本地运行看板时，可用 Serveo 生成稳定子域名：

```bash
# 1. 启动本地服务（另开终端保持运行）
cd /Users/wangqiqi/Desktop/运营宣推数据看板
python3 -m http.server 8080

# 2. 建立隧道（自定义子域名，每次相同）
ssh -R ops-dashboard:80:localhost:8080 serveo.net
```

访问：**https://ops-dashboard.serveo.net**（需保持终端运行）

---

## 方式三：Railway 部署

项目已包含 `railway.toml`，可连接 Railway 一键部署获得稳定链接。
