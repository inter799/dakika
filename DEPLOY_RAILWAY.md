# 部署到 Railway 指南（纯本地模式，无需登录、无需 API）

本项目已改造为**纯本地模式**：打开网页即可上传文档、生成题目、答题、存题库。
无任何登录、无外部 API 调用、无数据库依赖。

## 一、准备 GitHub 仓库

```bash
cd "c:\Users\1111\Desktop\测试"
git init
git add .
git commit -m "quiz generator - local mode"
git remote add origin https://github.com/你的用户名/quiz-generator.git
git push -u origin main
```

## 二、Railway 部署

1. 打开 https://railway.app ，**用 GitHub 登录**（建议用无痕窗口避免误登录错账号）。
2. **New Project → Deploy from GitHub repo**，选择你的仓库。
3. Railway 会自动识别 `Procfile`，直接开始构建部署，**无需任何额外配置**。
4. 部署完成后，Railway 会分配一个 `xxx.up.railway.app` 域名，点击即可访问。

> 环境变量：**无需设置任何变量**。所有数据存在容器临时磁盘 `/tmp/data`。
> 注意：Railway 容器重启/休眠后数据会清空，但程序本身正常运行。
> 若需要数据持久化，可在 Railway 控制台挂载 Volume 到 `/tmp/data`（可选）。

## 三、绑定自定义域名 gstest.cn.mt（DNSHE 注册）

1. Railway 控制台 → 你的项目 → **Settings → Custom Domains**。
2. 输入 `gstest.cn.mt`，点击 Add Domain。
3. Railway 会生成一个 **CNAME 目标地址**（形如 `xxx.up.railway.app` 或 `cname.railway.app`）。
4. 登录 DNSHE（你的域名注册商）控制台，进入 **DNS 解析 / 域名解析** 管理：
   - 添加一条记录：
     | 记录类型 | 主机记录 | 记录值（Railway 提供的 CNAME） | TTL |
     |---------|---------|-------------------------------|-----|
     | CNAME   | @       | （Railway 生成的 cname 地址）   | 600 |
   - 若 DNSHE 不支持 @ 用 CNAME，改用主机记录 `www` 并添加 URL 转发 `gstest.cn.mt → www.gstest.cn.mt`。
5. 等待 DNS 生效（通常 5–30 分钟）。Railway 会自动签发 SSL 证书，访问 `https://gstest.cn.mt` 即可。

## 四、本地运行（测试）

```bash
pip install -r requirements.txt
python app.py
# 浏览器打开 http://127.0.0.1:5000
```

## 五、已移除的内容

- ❌ 登录系统（`/login`、`/logout`、session、用户隔离目录）
- ❌ 所有 AI 模型 API（OpenAI / DeepSeek / Qwen / Claude / Gemini / Anthropic / 自定义）
- ❌ 网络搜题功能（DuckDuckGo 搜索、网页抓取）
- ❌ Supabase 依赖
- ✅ 保留：文件解析、本地规则生成题目、题库空间、错题本、答题卡

## 六、支持的文件格式

txt / md / json / csv / xml / log / yaml / html / doc / docx / pdf / rtf / xlsx / pptx / jpg / png / bmp / gif
