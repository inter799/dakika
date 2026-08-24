# dakika

> 从任意文档自动生成题目的本地化 Web 应用 — 无登录、无外部 API、纯本地存储。

dakika 是一个开箱即用的"题库生成器"：把 `PDF / Word / Excel / PPT / 图片 / 文本文档`丢进去，自动抽取内容、生成题目、做题、保存题库。整个过程 **零账号、零网络依赖**。

---

## ✨ 功能特性

- 📄 **多格式文档解析**：doc / docx / pdf / xlsx / xls / xlsm / pptx / rtf / txt / md / csv / xml / log / yaml / html / 图片
- 🤖 **本地规则引擎出题**：支持选择、判断、填空、简答等多种题型（无需大模型 API）
- 🗂️ **题库空间（Space）**：按文档自动分组管理题库
- ❌ **错题本**：自动汇总错题，单独复习
- 📊 **答题卡 / 答题结果分析**
- 🔒 **纯本地模式**：不上传任何信息到外部服务器
- 🌐 **响应式 Web UI**：浏览器即可使用，无需安装客户端

---

## 🚀 快速开始（本地运行）

### 方式一：Windows 一键启动（推荐）

双击 `start.bat`：

```bat
:: 自动完成：检测/安装依赖 → 启动服务 → 打开浏览器
start.bat
```

### 方式二：手动启动

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动应用
python app.py

# 3. 浏览器访问
open http://127.0.0.1:5000
```

> 需要 Python 3.9+。首次运行时会自动创建 `data/`（题库）和 `data/uploads/`（上传文件）目录。

---

## 🐳 部署到云端

本项目同时支持 **4 种部署方式**，任选其一：

| 平台 | 配置文件 | 一键部署 | 免费额度 |
|------|---------|---------|---------|
| **Render** | `render.yaml` | ✅ | 512 MB / 休眠 |
| **Railway** | `Dockerfile` + `Procfile` | ✅ | $5/月额度 |
| **Vercel** | `vercel.json` | ✅ | Serverless 适配 |
| **Docker** | `Dockerfile` | ✅ | 自托管 |

详细 Railway 部署步骤参见 [`DEPLOY_RAILWAY.md`](./DEPLOY_RAILWAY.md)。

---

## 📁 项目结构

```
dakika/
├── app.py                    # Flask 主程序（所有 API + 文本提取逻辑）
├── supabase_client.py        # （可选）外部数据客户端
├── requirements.txt          # Python 依赖清单
├── runtime.txt               # Python 版本声明
├── Procfile                  # Railway/Heroku 进程定义
├── Dockerfile                # Docker 构建
├── render.yaml               # Render 部署配置
├── vercel.json               # Vercel 部署配置
├── DEPLOY_RAILWAY.md         # Railway 部署指南
├── start.bat                 # Windows 一键启动脚本
├── templates/
│   └── index.html            # 主页面（单页 SPA）
└── data/                     # 运行时自动生成（题库、上传文件）
```

---

## 🔧 配置项（环境变量）

应用**开箱即用，无需任何环境变量**。可选变量：

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `DATA_DIR` | `<项目目录>/data` | 数据存储根目录 |
| `PORT` | `5000` | HTTP 端口（云平台自动注入） |
| `RAILWAY_ENVIRONMENT` / `RENDER` | — | 自动检测，自动切到 `/tmp/data` |

---

## 📚 支持的文件格式

| 类别 | 格式 |
|------|------|
| **Word** | `.doc` `.docx` |
| **PDF** | `.pdf` |
| **Excel** | `.xlsx` `.xlsm` `.xls` |
| **PowerPoint** | `.pptx` |
| **文本** | `.txt` `.md` `.json` `.csv` `.xml` `.log` `.yaml` `.yml` `.rtf` `.html` `.htm` |
| **图片** | `.jpg` `.jpeg` `.png` `.bmp` `.gif` |

> `.doc`（旧版 Word）在 Windows 上自动调用 Word/WPS COM 解析；在 Linux 上回退到二进制文本提取。

---

## 🛡️ 已主动移除的"过度设计"

为保证"开箱即用 + 零外部依赖"，本项目刻意剔除了：

- ❌ 登录系统 / Session / 用户隔离目录
- ❌ OpenAI / DeepSeek / Qwen / Claude / Gemini / 自定义大模型 API
- ❌ DuckDuckGo 联网搜题 / 网页抓取
- ❌ Supabase / 任何外部数据库

---

## 📜 License

MIT