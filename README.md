# PPT Light (pptmo)

一个轻量的本地 PPTX 编辑器，核心是**对话式、流式输出、AI 主导的精准改稿**：AI 会先读大纲、读页面、查表格和对象，再提出可审阅的精确操作；用户可以逐条确认，也可以打开“自动执行”直接写回 PPTX。

前端不依赖 React 或构建工具，后端使用 FastAPI。预览由 LibreOffice 转 PDF、PyMuPDF 渲染；修改通过 `python-pptx` 直接写回原文件，下载得到的是真正可编辑的 `.pptx`。

> 本仓库同时是一个 agent skill：根目录的 `SKILL.md` 让 AI 助手（ZCode / Claude Code 等）可以按流程部署、调用 API、配置远程访问。`scripts/setup.sh` 一键装环境。

## 快速开始

### 依赖

- Python 3.10+（在 3.12 与 3.14 上验证过）
- LibreOffice Impress / `soffice`（渲染需要）
- 可访问 OpenAI 兼容网关（AI 改稿需要；没有也能用手动编辑、版本、下载）

```bash
# 一键部署到任意目录（可选：直接带上 AI 配置）
bash scripts/setup.sh ~/apps/pptmo \
  --api-key <你的key> \
  --api-base https://open.bigmodel.cn/api/paas/v4 \
  --model glm-4.6

# 或手动：
cd backend
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8090
```

打开 http://127.0.0.1:8090 ，上传一个 `.pptx` 即可开始。健康检查：`curl http://127.0.0.1:8090/api/health`。

### 环境配置

复制 `backend/.env.example` 为 `backend/.env`（权限建议 600），填入任意 OpenAI 兼容服务：

```dotenv
DEEPSEEK_API_KEY=你的 key
DEEPSEEK_API_BASE=https://api.deepseek.com/v1   # 或智谱 https://open.bigmodel.cn/api/paas/v4
DEEPSEEK_MODEL=deepseek-chat                     # 或 glm-4.6 / glm-5.3-flash 等
AI_PROVIDER=deepseek
RENDER_SCALE=2
```

可选：`MAX_UPLOAD_MB=60`、`AI_MAX_STEPS=10`（2–16）、`AI_MAX_OPS=24`（4–60）。变量名带 DEEPSEEK 只是历史原因，实际接受任何兼容网关（DeepSeek、智谱 bigmodel、自建 LiteLLM 等）。前端永远不会收到 key。

## 核心流程

1. 在首页上传或打开一个 `.pptx` 会话。
2. 在右侧对话框说清页码、对象和期望效果，例如：
   `只改第 4 页：把项目名称里的“和”改成“与”并加粗；把“200 万（原）→200 万”改成“180 万（原）→180 万”，保留原格式。`
3. AI 以 ReAct 工具循环执行：读大纲 → 读页面 → 读表格/查对象 → 提交操作 → 总结。
4. 默认进入**审阅模式**：操作卡片默认全选，点“应用并生成新版本”后才写入。
5. 打开“自动执行”后，AI 完成后直接应用并生成新版本；SSE 会持续推送思考、工具步骤、写回与渲染状态。

## 功能

### 1. 对话式 AI Harness

- **真实流式输出**：SSE 推送 `start → step → token/thought/tool_call/tool_result → operations → summary → done`。
- **AI 主导、用户可控**：AI 不能直接改文件，只能调用工具并提交操作；是否自动执行由用户决定。
- **会话上下文**：同一文档最近 12 条对话带入下一轮；指令里出现“第 5 页”“首页”等自动定位。
- **待确认操作可恢复**：审阅模式下刷新页面，最近一组未应用操作恢复为待确认卡片。
- **冲突保护**：请求带 `base_version`，应用时服务端校验版本。

可用工具：`outline` / `read_slides`（单次≤4 页）/ `search_text` / `read_table` / `inspect_shapes` / `propose_operations`（单次≤6 条）/ `finish`。

### 2. 复杂 PPTX 的对象级控制

- 递归解析组合形状：组内文字可精准修改，几何操作被拒绝。
- 表格单元格拆成独立对象（ID 形如 `s3_id4_c0_1`），支持文字/样式/清空；单元格移动/删除被拒绝。
- 操作必须命中真实 `shape_id`，文本修改校验 `before` 与当前原文，避免“改错行”。
- 页码、页脚、logo 默认受保护；指令明确提到时才允许修改。
- 写入前统一重新渲染受影响页面，版本与预览同步更新。

### 3. 直接编辑

缩略图翻页、双击文字原位编辑、拖动/缩放文本框、检查器改字号/颜色/对齐/加粗、新增/删除文本框、版本历史可恢复、下载真实可编辑 PPTX。

### 4. 会话与版本

每次上传创建会话；`#doc=<id>` 刷新可回；会话记录含指令、回复、执行轨迹、操作与版本事件；服务重启后仍存在。

## API 摘要

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/api/health` | 健康检查、AI/LO 是否就绪、会话数 |
| `GET` | `/api/sessions` | 最近会话列表 |
| `POST` | `/api/docs/upload` | 上传 `.pptx` 并创建会话 |
| `GET` | `/api/docs/{id}` | 文档 + 对象 + 最近会话记录 |
| `PATCH` | `/api/docs/{id}` | 重命名会话 |
| `DELETE` | `/api/docs/{id}` | 删除会话及全部文件 |
| `POST` | `/api/docs/{id}/chat` | 流式对话改稿（SSE） |
| `POST` | `/api/docs/{id}/apply` | 校验并应用操作，生成新版本 |
| `GET` | `/api/docs/{id}/messages` | 会话记录 |
| `GET` | `/api/docs/{id}/versions` | 版本列表 |
| `POST` | `/api/docs/{id}/versions/{v}/restore` | 恢复到指定版本 |
| `GET` | `/api/docs/{id}/download` | 下载当前 PPTX |
| `GET` | `/api/docs/{id}/slides/{n}/image` | 页面渲染图 / 缩略图 |

`/chat` 请求示例与 SSE 事件说明见交互式文档 `/api/docs-ui`。

## 目录结构

```text
├── SKILL.md               # agent skill 入口（部署/调用/远程访问指引）
├── scripts/setup.sh       # 一键部署脚本
└── app/                   # 应用本体
    ├── backend/
    │   ├── main.py                 # FastAPI 路由、SSE 对话、静态托管
    │   ├── storage.py              # 文档/版本/会话/渲染文件布局与锁
    │   └── services/
    │       ├── parse.py            # 递归解析组合、表格单元格、文本对象、bbox
    │       ├── edits.py            # 操作校验与 python-pptx 写回
    │       ├── render.py           # LibreOffice → PDF → PNG/缩略图
    │       ├── agent.py            # 流式 ReAct 循环、工具、操作校验
    │       ├── streaming.py        # OpenAI 兼容接口的 SSE 流式调用
    │       └── ai.py               # AI 服务配置与通用文本工具
    ├── web/                        # 无构建 Vanilla JS 前端
    └── samples/demo.pptx           # 演示文件
```

数据落盘在 `app/backend/data/docs/<doc_id>/`：`meta.json`、`objects.json`、`messages.json`、`current.pptx`、`versions/`、`render/vN/`。每个历史版本保留渲染图，长期编辑后 `render/` 会增长，可按需清理旧版本目录。

## 已知边界

- 只支持文字层与文本框几何编辑；图片、图表、SmartArt 以渲染图展示，不能在网页里改。
- 表格单元格支持文字与样式，不支持单独移动/缩放/删除。
- 分组内文字只能改文字或样式。
- 渲染为同步任务，超大文稿首次打开需要数秒；单文件限制 60 MB（服务端下发，前端自动跟随）。
- AI 改稿依赖外部 OpenAI 兼容 API；网关不可达时手动编辑、版本、下载仍可用。

## 安全模型（v0.1.2 起）

应用内置 `RequestGuardMiddleware`，即使部署在受信网络也建议保留：

- **DNS rebinding 防护**：Host 不在白名单（回环 + `TRUSTED_HOSTS`）→ 421
- **浏览器 drive-by 防护**：跨站 POST/PATCH/DELETE（`Sec-Fetch-Site: cross-site` 或异源 `Origin`）→ 403
- **请求体限幅**：JSON 默认 2 MB、multipart 按上传上限 → 超限 413，防 OOM
- **参数上限**：patch 字符串 20k、解析 500 页 / 5000 对象 / 20 层组、渲染 500 页、版本保留 50 个、渲染队列 4、并发对话 2
- **错误脱敏**：服务器文件路径不出现在响应里
- **自动执行守卫**：删除/清空类操作默认不自动执行（`AUTO_APPLY_DESTRUCTIVE=1` 才放行）

通过反向代理（如 `tailscale serve`）访问时，把公网主机名加进 `TRUSTED_HOSTS` / `TRUSTED_ORIGINS`，否则浏览器请求会被 403。没有登录系统：请保持回环绑定或私有网络暴露，不要直接上公网。

## License

MIT
