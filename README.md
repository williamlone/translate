# Access 博客文章中英文翻译程序

本程序从 Access 数据库的源表中读取中文标题、摘要和正文，翻译成英文后写入同一数据库的目标表。数据库路径和 API Key 等敏感信息放在本地配置文件中；数据表名和固定术语不属于敏感信息，在代码和文档中直接保留。

---

## 1. 功能特点

1. **同一数据库内生成目标表**
   - 源表为 `blog_Article`，目标表为 `blog_Article_Trans`（表名由配置 `SOURCE_TABLE` / `TARGET_TABLE` 指定）。
   - 目标表复制源表的字段、类型和索引，并额外增加 5 个翻译状态字段；源表不会被修改。
   - 英文结果写入待翻译字段，同时同步写入对应的 `_EN` 字段（若存在）。

2. **自动续跑**
   - 按主键升序查找第一条未翻译、失败或源文已变化的记录。
   - `--limit N` 表示从该记录开始处理接下来的 N 篇文章。
   - 已成功翻译且源文未变化的记录不会重复翻译。

3. **准确记录进度**
   - 每条记录的状态写入目标表的 `log_TransStatus`、`log_TransModel`、`log_TransError`、`log_TransTime`、`log_TransHash`。
   - 每次运行的总览写入进度表 `blog_Article_Trans_Progress`（由配置 `PROGRESS_TABLE` 指定）。

4. **保护 HTML 与专有内容**
   - HTML 标签、属性、实体和 URL 由提示词要求原样保留；`script`/`style`/`pre`/`code` 等不可翻译整块用占位符保护。
   - 固定术语内置在程序中，也可以通过术语表文件覆盖或扩展。

5. **支持多种翻译引擎**
   - `qwen-mt-lite` / `qwen-mt-plus`：阿里云百炼
   - `doubao-seed-translation`：火山方舟豆包翻译模型
   - `deepseek`：DeepSeek 官方或火山方舟 OpenAI 兼容接口

---

## 2. 文件说明

| 文件 | 是否提交到 GitHub | 说明 |
|---|---|---|
| `translate_blog.py` | 是 | 主程序 |
| `requirements.txt` | 是 | Python 依赖清单 |
| `config.example.env` | 是 | 配置文件模板，敏感项为占位符 |
| `.gitignore` | 是 | 忽略本地配置和数据库文件 |
| `translate.env` | 否 | 本机真实配置，含数据库路径等敏感信息 |
| `*.mdb` / `*.accdb` | 否 | Access 数据库文件本身 |
| `_inspect.py` | 是 | 进度检查辅助脚本，从配置读取连接信息 |

---

## 3. 环境要求

- Windows 操作系统
- Python 3.9 或更高版本
- Microsoft Access ODBC 驱动：`Microsoft Access Driver (*.mdb, *.accdb)`（驱动名可在配置中覆盖）
- Python 依赖：`pyodbc>=5.1.0`

确认驱动是否可用：

```powershell
python -c "import pyodbc; print('\n'.join(pyodbc.drivers()))"
```

---

## 4. 安装依赖

```powershell
python -m pip install -r requirements.txt
```

---

## 5. 配置（重要）

复制配置模板并填写真实值：

```powershell
Copy-Item config.example.env translate.env
```

`translate.env` 已被 `.gitignore` 忽略，**只有模板会被提交**。请不要把真实数据库路径或 API Key 写进 `README.md`、代码或模板文件；数据表名和固定术语直接写在代码中。

程序启动时会自动读取脚本同目录下的 `translate.env`。已存在的系统环境变量优先级高于配置文件，因此也可以只在终端里临时设置密钥。

### 5.1 配置项说明

| 配置项 | 说明 |
|---|---|
| `BLOG_DB` | Access 数据库文件路径。相对路径按脚本所在目录解析 |
| `SOURCE_TABLE` | 源表名：`blog_Article` |
| `TARGET_TABLE` | 目标表名：`blog_Article_Trans` |
| `PROGRESS_TABLE` | 进度表名：`blog_Article_Trans_Progress` |
| `ACCESS_ODBC_DRIVER` | ODBC 驱动名，默认 `Microsoft Access Driver (*.mdb, *.accdb)` |
| `TRANSLATE_ENGINE` | 默认翻译引擎 |
| `GLOSSARY_FILE` | 术语表 JSON 路径，可留空 |
| `INSPECT_ENGINE` | `_inspect.py` 默认查询的引擎名 |
| `DEEPSEEK_API_KEY` / `DEEPSEEK_MODEL` / `DEEPSEEK_BASE_URL` | DeepSeek 官方 API |
| `ARK_API_KEY` | 火山方舟密钥，DeepSeek 与豆包通用/专用模型共用 |
| `DOUBAO_MODEL` / `DOUBAO_BASE_URL` | 豆包翻译专用模型 |
| `DASHSCOPE_API_KEY` / `QWEN_MODEL` / `QWEN_BASE_URL` | 阿里云百炼 Qwen-MT |
| `DEEPSEEK_MAX_TOKENS` | 单次请求最大输出 token 数 |
| `DEEPSEEK_DISABLE_THINKING` | `1` 表示关闭思考/推理模式 |

API Key 可以写在 `translate.env` 里（不要提交），也可以用环境变量设置，例如：

```powershell
$env:DEEPSEEK_API_KEY="你的API Key"
```

`--engine deepseek` 会根据 `DEEPSEEK_BASE_URL` 自动选择密钥：官方地址用 `DEEPSEEK_API_KEY`，火山方舟地址用 `ARK_API_KEY`。

### 5.2 术语表

程序内置了以下固定术语。如需覆盖或扩展，可通过 `--glossary` 或 `GLOSSARY_FILE` 指定自己的 JSON 文件：

```json
{
  "月光博客": "Moonlight Blog",
  "Google Earth": "Google Earth",
  "Web 2.0": "Web 2.0",
  "SEO": "SEO",
  "QQ": "QQ",
  "MSN": "MSN"
}
```

---

## 6. 快速开始

以下命令均在项目目录下执行。

先看将要翻译哪些记录（不调用 API、不修改数据库）：

```powershell
python translate_blog.py --dry-run --limit 10
```

正式翻译：

```powershell
python translate_blog.py --limit 10
```

指定引擎：

```powershell
python translate_blog.py --engine qwen-mt-lite --limit 10
python translate_blog.py --engine deepseek --limit 10
python translate_blog.py --engine doubao-seed-translation --limit 10
```

使用火山方舟的 OpenAI 兼容接口时，把地址设为配置中的 `DEEPSEEK_BASE_URL`，模型名按需填写；程序会自动使用 `ARK_API_KEY`。方舟账户需要先在控制台开通对应模型，否则会返回 `HTTP 404 ModelNotOpen`。

查看进度记录：

```powershell
python _inspect.py
```

---

## 7. 命令行参数

未显式传入时，以下参数取配置文件中的值；表中“来自配置”表示没有代码内置默认值。

| 参数 | 默认值 | 说明 |
|---|---:|---|
| `--db` | 来自配置 `BLOG_DB` | Access 数据库路径 |
| `--source-table` | 来自配置 `SOURCE_TABLE` | 源表名：`blog_Article` |
| `--target-table` | 来自配置 `TARGET_TABLE` | 目标表名：`blog_Article_Trans` |
| `--progress-table` | 来自配置 `PROGRESS_TABLE` | 进度表名：`blog_Article_Trans_Progress` |
| `--engine` | 来自配置 `TRANSLATE_ENGINE` | 翻译引擎 |
| `--limit` | `10` | 本次翻译的文章数量 |
| `--workers` | `4` | 单篇文章内并行翻译的块数 |
| `--article-workers` | `1` | 单批次内并行处理的文章数 |
| `--chunk-chars` | `600` | 单次 API 请求的最大字符数；HTML 密集文章可调到 400 |
| `--timeout` | `120` | API 请求超时秒数 |
| `--retries` | `4` | API 失败重试次数 |
| `--min-interval` | `0` | 两次 API 请求之间的最小间隔秒数 |
| `--env-file` | 自动查找 `translate.env` | 指定其他配置文件 |
| `--glossary` | 来自配置 `GLOSSARY_FILE` | 自定义术语 JSON 文件 |
| `--dry-run` | 关闭 | 只显示待翻译记录，不调用 API |
| `--init-only` | 关闭 | 只建表并引导已有译文，不调用 API |
| `--qwen-model` / `--qwen-base-url` | 来自配置 | Qwen 模型与接口 |
| `--doubao-model` / `--doubao-base-url` | 来自配置 | 豆包翻译模型与接口 |
| `--deepseek-model` / `--deepseek-base-url` | 来自配置 | DeepSeek / 方舟模型与接口 |
| `--deepseek-max-tokens` | 来自配置 | 输出 token 上限 |
| `--deepseek-disable-thinking` | 来自配置 | 关闭推理模式 |

---

## 8. 进度与续跑机制

### 8.1 源文哈希

程序对待翻译的三个字段计算 SHA-256，结果写入目标表 `blog_Article_Trans` 的 `log_TransHash`。源文修改后哈希变化，该记录会在下一轮被重新处理。

### 8.2 单条记录状态

目标表 `blog_Article_Trans` 使用以下字段记录单篇状态：

| 字段 | 含义 |
|---|---|
| `log_TransStatus` | `ok` 或 `failed` |
| `log_TransModel` | 使用的翻译引擎或模型 |
| `log_TransError` | 失败原因 |
| `log_TransTime` | 最后处理时间 |
| `log_TransHash` | 源文哈希 |

### 8.3 每次运行记录

进度表 `blog_Article_Trans_Progress` 记录 `run_id`、引擎、起止记录 ID、请求/成功/失败数量、开始与结束时间、状态、摘要以及最后一条成功记录的信息。续跑依据每条记录的状态和哈希，而不是只看最后一条成功记录，因此中间失败的条目会优先重试，后续已成功的条目不会重复翻译。

---

## 9. 输出示例

```text
Pending 10 record(s), log_ID: ...
[OK] log_ID=... (1.8s) -> Example Translated Title
...
deepseek: requested=10, succeeded=10, failed=0
```

退出码：

| 退出码 | 含义 |
|---:|---|
| `0` | 全部成功，或没有待翻译记录 |
| `1` | 程序级错误，例如缺少配置/密钥、数据库错误 |
| `2` | 本批存在翻译失败记录 |

---

## 10. 提交 GitHub 前的检查

1. 确认 `.gitignore` 生效：`git status --short` 中不应出现 `translate.env` 或数据库文件。
2. 代码、数据表名及配置模板都可以提交。
3. 若曾经误提交过真实配置或数据库，需要从 Git 历史中移除，并立即轮换相应 API Key。

---

## 11. 常见问题

### 提示缺少必要配置

说明 `translate.env` 不存在或没有填写 `BLOG_DB`、`SOURCE_TABLE`、`TARGET_TABLE`、`PROGRESS_TABLE`。复制模板并填写即可。

### 新数据库提示 source table lacks required columns

程序只要求源表包含 `log_ID` 和三个待翻译字段，不会向源表添加任何字段。5 个翻译状态字段只建在目标表中；如果源表不存在，程序会创建仅含必要字段的源表。目标表与进度表仍自动创建，无需手工建表。

### 提示缺少 ODBC 驱动

确认已安装 Microsoft Access Database Engine，且 `pyodbc.drivers()` 中包含配置里指定的驱动名。驱动名不同时修改 `ACCESS_ODBC_DRIVER`。

### 提示缺少 API Key

在 `translate.env` 或环境变量中设置对应引擎的密钥。程序会根据接口地址自动选择 DeepSeek 官方密钥或火山方舟密钥。

### 一直从同一条记录开始

检查该记录在 `blog_Article_Trans` 中的 `log_TransStatus` 是否为 `failed`，并查看 `log_TransError` 定位失败原因。

### 译文为空或提示 dropped placeholder tokens

使用推理模型时，思考过程可能占用全部输出 token。设置 `DEEPSEEK_DISABLE_THINKING=1`，并按需增大 `DEEPSEEK_MAX_TOKENS`。程序会校验占位符完整性，缺失时最多重试 5 次，并在重试请求中明确指出遗漏的占位符。若 HTML/URL 非常密集仍失败，可把 `--chunk-chars` 降到 400 或更低。

### 出现 429 限流

增大请求间隔：

```powershell
python translate_blog.py --limit 10 --min-interval 0.5 --retries 6
```
