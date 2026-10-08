# KKS 编码审核智能体

这是一个不依赖 WorkBuddy 的 KKS Excel 审核代码服务。它把附件中的 `kks-audit` Skill 固化为可重复运行的审核引擎：编码人员可以通过浏览器上传 `.xlsx`，输出面向电厂业务人员的《KKS 编码质量审核报告》和精简问题 Excel。AI 只作为不确定项的辅助复核，正式报告保留结论、证据和整改建议。

## 用户需求与附件规则的边界

- 用户需求：编码人员完成编码后，自行上传 Excel，得到可读的审核结论和问题清单。
- 附件 Skill：是本项目的正式审核规则来源，规则会在运行时执行并写入覆盖证明；它不授权自动修改源 Excel。
- 本目录：可运行的审核智能体；`source-skill/` 只保留附件原始副本，便于核对版本，不作为“未实现”的替代品。

当前实现边界：`kks-audit/SKILL.md` 是正式 Skill 规范，`kks-audit/scripts/audit_template.py` 会被运行器动态加载并执行；AI 按当前规则选择相关 Skill 内容，并结合 Excel 证据复核，参考资料不得脱离当前 Excel 证据单独定性。专业覆盖、编码深度和数量瀑布对比通过 `--compare-file` 执行；DM8/LOCATIONS 目前输出导入治理约束，但未连接真实数据库做导入验证。

## 方式一：启动网页服务（推荐）

在本目录执行：

```powershell
python .\server.py --host 127.0.0.1 --port 8080 --runs-dir .\runs
```

浏览器打开 `http://服务器地址:8080/`，上传 Excel 即可。默认只监听本机（`127.0.0.1`）；要让局域网同事访问，把 `--host` 改成 `0.0.0.0`（Docker 部署对应 `.env` 里的 `KKS_BIND_ADDR`），并务必先配置下面的管理令牌。

别人使用时不需要 WorkBuddy，只需要能访问这台服务器的地址。公开接口（无需令牌）：

- `GET /`：浏览器上传页面；
- `GET /healthz`：健康检查；
- `POST /api/audits`：接收上传并创建后台审核任务，支持 multipart 字段 `file`，返回 `run_id`；
- `GET /api/audits/{run_id}/status`：查询审核阶段、真实进度和 AI 全部候选复核状态；
- `GET /api/audits/{run_id}/{文件名}`：按上传 Excel 文件名生成的 HTML 报告或 Excel 问题清单。

管理接口（需要管理令牌，携带方式 `Authorization: Bearer <token>` 或 `?token=<token>`；未配置 `KKS_ADMIN_TOKEN` 时一律返回 403）：

- `GET/POST /api/config`：读取/保存 AI 配置；
- `GET/POST /api/skill`、`POST /api/skill/upload`：读取/保存/上传 Skill；
- `GET /api/logs`：运行日志；
- `POST /api/ai/test`：AI 连接测试；
- `GET /api/history`、`GET /api/history/download`：审计台账与汇总报告/JSON 下载。响应含原始文件名、内容哈希、审核时间和 P0/P1/P2 统计，属审核记录，因此同样纳入鉴权，不因页面展示需要而公开。
- `GET /api/history/artifact?run_id=<id>&kind=html|xlsx`：按台账记录下载该次审核生成的报告 HTML 或问题清单 Excel。台账页「报告」列用的就是这个接口；`run_id` 解析后必须仍在 runs 根目录内，产物已被保留策略清理时返回 404（`artifact_missing`）。

> 注意区分：业务人员拿到的报告直链 `GET /api/audits/{run_id}/{文件名}` 仍是**公开**的（面向电厂业务人员，不要求令牌），本次改动没有锁住它。台账页的导出走上面这个带令牌的接口。

## 统一依赖环境（uv）

项目统一使用 `uv` 管理 Python、运行依赖和 EXE 构建依赖。不要再分别使用系统 Python、pip 和另一套虚拟环境。

首次初始化：

```powershell
cd D:\项目\同海\kks-audit-agent
$env:UV_CACHE_DIR = Join-Path (Get-Location) ".uv-cache"
uv sync --group build
```

日常运行：

```powershell
$env:UV_CACHE_DIR = Join-Path (Get-Location) ".uv-cache"
uv run python .\server.py --host 0.0.0.0 --port 8080
```

命令行审核：

```powershell
$env:UV_CACHE_DIR = Join-Path (Get-Location) ".uv-cache"
uv run python .\run_audit.py "D:\QQ\8.20_舟山汽机专业.xlsx" --output-dir .\outputs\sample

# 可选：同时对比另一份厂/专业 Excel 的覆盖、深度和数量闭合差异
uv run python .\run_audit.py "D:\QQ\8.20_舟山汽机专业.xlsx" --compare-file "D:\QQ\另一份编码.xlsx" --output-dir .\outputs\compare
```

依赖定义在 `pyproject.toml`，锁定版本在 `uv.lock`，统一虚拟环境为项目目录下的 `.venv`。

## AI 配置边界

AI 地址、模型和密钥由管理员预先写入 `config/app_config.json`，打包时随 EXE 一起发布。网页左侧的“AI 配置”是管理员配置页，可修改接口地址、模型、启用状态和 API 密钥，并可直接测试连接。密钥输入框不会回显原文，只显示掩码；留空保存表示保持原密钥。

普通使用者只需要上传 Excel、查看报告和日志。管理接口（配置读写、Skill 读写上传、运行日志、AI 连接测试、审计台账与下载）统一由 `KKS_ADMIN_TOKEN` 保护：未配置该变量时这些接口一律返回 403（fail-closed），令牌缺失或不匹配返回 401，比较使用常量时间避免时序侧信道。令牌可用 `openssl rand -hex 24` 生成，写入 `.env` 后重启服务；浏览器端在「AI 配置」页填入同一令牌即可（保存在 localStorage，页面自动携带）。

服务默认只监听回环地址，不发布到所有宿主网卡；需要局域网访问时显式设置绑定地址（Docker 用 `KKS_BIND_ADDR`），并保证管理令牌已配置。API 密钥不要写入 Skill、Excel、报告或 Git。

## 审核产物的保留策略

每次审核会在 `runs/<run_id>/` 留下三份东西：上传副本、报告 HTML、问题清单 Excel，并在台账里记下 `run_id` 作为索引。

- 保留天数由 `KKS_ARTIFACT_RETENTION_DAYS` 控制，默认 **90 天**，设 `0` 表示不自动清理。
- 每次审核结束后按 `run_id` 里的开始时间判断，超期目录整目录删除，并把台账对应记录标记为 `archived`——**台账记录本身不会被删**，只是台账页「报告」列显示「已归档」，不能再在线下载。
- 只处理名字严格符合 `run_id` 格式（`YYYYmmddTHHMMSS_xxxxxxxx`）的目录，`runs/` 下的其他目录（回归样本、手工目录、日志）一律不碰。
- 升级到带 `run_id` 的版本后，服务启动时会按**输入文件内容哈希**给老记录回填一次 `run_id`（精确匹配、幂等）；匹配不上的老记录显示「—」，需要时请用 `runs/` 里的其他副本对照。

## 配置、Skill 和日志

启动网页服务后，使用者可以：

- 上传 Excel 并下载审核报告；
- 查看规则审核、AI 复核和报告生成的实时进度；
- 查看当前已加载的正式 Skill、审核参考文档和实际执行引擎；管理员可在页面编辑或上传允许的 Skill 文件；
- 查看最近运行日志。

AI 配置可通过管理员页面维护，也可直接编辑 `config/app_config.json`；日志保存在 `logs/kks-audit.log`，审核运行产物保存在 `runs/`。

AI 的三个附加能力用环境变量（或 `config/app_config.json` 的 `ai` 段）单独控制，默认都开启：

| 变量 | 作用 | 关闭时 |
|---|---|---|
| `AI_SUMMARY_ENABLED` | 单文件「AI 总体总结」 | 报告第九节不生成 |
| `AI_TREND_SUMMARY_ENABLED` | 台账「AI 跨文件趋势总结」 | 汇总报告不出趋势节 |
| `AI_CODE_SUGGEST_ENABLED` | **问题行的「AI 建议编码」** | 不生成，报告不出现该列 |
| `AI_CODE_SUGGEST_MAX` | 送生成的行数上限（`0` = 不限制） | 超限时按 P0>P1>P2 截断 |

「AI 建议编码」只对**判定为有问题的行**生成，命中规则的**标准条款原文**会随 prompt 一起送入，
生成结果再由本地复核（12 位骨架 / 禁用 `I`、`O` / 分类码必须落在两套标准字典内），
不通过的不写进报告。报告第十节与 Excel 的「AI建议编码」sheet 是它的落点，仅供参考，不替代规则结论。

## 规则依据与一致性核查

**规则依据只有两份，且都必须是现行有效文档**（登记在 `config/kks_rules.json` 的 `standard_basis`）：

| 文件 | 标准 | 状态 |
|---|---|---|
| `国标/GBT 50549-2020 电厂标识系统编码标准.pdf_by_PaddleOCR-VL-1.6.md` | **GB/T 50549-2020**《电厂标识系统编码标准》 | 现行国家标准，2021-03-01 实施；同时废止 GB/T 50549-2010 |
| `国标/中国中煤能源集团电厂标识系统编码应用细则（试行）A3.2.docx` | 中煤集团《电厂标识系统编码应用细则（试行）》 | 集团落地细则，附录作分类码补充 |

⚠️ **版本纪律**：GB/T 50549-2020 前言明确**同时废止 GB/T 50549-2010**。本项目**不得**再引用
继承 2010 体系的资料（2015 年《火电厂标识系统（KKS）编码导则》、旧版风电/光伏规范）——
同一分类码在两套体系里含义不同（如 `GP`：2020 版是泵，2010 版体系是照明检修配电箱），
混用会让判定与"建议编码"指向完全不同的设备。这些资料已连同其字典、分歧清单、KKS-36 一并移除。

结构上两套体系其实一致：项目沿用的 `G|F0 | F1F2F3 | F4F5 | A1A2 | AN1AN2AN3`
就是 GB/T 50549-2020 4.4.1 功能面标识三级编码（火电项目全厂码 H1 缺省，2+3+2+2+3=12 位）。

规则落地在 `config/kks_rules.json`、`run_audit.py`、`kks-audit/scripts/audit_template.py`、
`SKILL.md` 和 `config/kks_code_index.json` 五个地方。**改任何规则之后，先跑一次核查**：

```bash
python -B tools/check_rule_conflicts.py      # 退出码 0 = 无冲突；1 = 有 ERROR
python -B -m pytest tests/test_rule_conflicts.py -q
```

它检查五类问题，其中两类是这次专门为"版本纪律"加的：

| 类别 | 检查内容 |
|---|---|
| A 注册一致性 | 规则号在「定义 / 出处 / 代码 / SKILL」四处对齐；无幽灵规则、无"声明未实现" |
| B 配置自洽 | 配置常量与代码常量一致；骨架位数与分段位置一致；编号段区间不重叠；特殊码在字典内 |
| **C 版本纪律** | 依据里只许出现现行标准；索引里不得残留旧体系字典；文档不得把废止版本当依据 |
| **D 条款真实性** | 引用的条款号/表号必须在抽取出的标准语料里**真实存在**（防止虚构出处） |
| E 同条件重复 | 同一行被两条规则以不同等级判定时，必须已登记为互补关系 |

Skill 管理页支持编辑/上传 `SKILL.md`、`references/*.md` 和 `scripts/audit_template.py`，也支持上传只包含这些文件的 `.zip` 包。每次保存前会备份到 `skill_backups/`；修改 `audit_template.py` 后需要重启服务，修改 Markdown 后会在后续 AI 复核中使用。

## 打包为 Windows EXE

在有 Python 的开发机上执行：

```powershell
cd D:\项目\同海\kks-audit-agent
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1 -Clean
```

脚本内部会执行 `uv sync --group build`，再用同一个 `.venv` 调用 PyInstaller。

产物位于 `dist\KKS-Audit-Agent\KKS-Audit-Agent.exe`。把整个 `dist\KKS-Audit-Agent` 文件夹交给使用者，双击 `start-kks-audit.cmd` 即可启动并打开浏览器。使用者不需要 WorkBuddy，也不需要安装 Python；Skill、AI 配置和日志都在 EXE 同目录下维护。

也可以用命令行上传：

```powershell
curl.exe -F "file=@D:\QQ\8.20_舟山汽机专业.xlsx" http://127.0.0.1:8080/api/audits
```

## 管理员命令行运行

在项目 uv 环境中运行：

```powershell
$env:UV_CACHE_DIR = Join-Path (Get-Location) ".uv-cache"
uv run python .\run_audit.py "D:\QQ\8.20_舟山汽机专业.xlsx" --output-dir .\outputs\8.20_舟山汽机专业
```

输出：

- `{原文件名}_审核报告.html`：按“总体结论、关键质量指标、KKS 八维审核结果、P0/P1/P2 分析、整改建议、导入评估、详细问题清单”组织的正式报告；
- `{原文件名}_问题清单.xlsx`：只展示“等级、规则、行号、KKS、问题、整改建议”六列；AI 依据保留在隐藏的“AI复核（技术）”页，规则覆盖和结构证据保留在辅助页。

## 审核边界

- 原始 Excel 只读，不回写、不自动改码、不自动发布。
- 变长 KKS 层级按“父级列 + 子码前缀”校验，不按固定长度截断。
- 13 位 A/B 位置扩展码进入“需人工确认”，不直接判为坏码。
- 名称相似度只可作为人工抽检线索，不能自动判定同一设备。
- 规则命中与语义复核分开记录，公用/联络设备不因名称中出现另一机组号而误报。
- `valid_letter_pairs`、`letter_reallocations` 等项目专属字典在 `config/kks_rules.json` 中配置；Skill 未提供具体项目表时保持空表，不擅自编造迁移关系。
- “可导入”结论仅代表静态审核通过；未连接 DM8/LOCATIONS 做真实导入验证。
- AI 复核所有被规则标记为 `needs_review` 或语义候选的问题；不会自动关闭问题、改码或发布。
- EXE 使用“目录版”发布：管理员通过 AI 配置页维护接口和 Skill，使用者上传文件、查看全量问题清单和日志；当前没有内置登录认证，建议仅在本机或受控内网使用；不要把真实 API 密钥提交到 Git。

## 代码调用方式

程序可以直接调用 `run_audit.audit_file(input_path, output_dir)`，不需要启动网页服务。网页服务只是对这个函数增加了上传和下载接口。

## 部署建议

- 少量内部用户：在一台内网 Windows 电脑启动服务，浏览器共享地址。
- 多人长期使用：部署到内网服务器，并在反向代理或网关层增加登录、HTTPS、上传大小限制和运行目录清理策略。
- 生产环境：不要直接暴露到公网；当前版本没有用户登录、权限管理和自动清理历史文件。
