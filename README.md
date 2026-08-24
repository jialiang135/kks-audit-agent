# KKS 编码审核智能体

这是一个不依赖 WorkBuddy 的 KKS Excel 审核代码服务。它把附件中的 `kks-audit` Skill 固化为可重复运行的审核引擎：编码人员可以通过浏览器上传 `.xlsx`，输出 HTML 审核报告和 Excel 问题清单。配置模型接口后，它会对全部语义候选项执行一次 AI 二次复核。

## 用户需求与附件规则的边界

- 用户需求：编码人员完成编码后，自行上传 Excel，得到可读的审核结论和问题清单。
- 附件 Skill：是本项目的正式审核规则来源，规则会在运行时执行并写入覆盖证明；它不授权自动修改源 Excel。
- 本目录：可运行的审核智能体；`source-skill/` 只保留附件原始副本，便于核对版本，不作为“未实现”的替代品。

当前实现边界：`kks-audit/SKILL.md` 是正式 Skill 规范，`kks-audit/scripts/audit_template.py` 会被运行器动态加载并执行；AI 会加载正式 Skill 和 `references/*.md` 作为上下文，但参考资料不得脱离当前 Excel 证据单独定性。页面只读展示已加载的正式 Skill，不再维护独立的补充 Skill 文件。专业覆盖、编码深度和数量瀑布对比通过 `--compare-file` 执行；DM8/LOCATIONS 目前输出导入治理约束，但未连接真实数据库做导入验证。

## 方式一：启动网页服务（推荐）

在本目录执行：

```powershell
python .\server.py --host 0.0.0.0 --port 8080 --runs-dir .\runs
```

浏览器打开 `http://服务器地址:8080/`，上传 Excel 即可。只在本机使用时，把 `--host` 改成 `127.0.0.1`。

别人使用时不需要 WorkBuddy，只需要能访问这台服务器的地址。服务提供：

- `GET /`：浏览器上传页面；
- `GET /healthz`：健康检查；
- `POST /api/audits`：接收上传并创建后台审核任务，支持 multipart 字段 `file`，返回 `run_id`；
- `GET /api/audits/{run_id}/status`：查询审核阶段、真实进度和 AI 全部候选复核状态；
- `GET /api/audits/{run_id}/{文件名}`：按上传 Excel 文件名生成的 HTML 报告或 Excel 问题清单；

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

AI 地址、模型和密钥由管理员预先写入 `config/app_config.json`，打包时随 EXE 一起发布。网页左侧的“AI 配置”是管理员配置页，可修改接口地址、模型、启用状态、批次参数和 API 密钥，并可直接测试连接。密钥输入框不会回显原文，只显示掩码；留空保存表示保持原密钥。

普通使用者只需要上传 Excel、查看报告和日志。当前服务没有内置登录认证；如果要在局域网多人使用，应在反向代理或内网访问控制层保护“AI 配置”页面和 `/api/config`、`/api/ai/test` 接口。API 密钥不要写入 Skill、Excel、报告或 Git。

## 配置、Skill 和日志

启动网页服务后，使用者可以：

- 上传 Excel 并下载审核报告；
- 查看规则审核和 AI 每批复核的实时进度；
- 查看当前已加载的正式 Skill、审核参考文档和实际执行引擎；管理员可在页面编辑或上传允许的 Skill 文件；
- 查看最近运行日志。

AI 配置可通过管理员页面维护，也可直接编辑 `config/app_config.json`；日志保存在 `logs/kks-audit.log`，审核运行产物保存在 `runs/`。

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

- `{原文件名}_审核报告.html`：给编码人员看的报告；
- `{原文件名}_问题清单.xlsx`：Excel 问题清单，并包含“规则覆盖”页，列出正式 Skill 的执行状态、命中数和证据来源。

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
