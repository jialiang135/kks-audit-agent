---
name: kks-audit-agent
description: "对编码人员上传的电厂 KKS Excel 做只读审核，输出 P0/P1/P2 报告、问题清单和需人工确认项。"
---

# KKS 编码审核智能体

当用户提交 KKS 编码 Excel 并要求审核、检查、查重复、查父级或判断能否导入时：

1. 先只读识别 Sheet、主表、表头、数据行和关键列。
2. 使用 `run_audit.py` 生成 `audit.html`、`audit.json`、`issues.csv`。
3. 管理员可在网页“AI 配置”页或 `config/app_config.json` 配置接口地址、密钥和模型；`run_audit.py` 会通过 OpenAI-compatible 接口一次性复核全部语义候选项。普通使用者不应修改 AI 配置。
4. 若需要 Excel 问题清单，使用 `build_issue_workbook.mjs` 生成 `audit.xlsx`。
5. 逐条解释 P0/P1/P2；把规则命中但语义排除的内容放到“已澄清项”。
6. 不修改源 Excel，不自动改码，不自动合并设备，不直接发布或导入数据库。

如需让多人自行使用，启动 `server.py` 提供浏览器上传页和 `/api/audits` HTTP 接口；WorkBuddy 不是运行依赖。

具体规则和误报规避经验以 `source-skill/kks-audit/` 为参考，用户需求和项目安全边界优先。
