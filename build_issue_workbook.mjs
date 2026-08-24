import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const jsonPath = process.argv[2];
const outputPath = process.argv[3];
if (!jsonPath || !outputPath) {
  throw new Error("用法: node build_issue_workbook.mjs audit.json audit.xlsx");
}

const result = JSON.parse(await fs.readFile(jsonPath, "utf8"));
const workbook = Workbook.create();
const summary = workbook.worksheets.add("概览");
const issues = workbook.worksheets.add("问题清单");
const resolved = workbook.worksheets.add("已澄清项");
const sheets = workbook.worksheets.add("文件结构");

const teal = "#0F4C5C";
const lightTeal = "#EAF4F5";
const gray = "#F4F6F7";

function styleHeader(range) {
  range.format = {
    fill: teal,
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
    verticalAlignment: "center",
  };
  range.format.borders = { preset: "all", style: "thin", color: "#B8C8CC" };
}

function styleGrid(range) {
  range.format.borders = { preset: "insideHorizontal", style: "thin", color: "#D9E2E5" };
  range.format.wrapText = true;
  range.format.verticalAlignment = "top";
}

summary.showGridLines = false;
summary.mergeCells("A1:F1");
summary.getRange("A1").values = [["KKS 编码审核报告"]];
summary.getRange("A1:F1").format = { fill: teal, font: { bold: true, color: "#FFFFFF", size: 16 }, horizontalAlignment: "center", verticalAlignment: "center" };
summary.getRange("A2:F2").merge();
summary.getRange("A2").values = [[`源文件：${result.source_file}；主表：${result.sheet}；表头行：${result.header_row}`]];
summary.getRange("A2:F2").format = { fill: lightTeal, font: { color: "#36525B" }, wrapText: true };
summary.getRange("A4:B4").values = [["指标", "值"]];
styleHeader(summary.getRange("A4:B4"));
summary.getRange("A5:B12").values = [
  ["有效编码行", result.data_rows],
  ["唯一 KKS 码", result.metrics.unique_codes],
  ["问题总数", result.issue_count ?? result.issues.filter((x) => x.status !== "resolved").length],
  ["已澄清项", result.resolved_count ?? result.resolved_reviews.length],
  ["父级孤儿", result.metrics.orphan_rows],
  ["父子前缀不一致", result.metrics.prefix_mismatch_rows],
  ["12 位以上扩展码", result.metrics.long_code_rows],
];
styleGrid(summary.getRange("A5:B12"));
summary.getRange("A14:F14").merge();
summary.getRange("A14").values = [["总体结论"]];
summary.getRange("A14:F14").format = { fill: teal, font: { bold: true, color: "#FFFFFF" } };
summary.getRange("A15:F16").merge();
summary.getRange("A15").values = [[result.conclusion]];
summary.getRange("A15:F16").format = { fill: lightTeal, wrapText: true, verticalAlignment: "center" };
summary.getRange("A18:F18").merge();
summary.getRange("A18").values = [["结构与使用说明"]];
summary.getRange("A18:F18").format = { fill: teal, font: { bold: true, color: "#FFFFFF" } };
const notes = result.structural_notes.map((x) => `• ${x}`);
notes.push("• 源 Excel 未修改；13 位扩展码和历史映射均需按证据人工确认。", "• 本表未做 DM8/LOCATIONS 真实导入验证。");
summary.getRange(`A19:F${18 + notes.length}`).merge(true);
summary.getRange("A19").values = notes.map((x) => [x]);
summary.getRange(`A19:F${18 + notes.length}`).format = { fill: gray, wrapText: true, verticalAlignment: "top" };
summary.getRange("A:A").format.columnWidth = 24;
summary.getRange("B:B").format.columnWidth = 18;
summary.getRange("C:F").format.columnWidth = 18;
summary.getRange("A1:F1").format.rowHeight = 28;
summary.freezePanes.freezeRows(4);

const issueHeaders = ["状态", "规则", "类别", "Excel行号", "KKS码", "父级码", "原KKS码", "设备名称", "审核意见", "建议", "证据"];
issues.showGridLines = false;
issues.getRange(`A1:${String.fromCharCode(64 + issueHeaders.length)}1`).values = [issueHeaders];
styleHeader(issues.getRange(`A1:${String.fromCharCode(64 + issueHeaders.length)}1`));
const issueRows = result.issues.map((x) => issueHeaders.map((h) => ({
  "状态": x.status, "规则": x.rule_id, "类别": x.category, "Excel行号": x.excel_row,
  "KKS码": x.kks_code, "父级码": x.parent_code, "原KKS码": x.old_code, "设备名称": x.name,
  "审核意见": x.message, "建议": x.suggestion, "证据": x.evidence,
}[h] ?? "")));
if (issueRows.length) issues.getRange(`A2:K${issueRows.length + 1}`).values = issueRows;
styleGrid(issues.getRange(`A1:K${Math.max(1, issueRows.length + 1)}`));
issues.getRange(`A1:K${Math.max(1, issueRows.length + 1)}`).format.rowHeight = 32;
issues.getRange("A:A").format.columnWidth = 10;
issues.getRange("B:B").format.columnWidth = 14;
issues.getRange("C:C").format.columnWidth = 12;
issues.getRange("D:D").format.columnWidth = 24;
issues.getRange("E:E").format.columnWidth = 10;
issues.getRange("F:H").format.columnWidth = 18;
issues.getRange("I:I").format.columnWidth = 32;
issues.getRange("I:K").format.columnWidth = 48;
issues.tables.add(`A1:K${Math.max(1, issueRows.length + 1)}`, true, "KKSIssues");
issues.freezePanes.freezeRows(1);

const resolvedHeaders = ["状态", "规则", "类别", "Excel行号", "KKS码", "父级码", "原KKS码", "设备名称", "复核结论", "建议", "证据"];
resolved.showGridLines = false;
resolved.getRange(`A1:K1`).values = [resolvedHeaders];
styleHeader(resolved.getRange("A1:K1"));
const resolvedRows = result.resolved_reviews.map((x) => [x.status, x.rule_id, x.category, x.excel_row, x.kks_code, x.parent_code, x.old_code, x.name, x.message, x.suggestion, x.evidence]);
if (resolvedRows.length) resolved.getRange(`A2:K${resolvedRows.length + 1}`).values = resolvedRows;
styleGrid(resolved.getRange(`A1:K${Math.max(1, resolvedRows.length + 1)}`));
resolved.getRange("A:K").format.columnWidth = 22;
resolved.getRange("H:K").format.columnWidth = 48;
resolved.tables.add(`A1:K${Math.max(1, resolvedRows.length + 1)}`, true, "KKSResolved");
resolved.freezePanes.freezeRows(1);

const sheetHeaders = ["Sheet", "最大行", "最大列", "非空行", "首个非空行", "最后非空行", "表头预览"];
sheets.showGridLines = false;
sheets.getRange("A1:G1").values = [sheetHeaders];
styleHeader(sheets.getRange("A1:G1"));
const sheetRows = result.sheet_info.map((x) => [x.sheet, x.max_row, x.max_col, x.nonempty_rows, x.first_nonempty_row ?? "", x.last_nonempty_row ?? "", x.header_preview.join(" | ")]);
sheets.getRange(`A2:G${sheetRows.length + 1}`).values = sheetRows;
styleGrid(sheets.getRange(`A1:G${sheetRows.length + 1}`));
sheets.getRange("A:A").format.columnWidth = 18;
sheets.getRange("B:F").format.columnWidth = 14;
sheets.getRange("G:G").format.columnWidth = 60;
sheets.tables.add(`A1:G${sheetRows.length + 1}`, true, "KKSSheets");
sheets.freezePanes.freezeRows(1);

await fs.mkdir(new URL(`file://${outputPath.replaceAll("\\", "/")}`).pathname.replace(/\/$/, "").replace(/\/[^/]+$/, ""), { recursive: true }).catch(() => {});
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);

const check = await workbook.inspect({ kind: "table", sheetId: "概览", range: "A1:F20", tableMaxRows: 20, tableMaxCols: 6, maxChars: 5000 });
console.log(check.ndjson);
