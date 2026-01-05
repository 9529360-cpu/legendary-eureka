/**
 * Excel 命令执行器
 * v2.9.8: 从 App.tsx 提取的 Excel 命令执行逻辑
 *
 * 这个模块处理旧的 CopilotAction/ExcelCommand 模式
 * 长期计划是统一到 Agent 模式
 */

import type { CellValue, CopilotAction, WorkbookContext } from "../types";
import type { ChatResponse } from "../../services/ApiService";
import { getCommandRangeAddress, coerceCellValue, extractHeaders } from "./excel.utils";

// ========== 类型定义 ==========

type ExcelCommand = NonNullable<ChatResponse["excelCommand"]>;

interface CopilotResponse {
  message: string;
  actions: CopilotAction[];
}

interface ValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
  suggestions: string[];
  autoFixApplied: boolean;
  fixedParameters?: Record<string, unknown>;
}

// ========== 命令规范化 ==========

export function normalizeExcelCommandAction(command: ExcelCommand): string {
  const type = String(command.type || "").toLowerCase();
  const action = String(command.action || "").toLowerCase();

  if (type === "write" && action === "range") return "write_range";
  if (type === "write" && action === "cell") return "write_cell";
  if (type === "formula" && action === "set") return "set_formula";
  if (type === "table" && action === "create") return "create_table";
  if (type === "chart" && action === "create") return "create_chart";

  return action || type;
}

// ========== 数据构建辅助函数 ==========

function hasMatchingHeaderRow(values: CellValue[][], headers: string[]): boolean {
  if (values.length === 0) {
    return false;
  }
  if (values[0].length !== headers.length) {
    return false;
  }

  return headers.every((header, index) => String(values[0][index]) === String(header));
}

function getRequestedSampleCount(parameters: Record<string, unknown>): number | null {
  const p = parameters as Record<string, string | number | undefined>;
  const raw =
    p.sampleCount ?? p.sampleRows ?? p.sampleSize ?? p.rowCount ?? p.count ?? p.size ?? null;

  if (raw === null || raw === undefined) {
    return null;
  }

  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }

  return Math.min(Math.floor(parsed), 200);
}

function inferSampleValue(header: string, index: number): CellValue {
  const label = String(header || "");
  const lower = label.toLowerCase();
  const rowIndex = index + 1;

  if (label.includes("日期") || label.includes("时间") || lower.includes("date")) {
    const base = new Date();
    base.setDate(base.getDate() + index);
    const year = base.getFullYear();
    const month = String(base.getMonth() + 1).padStart(2, "0");
    const day = String(base.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  if (label.includes("数量") || lower.includes("qty")) {
    return rowIndex * 3;
  }
  if (
    label.includes("金额") ||
    label.includes("价格") ||
    label.includes("销售") ||
    label.includes("总计") ||
    label.includes("成本") ||
    label.includes("费用") ||
    lower.includes("price") ||
    lower.includes("amount")
  ) {
    return rowIndex * 25;
  }
  if (label.includes("率") || lower.includes("rate")) {
    return Number((0.05 * rowIndex).toFixed(2));
  }
  if (
    label.includes("人员") ||
    label.includes("员工") ||
    label.includes("姓名") ||
    label.includes("客户")
  ) {
    return `人员${rowIndex}`;
  }
  if (label.includes("地区") || label.includes("区域")) {
    return `区域${(rowIndex % 3) + 1}`;
  }
  if (label.includes("商品") || label.includes("产品") || label.includes("物品")) {
    return `商品${rowIndex}`;
  }
  if (label.includes("支付") || label.includes("方式") || label.includes("渠道")) {
    return `方式${(rowIndex % 3) + 1}`;
  }
  if (label.includes("部门") || label.includes("类别") || label.includes("分类")) {
    return `分类${(rowIndex % 3) + 1}`;
  }
  if (label.includes("编号") || lower.includes("id")) {
    return `NO-${String(rowIndex).padStart(3, "0")}`;
  }

  return `数据${rowIndex}`;
}

function buildSampleRows(headers: string[], count: number, offset: number = 0): CellValue[][] {
  const rows: CellValue[][] = [];
  for (let i = 0; i < count; i++) {
    const row: CellValue[] = [];
    for (let j = 0; j < headers.length; j++) {
      row.push(inferSampleValue(headers[j], offset + i));
    }
    rows.push(row);
  }
  return rows;
}

function mergeSampleRows(
  headers: string[],
  existing: CellValue[][],
  totalCount: number
): CellValue[][] {
  if (existing.length >= totalCount) {
    return existing;
  }

  const missing = totalCount - existing.length;
  return existing.concat(buildSampleRows(headers, missing, existing.length));
}

function buildValuesFromRaw(
  raw: unknown,
  headers: string[] | null,
  parameters: Record<string, unknown>
): CellValue[][] | null {
  if (raw === null || raw === undefined) {
    return null;
  }

  if (Array.isArray(raw)) {
    if (raw.length === 0) {
      return [];
    }

    if (Array.isArray(raw[0])) {
      return (raw as unknown[][]).map((row) => row.map(coerceCellValue));
    }

    if (typeof raw[0] === "object" && raw[0] !== null) {
      const rows = raw as Record<string, unknown>[];
      const resolvedHeaders = headers ?? Object.keys(rows[0] || {});
      if (!resolvedHeaders.length) {
        return null;
      }
      return [
        resolvedHeaders,
        ...rows.map((row) => resolvedHeaders.map((key) => coerceCellValue(row[key]))),
      ];
    }

    return [raw.map(coerceCellValue)];
  }

  if (typeof raw === "object") {
    const rawObject = raw as Record<string, unknown>;
    const nested =
      rawObject.values ??
      rawObject.data ??
      rawObject.rows ??
      rawObject.table ??
      rawObject.items ??
      rawObject.records ??
      rawObject.entries ??
      rawObject.list ??
      rawObject.sampleRows ??
      rawObject.sample_rows ??
      rawObject.sampleData ??
      rawObject.samples ??
      rawObject.examples ??
      rawObject.exampleData ??
      rawObject.example_data;
    const nestedHeaders = headers ?? extractHeaders(rawObject) ?? extractHeaders(parameters);
    return buildValuesFromRaw(nested, nestedHeaders, parameters);
  }

  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if (
      (trimmed.startsWith("[") && trimmed.endsWith("]")) ||
      (trimmed.startsWith("{") && trimmed.endsWith("}"))
    ) {
      try {
        const parsed = JSON.parse(trimmed);
        return buildValuesFromRaw(parsed, headers, parameters);
      } catch {
        return [[raw]];
      }
    }
    return [[raw]];
  }

  return null;
}

export function buildTabularValues(parameters: Record<string, unknown>): CellValue[][] | null {
  const headers = extractHeaders(parameters);
  const sampleCount = getRequestedSampleCount(parameters);
  const p = parameters as Record<string, unknown>;
  const raw =
    p.values ??
    p.data ??
    p.rows ??
    p.table ??
    p.items ??
    p.records ??
    p.entries ??
    p.list ??
    p.sampleRows ??
    p.sample_rows ??
    p.sampleData ??
    p.samples ??
    p.examples ??
    p.exampleData ??
    p.example_data ??
    null;

  const resolved = buildValuesFromRaw(raw, headers, parameters);
  if (!resolved) {
    if (!headers) {
      return null;
    }
    const samples = sampleCount ? buildSampleRows(headers, sampleCount) : [];
    return samples.length > 0 ? [headers, ...samples] : [headers];
  }

  if (headers && headers.length > 0) {
    const hasHeader = hasMatchingHeaderRow(resolved, headers);
    const dataRows = hasHeader ? resolved.slice(1) : resolved;
    const paddedRows = sampleCount ? mergeSampleRows(headers, dataRows, sampleCount) : dataRows;
    return [headers, ...paddedRows];
  }

  return resolved;
}

// ========== 命令标签 ==========

export function getExcelCommandLabel(command: ExcelCommand): string {
  const normalized = normalizeExcelCommandAction(command);
  let baseLabel = "执行操作";

  const labelMap: Record<string, string> = {
    // 写入操作
    writerange: "写入数据",
    write_range: "写入数据",
    insert: "插入数据",
    insert_data: "插入数据",
    writecell: "写入单元格",
    write_cell: "写入单元格",
    batch_write: "批量写入",
    batchwrite: "批量写入",
    // 公式操作
    setformula: "设置公式",
    set_formula: "设置公式",
    fill_formula: "填充公式",
    fillformula: "填充公式",
    batch_formula: "批量公式",
    batchformula: "批量公式",
    // 表格操作
    createtable: "创建表格",
    create_table: "创建表格",
    // 格式操作
    formatrange: "格式化区域",
    format_range: "格式化区域",
    conditional_format: "条件格式",
    conditionalformat: "条件格式",
    auto_fit: "自动调整",
    autofit: "自动调整",
    clear_format: "清除格式",
    clearformat: "清除格式",
    // 数据操作
    sortrange: "排序",
    sort_range: "排序",
    filterrange: "筛选",
    filter_range: "筛选",
    remove_duplicates: "删除重复项",
    removeduplicates: "删除重复项",
    find_replace: "查找替换",
    findreplace: "查找替换",
    clear_range: "清除区域",
    clearrange: "清除区域",
    // 图表操作
    createchart: "创建图表",
    create_chart: "创建图表",
    // 工作表操作
    create_sheet: "创建工作表",
    createsheet: "创建工作表",
    rename_sheet: "重命名工作表",
    renamesheet: "重命名工作表",
    copy_sheet: "复制工作表",
    copysheet: "复制工作表",
    delete_sheet: "删除工作表",
    deletesheet: "删除工作表",
    switch_sheet: "切换工作表",
    switchsheet: "切换工作表",
    // 跨表操作
    copy_to_sheet: "跨表复制",
    copytosheet: "跨表复制",
    merge_sheets: "合并工作表",
    mergesheets: "合并工作表",
    // 复合操作
    multi_step: "执行多步骤操作",
    multistep: "执行多步骤操作",
    // 数据透视表
    create_pivot_table: "创建数据透视表",
    createpivottable: "创建数据透视表",
    // 命名范围
    create_named_range: "创建命名范围",
    createnamedrange: "创建命名范围",
    delete_named_range: "删除命名范围",
    deletenamedrange: "删除命名范围",
    // 数据验证
    add_data_validation: "添加数据验证",
    adddatavalidation: "添加数据验证",
    remove_data_validation: "移除数据验证",
    removedatavalidation: "移除数据验证",
  };

  baseLabel = labelMap[normalized] || "执行操作";

  const address = getCommandRangeAddress(command.parameters || {});
  return address ? `${baseLabel} ${address}` : baseLabel;
}

// ========== 命令验证 ==========

export function validateAndFixCommand(
  command: ExcelCommand,
  context?: { workbookContext: WorkbookContext | null }
): ValidationResult {
  const result: ValidationResult = {
    valid: true,
    errors: [],
    warnings: [],
    suggestions: [],
    autoFixApplied: false,
  };

  const action = normalizeExcelCommandAction(command);
  const params = command.parameters || {};

  // 1. 通用验证：地址格式检查
  const addressFields = ["address", "sourceAddress", "targetAddress", "dataRange", "rangeAddress"];
  for (const field of addressFields) {
    if (params[field]) {
      const addr = String(params[field]);
      // 检查地址格式是否合法
      if (!/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(addr) && !/^'[^']+'![A-Za-z]+\d+/.test(addr)) {
        // 尝试自动修复：去除多余空格
        const fixed = addr.trim().replace(/\s+/g, "");
        if (/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(fixed)) {
          params[field] = fixed;
          result.autoFixApplied = true;
          result.warnings.push(`地址 "${addr}" 已自动修复为 "${fixed}"`);
        } else {
          result.errors.push(`无效的单元格地址格式: ${addr}`);
          result.valid = false;
        }
      }
    }
  }

  // 2. 操作特定验证
  switch (action) {
    case "write_range":
    case "writerange":
    case "insert_data":
    case "create_table":
    case "createtable": {
      // 验证数据不为空
      const data = params.data || params.values || params.rows;
      if (!data || (Array.isArray(data) && data.length === 0)) {
        result.errors.push("写入数据不能为空");
        result.valid = false;
      }
      // 验证headers参数
      if (!params.headers && action.includes("table")) {
        result.warnings.push("未指定表头，将使用默认表头");
      }
      break;
    }

    case "set_formula":
    case "setformula":
    case "fill_formula": {
      const formula = params.formula as string | undefined;
      if (!formula) {
        result.errors.push("公式不能为空");
        result.valid = false;
      } else if (typeof formula === "string") {
        // 自动添加等号
        if (!formula.startsWith("=")) {
          params.formula = "=" + formula;
          result.autoFixApplied = true;
          result.warnings.push("已自动添加公式前缀 '='");
        }
        // 检查括号匹配
        const openParens = (formula.match(/\(/g) || []).length;
        const closeParens = (formula.match(/\)/g) || []).length;
        if (openParens !== closeParens) {
          result.warnings.push(`公式括号不匹配: ${openParens}个'(' vs ${closeParens}个')'`);
        }
      }
      break;
    }

    case "create_chart":
    case "createchart": {
      const chartType = params.chartType || params.type;
      const validTypes = ["column", "bar", "line", "pie", "scatter", "area", "doughnut", "radar"];
      if (chartType && !validTypes.includes(String(chartType).toLowerCase())) {
        result.warnings.push(`未知图表类型 "${chartType}"，将使用柱状图`);
        params.chartType = "column";
        result.autoFixApplied = true;
      }
      if (!params.dataRange && !params.address) {
        result.suggestions.push("未指定数据范围，将使用当前选区");
      }
      break;
    }

    case "switch_sheet":
    case "rename_sheet":
    case "delete_sheet": {
      const sheetName = params.sheetName || params.name || params.oldName;
      if (sheetName && context?.workbookContext) {
        const exists = context.workbookContext.sheets.some((s) => s.name === sheetName);
        if (!exists) {
          result.errors.push(`工作表 "${sheetName}" 不存在`);
          result.suggestions.push(
            `可用工作表: ${context.workbookContext.sheets.map((s) => s.name).join(", ")}`
          );
          result.valid = false;
        }
      }
      break;
    }

    case "copy_to_sheet": {
      const targetSheet = params.targetSheet;
      if (targetSheet && context?.workbookContext) {
        const exists = context.workbookContext.sheets.some((s) => s.name === targetSheet);
        if (!exists) {
          result.warnings.push(`目标工作表 "${targetSheet}" 不存在，将自动创建`);
        }
      }
      break;
    }

    case "conditional_format": {
      const rule = params.rule;
      const validRules = [
        "greaterThan",
        "lessThan",
        "equalTo",
        "between",
        "containsText",
        "cellIsBlank",
      ];
      if (rule && !validRules.includes(rule as string)) {
        result.warnings.push(`未知条件格式规则 "${rule}"，将使用 greaterThan`);
        params.rule = "greaterThan";
        result.autoFixApplied = true;
      }
      break;
    }

    case "sort_range": {
      const column = params.column ?? params.sortColumn ?? params.columnIndex;
      if (column !== undefined && typeof column !== "number") {
        // 尝试转换为数字
        const num = parseInt(String(column), 10);
        if (!isNaN(num)) {
          params.column = num;
          result.autoFixApplied = true;
        } else {
          result.errors.push(`排序列索引必须是数字: ${column}`);
          result.valid = false;
        }
      }
      break;
    }

    case "add_data_validation": {
      const validationType = params.type;
      const validTypes = ["list", "number", "date", "textLength", "wholeNumber", "decimal"];
      if (validationType && !validTypes.includes(validationType as string)) {
        result.warnings.push(`未知验证类型 "${validationType}"，将使用列表验证`);
        params.type = "list";
        result.autoFixApplied = true;
      }
      break;
    }
  }

  // 3. 返回修复后的参数
  if (result.autoFixApplied) {
    result.fixedParameters = params;
  }

  return result;
}

// ========== 操作目标地址获取 ==========

export function getActionTargetAddress(action: CopilotAction): string | null {
  if (action.type === "writeRange" || action.type === "setFormula" || action.type === "writeCell") {
    return action.address || null;
  }

  if (action.type === "executeCommand" && action.command) {
    const params = action.command.parameters || {};
    return (
      params.address || params.rangeAddress || params.targetAddress || params.sourceAddress || null
    );
  }

  return null;
}

// ========== AI 响应转换 ==========

export function convertAiResponseToCopilotResponse(aiResponse: ChatResponse): CopilotResponse {
  if (!aiResponse.success) {
    return {
      message: aiResponse.error || aiResponse.fallback || "AI服务返回了错误响应",
      actions: [],
    };
  }

  const actions: CopilotAction[] = [];

  if (aiResponse.excelCommand && aiResponse.excelCommand.executable) {
    const normalized = normalizeExcelCommandAction(aiResponse.excelCommand);
    if (normalized === "unknown") {
      return {
        message: aiResponse.message,
        actions: [],
      };
    }

    const { type, action, parameters } = aiResponse.excelCommand;

    switch (type) {
      case "write":
        if (action === "range" && parameters.address && parameters.values) {
          actions.push({
            type: "writeRange",
            address: parameters.address,
            values: parameters.values,
          });
        } else if (action === "cell" && parameters.address && parameters.value !== undefined) {
          actions.push({
            type: "writeCell",
            address: parameters.address,
            value: parameters.value,
          });
        }
        break;

      case "formula":
        if (action === "set" && parameters.address && parameters.formula) {
          actions.push({
            type: "setFormula",
            address: parameters.address,
            formula: parameters.formula,
          });
        }
        break;
    }

    if (actions.length === 0) {
      actions.push({
        type: "executeCommand",
        command: aiResponse.excelCommand,
        label: getExcelCommandLabel(aiResponse.excelCommand),
      });
    }
  }

  return {
    message: aiResponse.message,
    actions,
  };
}
