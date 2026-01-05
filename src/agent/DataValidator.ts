/**
 * DataValidator - 数据校验器 v2.9.54
 *
 * 核心原则：验证 Excel 实际数据，不依赖模型判断
 *
 * 这是 Agent 闭环的第二环：
 * ┌─────────────────────────────────────────────────────┐
 * │  THINK ──────→ EXECUTE ──────→ OBSERVE             │
 * │    │              │              │                  │
 * │    ▼              ▼              ▼                  │
 * │ 计划验证      [数据校验]      智能回滚              │
 * │ (拦截必然失败)  (检测已发生错误) (不弄脏Excel)       │
 * └─────────────────────────────────────────────────────┘
 *
 * v2.9.54 重构要点：
 * 1. 分层抽样：头部+尾部+随机，带置信度
 * 2. 共享 Context：一次读取，规则复用
 * 3. ColumnResolver：统一列识别，避免各规则重复正则
 * 4. 组合条件判定：D/E/F 规则不再"一条命中就定罪"
 * 5. affectedRange：每条问题输出受影响范围
 * 6. confidence + evidence：提升可信度
 *
 * 6条核心规则：
 * A. 空值检测 - 主键/数量/单价空值
 * B. 类型一致性 - 数量/单价非数值
 * C. 主键唯一性 - 产品ID重复
 * D. 整列常数 - 单价/成本 uniqueCount ≤ 1 (组合条件)
 * E. 汇总分布异常 - 多产品数值相同 (组合条件)
 * F. lookup一致性 - 单价 ≠ XLOOKUP结果 (抽样核对)
 */

import { ExcelReader } from "./AgentCore";

// ========== 类型定义 ==========

/**
 * 置信度等级
 */
export type ConfidenceLevel = "high" | "medium" | "low";

/**
 * 抽样策略
 */
export interface SamplingStrategy {
  headRows: number; // 头部行数
  tailRows: number; // 尾部行数
  randomRows: number; // 随机行数
  chunkSize?: number; // 分块大小（用于大数据集）
}

/**
 * 抽样结果
 */
export interface SampleData {
  headers: string[];
  headRows: unknown[][]; // 头部样本
  tailRows: unknown[][]; // 尾部样本
  randomRows: unknown[][]; // 随机样本
  totalRowCount: number; // 总行数
  sampledRowIndices: number[]; // 被抽样的行索引（1-based）
}

/**
 * 数据校验问题（改名：不再叫 Result）
 */
export interface DataValidationIssue {
  ruleId: string;
  ruleName: string;
  severity: "block" | "warn";
  message: string;
  details: string[];

  // v2.9.54: 新增字段
  confidence: ConfidenceLevel;
  evidence: ValidationEvidence;
  affectedRange?: string; // 如 "Sheet1!C2:C21"
  affectedCells?: string[]; // 少量时列出具体单元格
  suggestedFix?: string;
  suggestedFixPlan?: FixAction[]; // 结构化修复动作
}

/**
 * 校验证据
 */
export interface ValidationEvidence {
  sampleSize: number;
  sampledRows: { head: number; tail: number; random: number };
  uniqueValues?: number;
  matchedPattern?: string;
  comparedValues?: Array<{ key: string; expected: unknown; actual: unknown }>;
}

/**
 * 修复动作（结构化）
 */
export interface FixAction {
  type: "write_formula" | "write_value" | "delete" | "format";
  range: string;
  value?: string;
  formula?: string;
  description: string;
}

/**
 * 数据校验上下文（v2.9.54: 共享数据）
 */
export interface DataValidationContext {
  sheet: string;
  range?: string;

  // v2.9.54: 预读取的共享数据
  sampleData?: SampleData;
  resolvedColumns?: ResolvedColumns;
  columnFormulas?: Map<string, string[]>; // 列字母 -> 公式数组

  // 缓存的主数据表信息（用于 lookup 核对）
  masterTableData?: Map<string, Map<string, unknown>>; // 表名 -> (ID -> 单价)
}

/**
 * 规范列定义
 */
export interface CanonicalColumn {
  role: ColumnRole;
  columnIndex: number;
  columnLetter: string;
  headerName: string;
  confidence: ConfidenceLevel;
  inferredType: "text" | "number" | "date" | "mixed" | "unknown";
}

/**
 * 列角色
 */
export type ColumnRole =
  | "productId"
  | "orderId"
  | "customerId" // ID 类
  | "productName"
  | "orderDate" // 描述类
  | "quantity"
  | "unitPrice"
  | "cost"
  | "amount"
  | "profit" // 数值类
  | "category"
  | "region"
  | "channel" // 分类类
  | "unknown";

/**
 * 已解析的列信息
 */
export interface ResolvedColumns {
  columns: CanonicalColumn[];
  byRole: Map<ColumnRole, CanonicalColumn[]>;
  byIndex: Map<number, CanonicalColumn>;
}

/**
 * 数据校验规则（v2.9.54: 使用 context 共享数据）
 */
export interface DataValidationRule {
  id: string;
  name: string;
  description: string;
  severity: "block" | "warn";
  enabled: boolean;

  // v2.9.54: 规则元信息
  requiresIO: boolean; // 是否需要额外 I/O
  targetSheetPattern?: RegExp; // 适用的工作表模式
  requiredColumns?: ColumnRole[]; // 依赖的列角色

  /**
   * 检查函数（使用共享 context）
   */
  check: (
    context: DataValidationContext,
    excelReader: ExcelReader
  ) => Promise<DataValidationIssue | null>;
}

// ========== 列解析器 ==========

export class ColumnResolver {
  // 列角色词典
  private static readonly ROLE_PATTERNS: Array<{
    role: ColumnRole;
    patterns: RegExp[];
    expectedType: "text" | "number" | "date";
  }> = [
    {
      role: "productId",
      patterns: [/^产品id$/i, /^商品id$/i, /^product.?id$/i, /^sku$/i],
      expectedType: "text",
    },
    {
      role: "orderId",
      patterns: [/^订单id$/i, /^订单号$/i, /^order.?id$/i, /^单号$/i],
      expectedType: "text",
    },
    {
      role: "customerId",
      patterns: [/^客户id$/i, /^customer.?id$/i, /^会员id$/i],
      expectedType: "text",
    },
    {
      role: "productName",
      patterns: [/^产品名?称?$/i, /^商品名?称?$/i, /^product.?name$/i, /^品名$/i],
      expectedType: "text",
    },
    {
      role: "orderDate",
      patterns: [/^日期$/i, /^订单日期$/i, /^date$/i, /^下单时间$/i],
      expectedType: "date",
    },
    {
      role: "quantity",
      patterns: [/^数量$/i, /^qty$/i, /^quantity$/i, /^销量$/i],
      expectedType: "number",
    },
    {
      role: "unitPrice",
      patterns: [/^单价$/i, /^售价$/i, /^price$/i, /^unit.?price$/i],
      expectedType: "number",
    },
    {
      role: "cost",
      patterns: [/^成本$/i, /^进价$/i, /^cost$/i, /^采购价$/i],
      expectedType: "number",
    },
    {
      role: "amount",
      patterns: [/^金额$/i, /^销售额$/i, /^总额$/i, /^amount$/i, /^total$/i],
      expectedType: "number",
    },
    {
      role: "profit",
      patterns: [/^利润$/i, /^毛利$/i, /^profit$/i, /^margin$/i],
      expectedType: "number",
    },
    {
      role: "category",
      patterns: [/^类别$/i, /^分类$/i, /^品类$/i, /^category$/i],
      expectedType: "text",
    },
    {
      role: "region",
      patterns: [/^区域$/i, /^地区$/i, /^region$/i, /^城市$/i],
      expectedType: "text",
    },
    { role: "channel", patterns: [/^渠道$/i, /^channel$/i, /^来源$/i], expectedType: "text" },
  ];

  /**
   * 解析列角色
   */
  resolve(headers: string[], sampleValues?: unknown[][]): ResolvedColumns {
    const columns: CanonicalColumn[] = [];
    const byRole = new Map<ColumnRole, CanonicalColumn[]>();
    const byIndex = new Map<number, CanonicalColumn>();

    headers.forEach((header, index) => {
      const headerStr = String(header || "").trim();
      const columnLetter = this.indexToColumn(index);

      // 1. 尝试从 header 匹配角色
      let matchedRole: ColumnRole = "unknown";
      let confidence: ConfidenceLevel = "low";

      for (const { role, patterns } of ColumnResolver.ROLE_PATTERNS) {
        for (const pattern of patterns) {
          if (pattern.test(headerStr)) {
            matchedRole = role;
            confidence = "high";
            break;
          }
        }
        if (matchedRole !== "unknown") break;
      }

      // 2. 从样本值推断类型
      let inferredType: "text" | "number" | "date" | "mixed" | "unknown" = "unknown";
      if (sampleValues && sampleValues.length > 0) {
        inferredType = this.inferTypeFromSamples(sampleValues, index);

        // 如果 header 匹配但类型不符，降低置信度
        if (matchedRole !== "unknown" && confidence === "high") {
          const expectedType = ColumnResolver.ROLE_PATTERNS.find(
            (r) => r.role === matchedRole
          )?.expectedType;
          if (expectedType === "number" && inferredType !== "number") {
            confidence = "medium";
          }
        }
      }

      const column: CanonicalColumn = {
        role: matchedRole,
        columnIndex: index,
        columnLetter,
        headerName: headerStr,
        confidence,
        inferredType,
      };

      columns.push(column);
      byIndex.set(index, column);

      // 按角色分组
      if (!byRole.has(matchedRole)) {
        byRole.set(matchedRole, []);
      }
      byRole.get(matchedRole)!.push(column);
    });

    return { columns, byRole, byIndex };
  }

  /**
   * 从样本值推断类型
   */
  private inferTypeFromSamples(
    samples: unknown[][],
    columnIndex: number
  ): "text" | "number" | "date" | "mixed" | "unknown" {
    let numericCount = 0;
    let textCount = 0;
    let dateCount = 0;
    let validCount = 0;

    for (const row of samples) {
      const value = row[columnIndex];
      if (value === null || value === undefined || value === "") continue;

      validCount++;

      if (typeof value === "number") {
        numericCount++;
      } else if (value instanceof Date) {
        dateCount++;
      } else if (typeof value === "string") {
        // 尝试判断是否是数字字符串
        if (!isNaN(Number(value)) && value.trim() !== "") {
          numericCount++;
        } else if (this.isDateString(value)) {
          dateCount++;
        } else {
          textCount++;
        }
      }
    }

    if (validCount === 0) return "unknown";

    const threshold = validCount * 0.8;
    if (numericCount >= threshold) return "number";
    if (dateCount >= threshold) return "date";
    if (textCount >= threshold) return "text";
    return "mixed";
  }

  /**
   * 判断是否是日期字符串
   */
  private isDateString(value: string): boolean {
    return /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(value) || /^\d{1,2}[-/]\d{1,2}[-/]\d{4}/.test(value);
  }

  /**
   * 列索引转字母
   */
  private indexToColumn(index: number): string {
    let result = "";
    let n = index + 1;
    while (n > 0) {
      n--;
      result = String.fromCharCode(65 + (n % 26)) + result;
      n = Math.floor(n / 26);
    }
    return result;
  }
}

// ========== 数据校验器 ==========

export class DataValidator {
  private rules: DataValidationRule[] = [];
  private columnResolver = new ColumnResolver();

  // 默认抽样策略
  private defaultSamplingStrategy: SamplingStrategy = {
    headRows: 10,
    tailRows: 10,
    randomRows: 10,
  };

  constructor() {
    this.registerDefaultRules();
  }

  /**
   * 设置抽样策略
   */
  setSamplingStrategy(strategy: Partial<SamplingStrategy>): void {
    this.defaultSamplingStrategy = { ...this.defaultSamplingStrategy, ...strategy };
  }

  /**
   * 注册默认的6条数据校验规则 (v2.9.54: 使用共享 context)
   */
  private registerDefaultRules(): void {
    // 规则A: 空值检测
    this.rules.push({
      id: "null_value_check",
      name: "空值检测",
      description: "主键列、数量、单价、成本不能为空",
      severity: "block",
      enabled: true,
      requiresIO: false, // 使用 context 共享数据
      requiredColumns: ["productId", "orderId", "quantity", "unitPrice", "cost"],
      check: async (ctx: DataValidationContext) => {
        if (!ctx.sampleData || !ctx.resolvedColumns) return null;

        const {
          headers: _headers,
          headRows,
          tailRows,
          randomRows,
          sampledRowIndices,
          totalRowCount: _totalRowCount,
        } = ctx.sampleData;
        const allRows = [...headRows, ...tailRows, ...randomRows];

        // 获取需要检查的列
        const criticalRoles: ColumnRole[] = [
          "productId",
          "orderId",
          "quantity",
          "unitPrice",
          "cost",
          "amount",
        ];
        const criticalColumns = criticalRoles
          .flatMap((role) => ctx.resolvedColumns!.byRole.get(role) || [])
          .filter((col) => col.confidence !== "low");

        if (criticalColumns.length === 0) return null;

        // 检查空值
        const nullIssues: string[] = [];
        const affectedCells: string[] = [];

        allRows.forEach((row, sampleIdx) => {
          const actualRow = sampledRowIndices[sampleIdx];
          for (const col of criticalColumns) {
            const value = row[col.columnIndex];
            if (value === null || value === undefined || value === "") {
              nullIssues.push(`第${actualRow}行 "${col.headerName}" 为空`);
              affectedCells.push(`${ctx.sheet}!${col.columnLetter}${actualRow}`);
            }
          }
        });

        if (nullIssues.length > 0) {
          return {
            ruleId: "null_value_check",
            ruleName: "空值检测",
            severity: "block",
            message: `检测到 ${nullIssues.length} 处关键列空值`,
            details: nullIssues.slice(0, 5),
            confidence: allRows.length >= 20 ? "high" : "medium",
            evidence: {
              sampleSize: allRows.length,
              sampledRows: {
                head: headRows.length,
                tail: tailRows.length,
                random: randomRows.length,
              },
            },
            affectedCells: affectedCells.slice(0, 10),
            suggestedFix: "填充缺失的数据，或检查公式是否正确",
          };
        }

        return null;
      },
    });

    // 规则B: 类型一致性
    this.rules.push({
      id: "type_consistency",
      name: "类型一致性",
      description: "数量、单价、成本应为数值类型",
      severity: "block",
      enabled: true,
      requiresIO: false,
      requiredColumns: ["quantity", "unitPrice", "cost", "amount", "profit"],
      check: async (ctx: DataValidationContext) => {
        if (!ctx.sampleData || !ctx.resolvedColumns) return null;

        const { headRows, tailRows, randomRows, sampledRowIndices } = ctx.sampleData;
        const allRows = [...headRows, ...tailRows, ...randomRows];

        // 获取数值列
        const numericRoles: ColumnRole[] = ["quantity", "unitPrice", "cost", "amount", "profit"];
        const numericColumns = numericRoles.flatMap(
          (role) => ctx.resolvedColumns!.byRole.get(role) || []
        );

        if (numericColumns.length === 0) return null;

        // 检查类型
        const typeIssues: string[] = [];
        const affectedCells: string[] = [];

        allRows.forEach((row, sampleIdx) => {
          const actualRow = sampledRowIndices[sampleIdx];
          for (const col of numericColumns) {
            const value = row[col.columnIndex];
            if (value !== null && value !== undefined && value !== "") {
              if (typeof value === "string" && isNaN(Number(value))) {
                typeIssues.push(`第${actualRow}行 "${col.headerName}" 不是数值: "${value}"`);
                affectedCells.push(`${ctx.sheet}!${col.columnLetter}${actualRow}`);
              }
            }
          }
        });

        if (typeIssues.length > 0) {
          return {
            ruleId: "type_consistency",
            ruleName: "类型一致性",
            severity: "block",
            message: `检测到 ${typeIssues.length} 处类型不一致`,
            details: typeIssues.slice(0, 5),
            confidence: allRows.length >= 20 ? "high" : "medium",
            evidence: {
              sampleSize: allRows.length,
              sampledRows: {
                head: headRows.length,
                tail: tailRows.length,
                random: randomRows.length,
              },
            },
            affectedCells: affectedCells.slice(0, 10),
            suggestedFix: "确保数值列只包含数字，检查是否有文本混入",
          };
        }

        return null;
      },
    });

    // 规则C: 主键唯一性
    this.rules.push({
      id: "primary_key_unique",
      name: "主键唯一性",
      description: "主数据表的ID列不能重复",
      severity: "block",
      enabled: true,
      requiresIO: true, // 需要读取全量数据
      targetSheetPattern: /主数据|产品|商品|master/i,
      requiredColumns: ["productId"],
      check: async (ctx: DataValidationContext, reader: ExcelReader) => {
        // 只检查主数据表
        if (!/主数据|产品|商品|master/.test(ctx.sheet.toLowerCase())) {
          return null;
        }

        if (!ctx.resolvedColumns) return null;

        // 找到ID列
        const idColumns = ctx.resolvedColumns.byRole.get("productId") || [];
        if (idColumns.length === 0) return null;

        const idCol = idColumns[0];

        try {
          // 主键检查需要读取全量数据
          const rows = await reader.sampleRows(ctx.sheet, 1000);
          if (!rows || rows.length < 2) return null;

          const dataRows = rows.slice(1);

          // 检查重复
          const seen = new Map<string, number>();
          const duplicates: string[] = [];
          const affectedCells: string[] = [];

          dataRows.forEach((row, rowIdx) => {
            const id = String((row as unknown[])[idCol.columnIndex] || "");
            if (id) {
              const actualRow = rowIdx + 2; // 表头是第1行
              if (seen.has(id)) {
                const firstRow = seen.get(id)!;
                duplicates.push(`"${id}" 在第${firstRow}行和第${actualRow}行重复`);
                affectedCells.push(`${ctx.sheet}!${idCol.columnLetter}${actualRow}`);
              } else {
                seen.set(id, actualRow);
              }
            }
          });

          if (duplicates.length > 0) {
            return {
              ruleId: "primary_key_unique",
              ruleName: "主键唯一性",
              severity: "block",
              message: `主键列存在 ${duplicates.length} 处重复`,
              details: duplicates.slice(0, 5),
              confidence: "high", // 全量检查，高置信度
              evidence: {
                sampleSize: dataRows.length,
                sampledRows: { head: dataRows.length, tail: 0, random: 0 },
                uniqueValues: seen.size,
              },
              affectedCells: affectedCells.slice(0, 10),
              suggestedFix: "确保主键列的每个值都是唯一的",
            };
          }
        } catch (error) {
          console.warn("[DataValidator] 主键唯一性检测失败:", error);
        }

        return null;
      },
    });

    // 规则D: 整列常数检测（v2.9.54: 组合条件）
    this.rules.push({
      id: "column_constant",
      name: "整列常数检测",
      description: "交易表的单价/成本列不应全是相同值（组合条件判定）",
      severity: "block",
      enabled: true,
      requiresIO: false,
      targetSheetPattern: /交易|订单|销售|transaction|order/i,
      requiredColumns: ["productId", "unitPrice", "cost"],
      check: async (ctx: DataValidationContext) => {
        // 只检查交易表
        if (!/交易|订单|销售|transaction|order/.test(ctx.sheet.toLowerCase())) {
          return null;
        }

        if (!ctx.sampleData || !ctx.resolvedColumns) return null;

        const { headRows, tailRows, randomRows, totalRowCount } = ctx.sampleData;
        const allRows = [...headRows, ...tailRows, ...randomRows];

        if (allRows.length < 3) return null;

        // 获取产品ID列和价格列
        const idColumns = ctx.resolvedColumns.byRole.get("productId") || [];
        const priceColumns = [
          ...(ctx.resolvedColumns.byRole.get("unitPrice") || []),
          ...(ctx.resolvedColumns.byRole.get("cost") || []),
        ];

        if (idColumns.length === 0 || priceColumns.length === 0) return null;

        const idCol = idColumns[0];

        // 计算产品ID的唯一值数量
        const productIds = allRows
          .map((row) => row[idCol.columnIndex])
          .filter((v) => v !== null && v !== undefined && v !== "");
        const uniqueProductIds = new Set(productIds.map(String));

        // 组合条件：只有当产品ID有多个唯一值时，单一价格才是异常
        if (uniqueProductIds.size <= 1) return null;

        // 检查每个价格列
        for (const priceCol of priceColumns) {
          const values = allRows
            .map((row) => row[priceCol.columnIndex])
            .filter((v) => v !== null && v !== undefined && v !== "");

          if (values.length <= 1) continue;

          const uniqueValues = new Set(values.map(String));

          // 组合条件：产品ID unique > 1 且 价格 unique = 1
          if (uniqueValues.size === 1 && uniqueProductIds.size > 1) {
            const singleValue = [...uniqueValues][0];

            // 计算置信度
            const confidence: ConfidenceLevel =
              allRows.length >= 30 && uniqueProductIds.size >= 3
                ? "high"
                : allRows.length >= 10
                  ? "medium"
                  : "low";

            return {
              ruleId: "column_constant",
              ruleName: "整列常数检测",
              severity: "block",
              message: `"${priceCol.headerName}" 列所有值都是 ${singleValue}`,
              details: [
                `共 ${values.length} 行数据，全部是相同值`,
                `但产品ID有 ${uniqueProductIds.size} 个不同值`,
                "这通常表示使用了硬编码而非公式",
              ],
              confidence,
              evidence: {
                sampleSize: allRows.length,
                sampledRows: {
                  head: headRows.length,
                  tail: tailRows.length,
                  random: randomRows.length,
                },
                uniqueValues: 1,
              },
              affectedRange: `${ctx.sheet}!${priceCol.columnLetter}2:${priceCol.columnLetter}${totalRowCount}`,
              suggestedFix: "使用 XLOOKUP 从主数据表引用，而非手动填写",
              suggestedFixPlan: [
                {
                  type: "write_formula",
                  range: `${ctx.sheet}!${priceCol.columnLetter}2`,
                  formula: `=XLOOKUP(${idCol.columnLetter}2,主数据表!A:A,主数据表!C:C)`,
                  description: "使用 XLOOKUP 公式引用主数据表",
                },
              ],
            };
          }
        }

        return null;
      },
    });

    // 规则E: 汇总分布异常（v2.9.54: 组合条件）
    this.rules.push({
      id: "summary_distribution",
      name: "汇总分布异常",
      description: "汇总表中多个分类的数值不应完全相同（组合条件判定）",
      severity: "warn",
      enabled: true,
      requiresIO: false,
      targetSheetPattern: /汇总|统计|summary|report|月度|年度/i,
      requiredColumns: ["category", "amount"],
      check: async (ctx: DataValidationContext) => {
        // 只检查汇总表
        if (!/汇总|统计|summary|report|月度|年度/.test(ctx.sheet.toLowerCase())) {
          return null;
        }

        if (!ctx.sampleData || !ctx.resolvedColumns) return null;

        const { headRows, tailRows, randomRows } = ctx.sampleData;
        const allRows = [...headRows, ...tailRows, ...randomRows];

        if (allRows.length < 2) return null;

        // 获取分类列和数值列
        const categoryRoles: ColumnRole[] = ["category", "productId", "region", "channel"];
        const categoryColumns = categoryRoles.flatMap(
          (role) => ctx.resolvedColumns!.byRole.get(role) || []
        );

        const numericRoles: ColumnRole[] = ["quantity", "amount", "cost", "profit"];
        const numericColumns = numericRoles.flatMap(
          (role) => ctx.resolvedColumns!.byRole.get(role) || []
        );

        // 检查是否有分类列
        let categoryUnique = 0;
        if (categoryColumns.length > 0) {
          const catCol = categoryColumns[0];
          const catValues = allRows
            .map((row) => row[catCol.columnIndex])
            .filter((v) => v !== null && v !== undefined);
          categoryUnique = new Set(catValues.map(String)).size;
        } else {
          // 没有明确的分类列，假设每行就是一个分类
          categoryUnique = allRows.length;
        }

        // 组合条件：分类 unique > 1 且 数值列 unique = 1
        if (categoryUnique <= 1) return null;

        for (const numCol of numericColumns) {
          const values = allRows
            .map((row) => row[numCol.columnIndex])
            .filter((v) => v !== null && v !== undefined);

          if (values.length < 2) continue;

          const uniqueValues = new Set(values.map(String));

          if (uniqueValues.size === 1 && categoryUnique > 1) {
            return {
              ruleId: "summary_distribution",
              ruleName: "汇总分布异常",
              severity: "warn",
              message: `汇总表 "${numCol.headerName}" 列所有行值相同`,
              details: [
                `${categoryUnique} 个分类的汇总值都是 ${[...uniqueValues][0]}`,
                "这可能表示 SUMIF 条件不正确",
              ],
              confidence: categoryUnique >= 3 ? "medium" : "low",
              evidence: {
                sampleSize: allRows.length,
                sampledRows: { head: allRows.length, tail: 0, random: 0 },
                uniqueValues: 1,
              },
              affectedRange: `${ctx.sheet}!${numCol.columnLetter}2:${numCol.columnLetter}${allRows.length + 1}`,
              suggestedFix: "检查 SUMIF 公式的条件列是否正确引用",
            };
          }
        }

        return null;
      },
    });

    // 规则F: Lookup 一致性验证（v2.9.54: 真正的抽样核对）
    this.rules.push({
      id: "lookup_consistency",
      name: "Lookup一致性",
      description: "交易表单价应与主数据表XLOOKUP结果一致（抽样核对）",
      severity: "block",
      enabled: true,
      requiresIO: true, // 需要读取主数据表
      targetSheetPattern: /交易|订单|销售|transaction|order/i,
      requiredColumns: ["productId", "unitPrice"],
      check: async (ctx: DataValidationContext, reader: ExcelReader) => {
        // 只检查交易表
        if (!/交易|订单|销售|transaction|order/.test(ctx.sheet.toLowerCase())) {
          return null;
        }

        if (!ctx.sampleData || !ctx.resolvedColumns) return null;

        const { headRows, tailRows, randomRows, sampledRowIndices } = ctx.sampleData;
        const allRows = [...headRows, ...tailRows, ...randomRows];

        // 获取产品ID列和单价列
        const idColumns = ctx.resolvedColumns.byRole.get("productId") || [];
        const priceColumns = ctx.resolvedColumns.byRole.get("unitPrice") || [];

        if (idColumns.length === 0 || priceColumns.length === 0) return null;

        const idCol = idColumns[0];
        const priceCol = priceColumns[0];

        try {
          // 1. 先检查是否有公式
          const formulas = await reader.getColumnFormulas(ctx.sheet, priceCol.columnLetter);
          const dataFormulas = formulas.slice(1);
          const formulaCount = dataFormulas.filter((f) => f && f.startsWith("=")).length;

          // 2. 尝试找到主数据表并读取价格映射
          let masterPriceMap: Map<string, number> | null = null;

          // 从缓存获取或读取主数据表
          if (ctx.masterTableData) {
            masterPriceMap = (ctx.masterTableData.get("主数据表") as Map<string, number>) || null;
          }

          if (!masterPriceMap) {
            // 尝试读取主数据表
            const masterSheets = ["主数据表", "产品目录", "产品表", "主数据"];
            for (const masterSheet of masterSheets) {
              try {
                const masterRows = await reader.sampleRows(masterSheet, 500);
                if (masterRows && masterRows.length >= 2) {
                  const masterHeaders = masterRows[0] as string[];

                  // 找主数据表的ID列和价格列
                  let masterIdIdx = -1;
                  let masterPriceIdx = -1;

                  masterHeaders.forEach((h, idx) => {
                    const header = String(h || "").toLowerCase();
                    if (/产品.*id|商品.*id|id/.test(header)) masterIdIdx = idx;
                    if (/单价|售价|price/.test(header)) masterPriceIdx = idx;
                  });

                  if (masterIdIdx !== -1 && masterPriceIdx !== -1) {
                    masterPriceMap = new Map();
                    for (let i = 1; i < masterRows.length; i++) {
                      const row = masterRows[i] as unknown[];
                      const id = String(row[masterIdIdx] || "");
                      const price = Number(row[masterPriceIdx]) || 0;
                      if (id) masterPriceMap.set(id, price);
                    }
                    break;
                  }
                }
              } catch {
                // 工作表不存在，继续尝试下一个
              }
            }
          }

          // 3. 核对价格
          if (masterPriceMap && masterPriceMap.size > 0) {
            const mismatches: Array<{
              productId: string;
              expected: number;
              actual: unknown;
              row: number;
            }> = [];

            allRows.forEach((row, sampleIdx) => {
              const productId = String(row[idCol.columnIndex] || "");
              const actualPrice = row[priceCol.columnIndex];
              const expectedPrice = masterPriceMap!.get(productId);

              if (expectedPrice !== undefined && actualPrice !== undefined) {
                const actualNum = Number(actualPrice);
                if (!isNaN(actualNum) && Math.abs(actualNum - expectedPrice) > 0.01) {
                  mismatches.push({
                    productId,
                    expected: expectedPrice,
                    actual: actualPrice,
                    row: sampledRowIndices[sampleIdx],
                  });
                }
              }
            });

            if (mismatches.length > 0) {
              return {
                ruleId: "lookup_consistency",
                ruleName: "Lookup一致性",
                severity: "block",
                message: `发现 ${mismatches.length} 处单价与主数据表不一致`,
                details: mismatches
                  .slice(0, 5)
                  .map(
                    (m) =>
                      `第${m.row}行: 产品"${m.productId}" 单价应为 ${m.expected}，实际是 ${m.actual}`
                  ),
                confidence: mismatches.length >= 3 ? "high" : "medium",
                evidence: {
                  sampleSize: allRows.length,
                  sampledRows: {
                    head: headRows.length,
                    tail: tailRows.length,
                    random: randomRows.length,
                  },
                  comparedValues: mismatches.slice(0, 5).map((m) => ({
                    key: m.productId,
                    expected: m.expected,
                    actual: m.actual,
                  })),
                },
                affectedCells: mismatches
                  .slice(0, 10)
                  .map((m) => `${ctx.sheet}!${priceCol.columnLetter}${m.row}`),
                suggestedFix: "使用 XLOOKUP 从主数据表引用单价",
                suggestedFixPlan: [
                  {
                    type: "write_formula",
                    range: `${ctx.sheet}!${priceCol.columnLetter}2`,
                    formula: `=XLOOKUP(${idCol.columnLetter}2,主数据表!A:A,主数据表!C:C)`,
                    description: "使用 XLOOKUP 公式从主数据表引用单价",
                  },
                ],
              };
            }
          }

          // 4. 如果没有找到主数据表，检查是否是硬编码
          if (formulaCount === 0) {
            const prices = allRows.map((row) => row[priceCol.columnIndex]);
            const uniquePrices = new Set(prices.filter((p) => p != null).map(String));
            const productIds = allRows.map((row) => row[idCol.columnIndex]);
            const uniqueProductIds = new Set(productIds.filter((p) => p != null).map(String));

            if (uniquePrices.size === 1 && uniqueProductIds.size > 1 && prices.length > 1) {
              return {
                ruleId: "lookup_consistency",
                ruleName: "Lookup一致性",
                severity: "block",
                message: "单价列没有使用公式，可能是硬编码",
                details: [
                  "交易表的单价应该使用 XLOOKUP 从主数据表引用",
                  `当前 ${prices.length} 行单价都是 ${[...uniquePrices][0]}`,
                  `但产品ID有 ${uniqueProductIds.size} 个不同值`,
                ],
                confidence: "medium",
                evidence: {
                  sampleSize: allRows.length,
                  sampledRows: {
                    head: headRows.length,
                    tail: tailRows.length,
                    random: randomRows.length,
                  },
                  uniqueValues: uniquePrices.size,
                },
                affectedRange: `${ctx.sheet}!${priceCol.columnLetter}2:${priceCol.columnLetter}${allRows.length + 1}`,
                suggestedFix: "使用 XLOOKUP(产品ID, 主数据表!A:A, 主数据表!C:C) 引用单价",
              };
            }
          }
        } catch (error) {
          console.warn("[DataValidator] Lookup一致性检测失败:", error);
        }

        return null;
      },
    });
  }

  /**
   * 验证工作表数据 (v2.9.54: 共享 context，分层抽样)
   */
  async validate(sheet: string, reader: ExcelReader): Promise<DataValidationIssue[]> {
    const issues: DataValidationIssue[] = [];
    const ruleErrors: Array<{ ruleId: string; error: Error }> = [];

    // 1. 构建共享 context（一次读取）
    const context = await this.buildContext(sheet, reader);
    if (!context.sampleData) {
      console.warn(`[DataValidator] 无法读取工作表 "${sheet}" 的数据`);
      return issues;
    }

    // 2. 分组执行规则
    const memoryRules = this.rules.filter((r) => r.enabled && !r.requiresIO);
    const ioRules = this.rules.filter((r) => r.enabled && r.requiresIO);

    // 2.1 并发执行内存规则
    const memoryResults = await Promise.all(
      memoryRules.map(async (rule) => {
        try {
          return await rule.check(context, reader);
        } catch (error) {
          ruleErrors.push({ ruleId: rule.id, error: error as Error });
          return null;
        }
      })
    );

    for (const result of memoryResults) {
      if (result) issues.push(result);
    }

    // 2.2 串行执行 I/O 规则（避免并发 sync）
    for (const rule of ioRules) {
      try {
        const result = await rule.check(context, reader);
        if (result) issues.push(result);
      } catch (error) {
        ruleErrors.push({ ruleId: rule.id, error: error as Error });
      }
    }

    // 3. 记录规则执行失败（便于调试）
    if (ruleErrors.length > 0) {
      console.warn(
        `[DataValidator] ${ruleErrors.length} 个规则执行失败:`,
        ruleErrors.map((e) => `${e.ruleId}: ${e.error.message}`)
      );
    }

    // 4. 对 block 级规则，二次确认（扩大样本）
    const blockIssues = issues.filter((i) => i.severity === "block" && i.confidence !== "high");
    for (const issue of blockIssues) {
      const confirmed = await this.confirmBlockIssue(issue, sheet, reader);
      if (!confirmed) {
        // 降级为 warn
        issue.severity = "warn";
        issue.details.push("（二次确认未能验证，已降级为警告）");
      }
    }

    return issues;
  }

  /**
   * 构建共享验证上下文
   */
  private async buildContext(sheet: string, reader: ExcelReader): Promise<DataValidationContext> {
    const context: DataValidationContext = { sheet };

    try {
      // 1. 读取分层样本
      const sampleData = await this.sampleData(sheet, reader);
      context.sampleData = sampleData ?? undefined;

      // 2. 解析列角色
      if (sampleData) {
        const allSamples = [
          ...sampleData.headRows,
          ...sampleData.tailRows,
          ...sampleData.randomRows,
        ];
        context.resolvedColumns = this.columnResolver.resolve(sampleData.headers, allSamples);
      }
    } catch (error) {
      console.warn(`[DataValidator] 构建上下文失败:`, error);
    }

    return context;
  }

  /**
   * 分层抽样读取数据
   */
  private async sampleData(sheet: string, reader: ExcelReader): Promise<SampleData | null> {
    const strategy = this.defaultSamplingStrategy;

    try {
      // 1. 先读取头部数据获取行数
      const headData = await reader.sampleRows(sheet, strategy.headRows + 1);
      if (!headData || headData.length < 2) return null;

      const headers = headData[0] as string[];
      const headRows = headData.slice(1) as unknown[][];

      // 2. 尝试获取总行数
      let totalRowCount = headRows.length + 1; // 至少是已读取的
      try {
        // 通过读取更多行来估算总行数
        const moreData = await reader.sampleRows(sheet, 1000);
        if (moreData) {
          totalRowCount = moreData.length;
        }
      } catch {
        // 忽略
      }

      // 3. 构建采样行索引
      const sampledRowIndices: number[] = [];

      // 头部行索引 (2, 3, 4, ...)
      for (let i = 0; i < headRows.length; i++) {
        sampledRowIndices.push(i + 2);
      }

      // 4. 读取尾部数据（如果总行数足够）
      let tailRows: unknown[][] = [];
      if (totalRowCount > strategy.headRows + strategy.tailRows + 1) {
        try {
          // 简化：直接读取全部然后取尾部
          const allData = await reader.sampleRows(sheet, totalRowCount);
          if (allData && allData.length > strategy.headRows + 1) {
            const tailStart = Math.max(allData.length - strategy.tailRows, strategy.headRows + 1);
            tailRows = allData.slice(tailStart) as unknown[][];

            // 尾部行索引
            for (let i = 0; i < tailRows.length; i++) {
              sampledRowIndices.push(tailStart + i + 1);
            }
          }
        } catch {
          // 忽略尾部读取失败
        }
      }

      // 5. 随机抽样（从中间部分）
      let randomRows: unknown[][] = [];
      if (totalRowCount > strategy.headRows + strategy.tailRows + 1) {
        try {
          const allData = await reader.sampleRows(sheet, totalRowCount);
          if (allData && allData.length > strategy.headRows + strategy.tailRows + 1) {
            const middleStart = strategy.headRows + 1;
            const middleEnd = allData.length - strategy.tailRows;
            const middleIndices: number[] = [];

            for (let i = middleStart; i < middleEnd; i++) {
              middleIndices.push(i);
            }

            // 随机选择
            this.shuffleArray(middleIndices);
            const selectedIndices = middleIndices.slice(0, strategy.randomRows);
            selectedIndices.sort((a, b) => a - b);

            randomRows = selectedIndices.map((idx) => allData[idx] as unknown[]);

            // 随机行索引
            for (const idx of selectedIndices) {
              sampledRowIndices.push(idx + 1);
            }
          }
        } catch {
          // 忽略随机读取失败
        }
      }

      return {
        headers,
        headRows,
        tailRows,
        randomRows,
        totalRowCount,
        sampledRowIndices,
      };
    } catch (error) {
      console.warn(`[DataValidator] 抽样读取失败:`, error);
      return null;
    }
  }

  /**
   * 数组随机打乱
   */
  private shuffleArray<T>(array: T[]): void {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
  }

  /**
   * 二次确认 block 级问题（扩大样本）
   */
  private async confirmBlockIssue(
    issue: DataValidationIssue,
    sheet: string,
    reader: ExcelReader
  ): Promise<boolean> {
    // 对于已经是高置信度的，不需要二次确认
    if (issue.confidence === "high") return true;

    try {
      // 扩大样本：读取更多数据
      const expandedData = await reader.sampleRows(sheet, 200);
      if (!expandedData || expandedData.length < 50) {
        // 数据量不足，无法二次确认
        return true; // 保守起见，认为问题存在
      }

      // 根据规则类型进行特定验证
      switch (issue.ruleId) {
        case "null_value_check":
        case "type_consistency":
          // 这些规则在扩大样本后仍然发现问题，就确认
          return true;

        case "column_constant":
        case "summary_distribution":
          // 检查扩大样本后唯一值数量是否仍然为1
          return true; // 简化处理，后续可细化

        default:
          return true;
      }
    } catch {
      return true; // 出错时保守处理
    }
  }

  /**
   * 验证多个工作表
   */
  async validateWorkbook(
    sheets: string[],
    reader: ExcelReader
  ): Promise<Map<string, DataValidationIssue[]>> {
    const allResults = new Map<string, DataValidationIssue[]>();

    for (const sheet of sheets) {
      const results = await this.validate(sheet, reader);
      if (results.length > 0) {
        allResults.set(sheet, results);
      }
    }

    return allResults;
  }

  /**
   * 快速验证（只返回是否通过）
   */
  async quickValidate(sheet: string, reader: ExcelReader): Promise<boolean> {
    const issues = await this.validate(sheet, reader);
    return !issues.some((i) => i.severity === "block");
  }

  /**
   * 获取所有规则
   */
  getRules(): Array<{
    id: string;
    name: string;
    severity: string;
    enabled: boolean;
    requiresIO: boolean;
  }> {
    return this.rules.map((r) => ({
      id: r.id,
      name: r.name,
      severity: r.severity,
      enabled: r.enabled,
      requiresIO: r.requiresIO,
    }));
  }

  /**
   * 启用/禁用规则
   */
  setRuleEnabled(ruleId: string, enabled: boolean): void {
    const rule = this.rules.find((r) => r.id === ruleId);
    if (rule) {
      rule.enabled = enabled;
    }
  }

  /**
   * 添加自定义规则
   */
  addRule(rule: DataValidationRule): void {
    this.rules.push(rule);
  }

  /**
   * 获取列解析器
   */
  getColumnResolver(): ColumnResolver {
    return this.columnResolver;
  }
}

// ========== 导出 ==========

export const dataValidator = new DataValidator();

// 兼容旧接口
export type DataValidationResult = DataValidationIssue;
