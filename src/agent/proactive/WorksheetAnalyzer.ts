/**
 * WorksheetAnalyzer - 工作表智能分析器
 *
 * 主动扫描表格结构，形成洞察判断
 * 像人一样先观察、理解、再建议
 *
 * @module agent/proactive/WorksheetAnalyzer
 */

// 注意：这里直接使用 Excel.run 而不是 excelRun 封装
// 因为我们需要返回 WorksheetAnalysis 而不是 ToolResult

// ========== 类型定义 ==========

/**
 * 单元格信息
 */
export interface CellInfo {
  address: string;
  value: unknown;
  text: string;
  numberFormat: string;
  hasFormula: boolean;
  formula?: string;
  isMerged: boolean;
  isHeader: boolean;
  dataType: "empty" | "text" | "number" | "date" | "boolean" | "error" | "mixed";
  // 格式问题
  isTextFormattedNumber: boolean; // 数字被存成文本
  hasInconsistentFormat: boolean;
}

/**
 * 列分析结果
 */
export interface ColumnAnalysis {
  index: number;
  letter: string;
  header: string | null;
  dataType: "text" | "number" | "date" | "mixed" | "empty";
  hasHeader: boolean;
  totalRows: number;
  emptyCount: number;
  fillRate: number; // 0-1，非空比例
  // 格式问题
  textNumberCount: number; // 被存成文本的数字
  formatInconsistencyRate: number; // 格式不一致率
  uniqueFormats: string[];
  // 内容特征
  seemsLikeDate: boolean;
  seemsLikeCurrency: boolean;
  seemsLikePercentage: boolean;
  seemsLikeID: boolean;
  sampleValues: string[];
}

/**
 * 行分析结果
 */
export interface RowAnalysis {
  index: number;
  isEmpty: boolean;
  isHeader: boolean;
  isSubtotal: boolean;
  cellCount: number;
  filledCount: number;
}

/**
 * 合并单元格信息
 */
export interface MergedCellInfo {
  address: string;
  rowCount: number;
  columnCount: number;
  value: unknown;
}

/**
 * 数据区域信息
 */
export interface DataRegion {
  address: string;
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  rowCount: number;
  columnCount: number;
  hasHeader: boolean;
  headerRow: number | null;
  dataStartRow: number;
}

/**
 * 表格结构类型
 */
export type TableStructure =
  | "simple_list"      // 简单列表（无表头）
  | "standard_table"   // 标准表格（有表头）
  | "multi_header"     // 多行表头
  | "pivot_style"      // 透视表风格
  | "matrix"           // 矩阵（行列都有标签）
  | "free_form"        // 自由格式（不规则）
  | "empty";           // 空表

/**
 * 完整分析结果
 */
export interface WorksheetAnalysis {
  // 基本信息
  sheetName: string;
  usedRange: string;
  analyzedAt: Date;

  // 区域分析
  dataRegion: DataRegion;
  structure: TableStructure;

  // 列分析
  columns: ColumnAnalysis[];
  totalColumns: number;

  // 行分析
  totalRows: number;
  emptyRowIndices: number[];
  headerRowIndex: number | null;

  // 问题检测
  issues: AnalysisIssue[];

  // 合并单元格
  mergedCells: MergedCellInfo[];
  hasMergedCells: boolean;

  // 数据质量
  overallQuality: number; // 0-100
  qualityFactors: {
    formatConsistency: number;
    dataCompleteness: number;
    structureClarity: number;
    typeCorrectness: number;
  };
}

/**
 * 分析发现的问题
 */
export interface AnalysisIssue {
  type: IssueType;
  severity: "low" | "medium" | "high";
  location: string; // 位置描述
  description: string;
  affectedRange?: string;
  suggestedFix: string;
  autoFixable: boolean;
}

export type IssueType =
  | "text_formatted_numbers"    // 数字存成文本
  | "inconsistent_format"       // 格式不一致
  | "merged_cells"              // 合并单元格
  | "empty_rows"                // 空行
  | "empty_columns"             // 空列
  | "missing_header"            // 缺少表头
  | "mixed_data_types"          // 混合数据类型
  | "header_data_mixed"         // 表头和数据混在一起
  | "irregular_structure"       // 不规则结构
  | "duplicate_headers"         // 重复的表头名
  | "date_as_text"              // 日期存成文本
  | "trailing_spaces"           // 尾部空格
  | "hidden_characters";        // 隐藏字符

// ========== 分析器类 ==========

export class WorksheetAnalyzer {
  private maxRowsToScan = 1000;
  private maxColumnsToScan = 50;
  private sampleSize = 100; // 采样行数

  constructor(options?: { maxRows?: number; maxColumns?: number; sampleSize?: number }) {
    if (options?.maxRows) this.maxRowsToScan = options.maxRows;
    if (options?.maxColumns) this.maxColumnsToScan = options.maxColumns;
    if (options?.sampleSize) this.sampleSize = options.sampleSize;
  }

  /**
   * 静默分析当前工作表
   */
  async analyze(sheetName?: string): Promise<WorksheetAnalysis> {
    return await Excel.run(async (ctx) => {
      const sheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();

      sheet.load("name");
      const usedRange = sheet.getUsedRange();
      usedRange.load(["address", "rowCount", "columnCount", "values", "numberFormat", "text"]);

      // 加载合并单元格
      // Note: Excel JS API 对合并单元格支持有限，这里简化处理

      await ctx.sync();

      const sheetNameResult = sheet.name;
      const address = usedRange.address;
      const values = usedRange.values as unknown[][];
      const formats = usedRange.numberFormat as string[][];
      const texts = usedRange.text as string[][];
      const rowCount = Math.min(usedRange.rowCount, this.maxRowsToScan);
      const colCount = Math.min(usedRange.columnCount, this.maxColumnsToScan);

      // 分析数据
      const columns = this.analyzeColumns(values, formats, texts, rowCount, colCount);
      const headerRowIndex = this.detectHeaderRow(values, columns, rowCount, colCount);
      const emptyRowIndices = this.findEmptyRows(values, rowCount, colCount);
      const structure = this.detectStructure(values, columns, headerRowIndex, rowCount, colCount);
      const issues = this.detectIssues(values, formats, texts, columns, headerRowIndex, rowCount, colCount);
      const mergedCells = await this.detectMergedCells(sheet, ctx);
      const dataRegion = this.calculateDataRegion(address, headerRowIndex, rowCount, colCount);
      const qualityFactors = this.calculateQualityFactors(columns, issues, headerRowIndex);

      return {
        sheetName: sheetNameResult,
        usedRange: address,
        analyzedAt: new Date(),
        dataRegion,
        structure,
        columns,
        totalColumns: colCount,
        totalRows: rowCount,
        emptyRowIndices,
        headerRowIndex,
        issues,
        mergedCells,
        hasMergedCells: mergedCells.length > 0,
        overallQuality: Math.round(
          (qualityFactors.formatConsistency +
            qualityFactors.dataCompleteness +
            qualityFactors.structureClarity +
            qualityFactors.typeCorrectness) / 4
        ),
        qualityFactors,
      };
    });
  }

  /**
   * 分析列
   */
  private analyzeColumns(
    values: unknown[][],
    formats: string[][],
    texts: string[][],
    rowCount: number,
    colCount: number
  ): ColumnAnalysis[] {
    const columns: ColumnAnalysis[] = [];

    for (let c = 0; c < colCount; c++) {
      const columnLetter = this.columnIndexToLetter(c);
      let emptyCount = 0;
      let textNumberCount = 0;
      const formatSet = new Set<string>();
      const sampleValues: string[] = [];
      let hasNumeric = false;
      let hasText = false;
      let hasDate = false;

      for (let r = 0; r < rowCount; r++) {
        const val = values[r]?.[c];
        const fmt = formats[r]?.[c] || "General";
        const txt = texts[r]?.[c] || "";

        if (val === null || val === undefined || val === "") {
          emptyCount++;
          continue;
        }

        formatSet.add(fmt);

        // 采样值
        if (sampleValues.length < 5 && r > 0) {
          sampleValues.push(String(val).substring(0, 50));
        }

        // 检测数据类型
        if (typeof val === "number") {
          hasNumeric = true;
        } else if (typeof val === "string") {
          // 检查是否是文本格式的数字
          if (this.looksLikeNumber(val)) {
            textNumberCount++;
            hasNumeric = true;
          } else if (this.looksLikeDate(val)) {
            hasDate = true;
          } else {
            hasText = true;
          }
        }
      }

      // 判断列类型
      let dataType: ColumnAnalysis["dataType"] = "empty";
      if (emptyCount < rowCount) {
        if (hasNumeric && !hasText && !hasDate) dataType = "number";
        else if (hasDate && !hasText && !hasNumeric) dataType = "date";
        else if (hasText && !hasNumeric && !hasDate) dataType = "text";
        else dataType = "mixed";
      }

      // 获取表头（假设第一行是表头）
      const headerValue = values[0]?.[c];
      const header = typeof headerValue === "string" ? headerValue : null;

      columns.push({
        index: c,
        letter: columnLetter,
        header,
        dataType,
        hasHeader: header !== null && header !== "",
        totalRows: rowCount,
        emptyCount,
        fillRate: (rowCount - emptyCount) / rowCount,
        textNumberCount,
        formatInconsistencyRate: formatSet.size > 1 ? (formatSet.size - 1) / formatSet.size : 0,
        uniqueFormats: Array.from(formatSet),
        seemsLikeDate: hasDate || this.columnNameSuggestsDate(header),
        seemsLikeCurrency: this.columnNameSuggestsCurrency(header) || this.formatSuggestsCurrency(Array.from(formatSet)),
        seemsLikePercentage: this.formatSuggestsPercentage(Array.from(formatSet)),
        seemsLikeID: this.columnNameSuggestsID(header),
        sampleValues,
      });
    }

    return columns;
  }

  /**
   * 检测表头行
   */
  private detectHeaderRow(
    values: unknown[][],
    columns: ColumnAnalysis[],
    rowCount: number,
    colCount: number
  ): number | null {
    if (rowCount === 0) return null;

    // 检查第一行是否全是文本
    const firstRow = values[0];
    if (!firstRow) return null;

    let textCount = 0;
    let numericCount = 0;

    for (let c = 0; c < colCount; c++) {
      const val = firstRow[c];
      if (typeof val === "string" && val.trim() !== "") {
        textCount++;
      } else if (typeof val === "number") {
        numericCount++;
      }
    }

    // 如果第一行大部分是文本，且第二行有数值，则认为第一行是表头
    if (textCount > colCount * 0.5 && rowCount > 1) {
      const secondRow = values[1];
      let secondRowNumeric = 0;
      for (let c = 0; c < colCount; c++) {
        if (typeof secondRow?.[c] === "number") secondRowNumeric++;
      }
      if (secondRowNumeric > 0 || textCount > numericCount) {
        return 0;
      }
    }

    return null;
  }

  /**
   * 检测表格结构
   */
  private detectStructure(
    values: unknown[][],
    columns: ColumnAnalysis[],
    headerRowIndex: number | null,
    rowCount: number,
    colCount: number
  ): TableStructure {
    if (rowCount === 0 || colCount === 0) return "empty";

    const hasHeader = headerRowIndex !== null;
    const columnsWithHeaders = columns.filter((c) => c.hasHeader).length;

    // 检查是否是矩阵（第一列也是标签）
    const firstColumnAllText = columns[0]?.dataType === "text";
    if (hasHeader && firstColumnAllText && colCount > 2) {
      return "matrix";
    }

    // 检查是否是标准表格
    if (hasHeader && columnsWithHeaders > colCount * 0.7) {
      return "standard_table";
    }

    // 检查是否是简单列表
    if (!hasHeader && colCount <= 3) {
      return "simple_list";
    }

    // 检查是否是透视表风格
    const hasSubtotals = this.hasSubtotalRows(values, rowCount, colCount);
    if (hasSubtotals) {
      return "pivot_style";
    }

    return "free_form";
  }

  /**
   * 检测问题
   */
  private detectIssues(
    values: unknown[][],
    formats: string[][],
    texts: string[][],
    columns: ColumnAnalysis[],
    headerRowIndex: number | null,
    rowCount: number,
    colCount: number
  ): AnalysisIssue[] {
    const issues: AnalysisIssue[] = [];

    // 1. 检测文本格式的数字
    for (const col of columns) {
      if (col.textNumberCount > 0 && col.textNumberCount > col.totalRows * 0.1) {
        issues.push({
          type: "text_formatted_numbers",
          severity: "high",
          location: `${col.letter} 列`,
          description: `${col.textNumberCount} 个数值被存储为文本格式`,
          affectedRange: `${col.letter}:${col.letter}`,
          suggestedFix: "将文本转换为数值格式",
          autoFixable: true,
        });
      }
    }

    // 2. 检测格式不一致
    for (const col of columns) {
      if (col.formatInconsistencyRate > 0.3) {
        issues.push({
          type: "inconsistent_format",
          severity: "medium",
          location: `${col.letter} 列`,
          description: `存在 ${col.uniqueFormats.length} 种不同格式`,
          affectedRange: `${col.letter}:${col.letter}`,
          suggestedFix: "统一列格式",
          autoFixable: true,
        });
      }
    }

    // 3. 检测缺少表头
    if (headerRowIndex === null && rowCount > 5) {
      issues.push({
        type: "missing_header",
        severity: "medium",
        location: "整个表格",
        description: "未检测到表头行",
        suggestedFix: "添加描述性表头",
        autoFixable: false,
      });
    }

    // 4. 检测混合数据类型
    for (const col of columns) {
      if (col.dataType === "mixed" && col.header) {
        issues.push({
          type: "mixed_data_types",
          severity: "medium",
          location: `${col.letter} 列 (${col.header})`,
          description: "列中包含混合数据类型",
          affectedRange: `${col.letter}:${col.letter}`,
          suggestedFix: "统一数据类型或拆分列",
          autoFixable: false,
        });
      }
    }

    // 5. 检测空行
    const emptyRows = this.findEmptyRows(values, rowCount, colCount);
    if (emptyRows.length > 0 && emptyRows.length < rowCount * 0.1) {
      issues.push({
        type: "empty_rows",
        severity: "low",
        location: `第 ${emptyRows.slice(0, 3).map((r) => r + 1).join(", ")}${emptyRows.length > 3 ? " 等" : ""} 行`,
        description: `存在 ${emptyRows.length} 个空行`,
        suggestedFix: "删除空行以保持数据连续",
        autoFixable: true,
      });
    }

    // 6. 检测空列
    const emptyColumns = columns.filter((c) => c.fillRate < 0.1);
    if (emptyColumns.length > 0) {
      issues.push({
        type: "empty_columns",
        severity: "low",
        location: `${emptyColumns.map((c) => c.letter).join(", ")} 列`,
        description: `存在 ${emptyColumns.length} 个几乎为空的列`,
        suggestedFix: "删除或隐藏空列",
        autoFixable: true,
      });
    }

    return issues;
  }

  /**
   * 检测合并单元格
   */
  private async detectMergedCells(
    sheet: Excel.Worksheet,
    ctx: Excel.RequestContext
  ): Promise<MergedCellInfo[]> {
    // Note: Excel JS API 对合并单元格的检测有限
    // 这里返回空数组，实际实现需要遍历检查
    // 可以通过 range.getMergedAreasOrNullObject() 来检测
    try {
      const usedRange = sheet.getUsedRange();
      usedRange.load("address");
      await ctx.sync();

      // 简化实现：检查是否有合并区域
      // 完整实现需要遍历每个单元格
      return [];
    } catch {
      return [];
    }
  }

  /**
   * 查找空行
   */
  private findEmptyRows(values: unknown[][], rowCount: number, colCount: number): number[] {
    const emptyRows: number[] = [];
    for (let r = 0; r < rowCount; r++) {
      let isEmpty = true;
      for (let c = 0; c < colCount; c++) {
        const val = values[r]?.[c];
        if (val !== null && val !== undefined && val !== "") {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) emptyRows.push(r);
    }
    return emptyRows;
  }

  /**
   * 检测是否有小计行
   */
  private hasSubtotalRows(values: unknown[][], rowCount: number, colCount: number): boolean {
    const subtotalKeywords = ["小计", "合计", "总计", "Subtotal", "Total", "Sum"];
    for (let r = 0; r < rowCount; r++) {
      const firstCell = String(values[r]?.[0] || "");
      if (subtotalKeywords.some((kw) => firstCell.includes(kw))) {
        return true;
      }
    }
    return false;
  }

  /**
   * 计算数据区域
   */
  private calculateDataRegion(
    address: string,
    headerRowIndex: number | null,
    rowCount: number,
    colCount: number
  ): DataRegion {
    // 解析地址获取起始位置
    const match = address.match(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
    const startRow = match ? parseInt(match[2]) : 1;
    const endRow = startRow + rowCount - 1;
    const startColumn = match ? this.letterToColumnIndex(match[1]) : 0;
    const endColumn = startColumn + colCount - 1;

    return {
      address,
      startRow,
      endRow,
      startColumn,
      endColumn,
      rowCount,
      columnCount: colCount,
      hasHeader: headerRowIndex !== null,
      headerRow: headerRowIndex !== null ? startRow + headerRowIndex : null,
      dataStartRow: headerRowIndex !== null ? startRow + headerRowIndex + 1 : startRow,
    };
  }

  /**
   * 计算质量因素
   */
  private calculateQualityFactors(
    columns: ColumnAnalysis[],
    issues: AnalysisIssue[],
    headerRowIndex: number | null
  ): WorksheetAnalysis["qualityFactors"] {
    // 格式一致性
    const avgFormatConsistency = columns.length > 0
      ? columns.reduce((sum, c) => sum + (1 - c.formatInconsistencyRate), 0) / columns.length
      : 1;

    // 数据完整性
    const avgFillRate = columns.length > 0
      ? columns.reduce((sum, c) => sum + c.fillRate, 0) / columns.length
      : 0;

    // 结构清晰度
    const structureClarity = headerRowIndex !== null ? 0.9 : 0.5;

    // 类型正确性
    const mixedColumns = columns.filter((c) => c.dataType === "mixed").length;
    const typeCorrectness = columns.length > 0
      ? 1 - mixedColumns / columns.length
      : 1;

    // 根据问题数量调整
    const issuePenalty = Math.min(issues.length * 0.05, 0.3);

    return {
      formatConsistency: Math.round((avgFormatConsistency - issuePenalty * 0.5) * 100),
      dataCompleteness: Math.round(avgFillRate * 100),
      structureClarity: Math.round((structureClarity - issuePenalty * 0.3) * 100),
      typeCorrectness: Math.round((typeCorrectness - issuePenalty * 0.2) * 100),
    };
  }

  // ========== 辅助方法 ==========

  private columnIndexToLetter(index: number): string {
    let letter = "";
    let temp = index;
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    return letter;
  }

  private letterToColumnIndex(letter: string): number {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
      index = index * 26 + (letter.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  private looksLikeNumber(value: string): boolean {
    if (!value || typeof value !== "string") return false;
    const cleaned = value.replace(/[,\s¥$€£%]/g, "").trim();
    return !isNaN(parseFloat(cleaned)) && isFinite(parseFloat(cleaned));
  }

  private looksLikeDate(value: string): boolean {
    if (!value || typeof value !== "string") return false;
    // 常见日期格式
    const datePatterns = [
      /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/,  // 2024-01-01
      /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/,  // 01/01/2024
      /^\d{4}年\d{1,2}月\d{1,2}日$/,    // 2024年1月1日
    ];
    return datePatterns.some((p) => p.test(value.trim()));
  }

  private columnNameSuggestsDate(name: string | null): boolean {
    if (!name) return false;
    const dateKeywords = ["日期", "时间", "date", "time", "年", "月", "created", "updated"];
    return dateKeywords.some((kw) => name.toLowerCase().includes(kw.toLowerCase()));
  }

  private columnNameSuggestsCurrency(name: string | null): boolean {
    if (!name) return false;
    const currencyKeywords = ["金额", "价格", "费用", "收入", "支出", "amount", "price", "cost", "revenue"];
    return currencyKeywords.some((kw) => name.toLowerCase().includes(kw.toLowerCase()));
  }

  private columnNameSuggestsID(name: string | null): boolean {
    if (!name) return false;
    const idKeywords = ["编号", "ID", "编码", "序号", "code", "number", "no."];
    return idKeywords.some((kw) => name.toLowerCase().includes(kw.toLowerCase()));
  }

  private formatSuggestsCurrency(formats: string[]): boolean {
    return formats.some((f) => /[¥$€£]/.test(f) || f.includes("Currency"));
  }

  private formatSuggestsPercentage(formats: string[]): boolean {
    return formats.some((f) => f.includes("%"));
  }
}

// ========== 导出工厂函数 ==========

export function createWorksheetAnalyzer(options?: {
  maxRows?: number;
  maxColumns?: number;
  sampleSize?: number;
}): WorksheetAnalyzer {
  return new WorksheetAnalyzer(options);
}
