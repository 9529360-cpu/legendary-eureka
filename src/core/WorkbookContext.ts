/**
 * WorkbookContext - Excel 工作簿感知层
 *
 * 功能：
 * 1. 结构感知 - 了解工作簿的整体结构
 * 2. 数据感知 - 自动识别数据区域和表头
 * 3. 关系感知 - 理解公式依赖关系
 * 4. 实时感知 - 追踪用户操作和选区变化
 *
 * 设计目标：让 AI 像大脑了解身体一样了解 Excel
 */

/* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */

// ==================== 类型定义 ====================

/**
 * 工作表信息
 */
export interface SheetInfo {
  name: string;
  id: string;
  position: number;
  visibility: "visible" | "hidden" | "veryHidden";
  usedRange: RangeInfo | null;
  rowCount: number;
  columnCount: number;
  hasData: boolean;
  tables: string[];
  charts: number;
}

/**
 * 范围信息
 */
export interface RangeInfo {
  address: string;
  addressLocal: string;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  isEmpty: boolean;
  hasFormulas: boolean;
  hasValues: boolean;
}

/**
 * 表格信息 (Excel Table 对象)
 */
export interface TableInfo {
  id: string;
  name: string;
  sheetName: string;
  range: string;
  headerRowRange: string;
  dataBodyRange: string;
  totalRowRange: string | null;
  columns: TableColumnInfo[];
  rowCount: number;
  showHeaders: boolean;
  showTotals: boolean;
  style: string;
}

/**
 * 表格列信息
 */
export interface TableColumnInfo {
  id: number;
  name: string;
  index: number;
  dataType: "text" | "number" | "date" | "boolean" | "mixed" | "unknown";
  hasFormula: boolean;
  sampleValues: any[];
}

/**
 * 命名范围信息
 */
export interface NamedRangeInfo {
  name: string;
  scope: "workbook" | "worksheet";
  scopeSheet?: string;
  reference: string;
  value: any;
  isValid: boolean;
}

/**
 * 数据区域 (自动识别)
 */
export interface DataRegion {
  id: string;
  sheetName: string;
  range: string;
  type: "table" | "list" | "matrix" | "single" | "sparse";
  hasHeaders: boolean;
  headerRow?: number;
  headers?: string[];
  dataRows: number;
  dataColumns: number;
  confidence: number;
}

/**
 * 公式依赖信息
 */
export interface FormulaDependency {
  cell: string;
  formula: string;
  precedents: string[]; // 该单元格依赖的单元格
  dependents: string[]; // 依赖该单元格的单元格
  isVolatile: boolean;
  hasExternalRef: boolean;
}

/**
 * 最近变更记录
 */
export interface RecentChange {
  id: string;
  timestamp: Date;
  type: "value" | "format" | "structure" | "selection";
  sheetName: string;
  range: string;
  oldValue?: any;
  newValue?: any;
  description: string;
}

/**
 * 完整的工作簿上下文
 */
export interface WorkbookContextData {
  // 基本信息
  workbookName: string;
  lastModified: Date;

  // 结构感知
  sheets: SheetInfo[];
  activeSheet: string;
  tables: TableInfo[];
  namedRanges: NamedRangeInfo[];

  // 数据感知
  dataRegions: DataRegion[];

  // 实时感知
  currentSelection: RangeInfo | null;
  selectionValues: any[][] | null;

  // 公式感知
  formulaDependencies: FormulaDependency[];

  // 变更追踪
  recentChanges: RecentChange[];

  // 元数据
  capturedAt: Date;
  captureDepth: "shallow" | "medium" | "deep";
}

// ==================== 工作簿上下文类 ====================

/**
 * WorkbookContext 类 - Excel 感知核心
 */
export class WorkbookContext {
  private context: Excel.RequestContext;
  private cachedContext: WorkbookContextData | null = null;
  private lastCaptureTime: Date | null = null;
  private cacheValidityMs: number = 5000; // 缓存有效期 5 秒
  private recentChanges: RecentChange[] = [];
  private maxRecentChanges: number = 50;

  constructor(context: Excel.RequestContext) {
    this.context = context;
  }

  // ==================== 主要 API ====================

  /**
   * 获取完整的工作簿上下文
   */
  async getFullContext(
    depth: "shallow" | "medium" | "deep" = "medium"
  ): Promise<WorkbookContextData> {
    // 检查缓存
    if (this.isCacheValid()) {
      return this.cachedContext!;
    }

    const contextData: WorkbookContextData = {
      workbookName: "",
      lastModified: new Date(),
      sheets: [],
      activeSheet: "",
      tables: [],
      namedRanges: [],
      dataRegions: [],
      currentSelection: null,
      selectionValues: null,
      formulaDependencies: [],
      recentChanges: [...this.recentChanges],
      capturedAt: new Date(),
      captureDepth: depth,
    };

    try {
      // 获取工作簿基本信息
      const workbook = this.context.workbook;
      workbook.load("name");

      // 获取活动工作表
      const activeSheet = workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");

      // 获取所有工作表
      const sheets = workbook.worksheets;
      sheets.load("items/name, items/id, items/position, items/visibility");

      // 获取当前选区
      const selection = workbook.getSelectedRange();
      selection.load("address, addressLocal, rowCount, columnCount, cellCount, values, formulas");

      // 获取命名范围
      const names = workbook.names;
      names.load("items/name, items/type, items/value, items/visible");

      await this.context.sync();

      // 填充基本信息

      contextData.workbookName = workbook.name || "未命名工作簿";

      contextData.activeSheet = activeSheet.name;

      // 填充工作表信息

      contextData.sheets = await this.processSheets(sheets.items, depth);

      // 填充选区信息
      contextData.currentSelection = {
        address: selection.address,

        addressLocal: selection.addressLocal,

        rowCount: selection.rowCount,

        columnCount: selection.columnCount,

        cellCount: selection.cellCount,
        isEmpty: false,
        hasFormulas: false,
        hasValues: false,
      };

      contextData.selectionValues = selection.values;

      // 填充命名范围

      contextData.namedRanges = names.items.map((name) => ({
        name: name.name,
        scope: "workbook" as const,
        reference: String(name.value),
        value: name.value,
        isValid: name.visible,
      }));

      // 中等深度：获取表格信息
      if (depth === "medium" || depth === "deep") {
        contextData.tables = await this.getAllTables();
      }

      // 深度：获取数据区域和公式依赖
      if (depth === "deep") {
        contextData.dataRegions = await this.detectDataRegions(contextData.sheets);
        contextData.formulaDependencies = await this.analyzeFormulaDependencies();
      }

      // 更新缓存
      this.cachedContext = contextData;
      this.lastCaptureTime = new Date();

      return contextData;
    } catch (error) {
      console.error("获取工作簿上下文失败:", error);
      throw error;
    }
  }

  /**
   * 获取简化的上下文摘要（用于 AI Prompt）
   */
  async getContextSummary(): Promise<string> {
    const ctx = await this.getFullContext("medium");

    const lines: string[] = [];

    lines.push(`## 工作簿: ${ctx.workbookName}`);
    lines.push("");

    // 工作表摘要
    lines.push(`### 工作表 (${ctx.sheets.length} 个)`);
    for (const sheet of ctx.sheets) {
      const tableCount = sheet.tables.length > 0 ? `, ${sheet.tables.length} 个表格` : "";
      const chartCount = sheet.charts > 0 ? `, ${sheet.charts} 个图表` : "";
      const activeMarker = sheet.name === ctx.activeSheet ? " [当前]" : "";
      lines.push(
        `- **${sheet.name}**${activeMarker}: ${sheet.rowCount}行 × ${sheet.columnCount}列${tableCount}${chartCount}`
      );
    }
    lines.push("");

    // 当前选区
    if (ctx.currentSelection) {
      lines.push(`### 当前选区`);
      lines.push(`- 地址: ${ctx.currentSelection.address}`);
      lines.push(
        `- 大小: ${ctx.currentSelection.rowCount}行 × ${ctx.currentSelection.columnCount}列`
      );

      // 显示选区数据预览
      if (ctx.selectionValues && ctx.selectionValues.length > 0) {
        const previewRows = Math.min(5, ctx.selectionValues.length);
        lines.push(`- 数据预览:`);
        for (let i = 0; i < previewRows; i++) {
          const row = ctx.selectionValues[i]
            .slice(0, 5)
            .map((v) => String(v).slice(0, 15))
            .join(" | ");
          lines.push(`  ${row}`);
        }
        if (ctx.selectionValues.length > 5) {
          lines.push(`  ... 还有 ${ctx.selectionValues.length - 5} 行`);
        }
      }
      lines.push("");
    }

    // 表格摘要
    if (ctx.tables.length > 0) {
      lines.push(`### Excel 表格 (${ctx.tables.length} 个)`);
      for (const table of ctx.tables) {
        lines.push(
          `- **${table.name}** (${table.sheetName}): ${table.rowCount}行, 列: ${table.columns.map((c) => c.name).join(", ")}`
        );
      }
      lines.push("");
    }

    // 命名范围
    if (ctx.namedRanges.length > 0) {
      lines.push(`### 命名范围 (${ctx.namedRanges.length} 个)`);
      for (const nr of ctx.namedRanges.slice(0, 10)) {
        lines.push(`- ${nr.name}: ${nr.reference}`);
      }
      if (ctx.namedRanges.length > 10) {
        lines.push(`... 还有 ${ctx.namedRanges.length - 10} 个`);
      }
      lines.push("");
    }

    return lines.join("\n");
  }

  /**
   * 快速获取当前选区信息
   */
  async getCurrentSelection(): Promise<{
    range: RangeInfo;
    values: any[][];
    formulas: any[][];
    sheetName: string;
  } | null> {
    try {
      const workbook = this.context.workbook;
      const activeSheet = workbook.worksheets.getActiveWorksheet();
      const selection = workbook.getSelectedRange();

      activeSheet.load("name");
      selection.load("address, addressLocal, rowCount, columnCount, cellCount, values, formulas");

      await this.context.sync();

      /* eslint-disable office-addins/call-sync-before-read */
      return {
        range: {
          address: selection.address,
          addressLocal: selection.addressLocal,
          rowCount: selection.rowCount,
          columnCount: selection.columnCount,
          cellCount: selection.cellCount,
          isEmpty: selection.values.every((row) =>
            row.every((cell) => cell === "" || cell === null)
          ),
          hasFormulas: selection.formulas.some((row) =>
            row.some((cell) => String(cell).startsWith("="))
          ),
          hasValues: selection.values.some((row) =>
            row.some((cell) => cell !== "" && cell !== null)
          ),
        },
        values: selection.values,
        formulas: selection.formulas,
        sheetName: activeSheet.name,
      };
      /* eslint-enable office-addins/call-sync-before-read */
    } catch (error) {
      console.error("获取选区失败:", error);
      return null;
    }
  }

  /**
   * 检测数据表的表头
   */
  async detectHeaders(rangeAddress: string): Promise<{
    hasHeaders: boolean;
    headers: string[];
    confidence: number;
  }> {
    try {
      const sheet = this.context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      range.load("values, rowCount");
      await this.context.sync();

      // eslint-disable-next-line office-addins/call-sync-before-read
      const values = range.values;
      if (values.length === 0) {
        return { hasHeaders: false, headers: [], confidence: 0 };
      }

      const firstRow = values[0];
      const secondRow = values.length > 1 ? values[1] : null;

      // 表头检测规则
      let headerScore = 0;

      // 1. 第一行全是文本
      const firstRowAllText = firstRow.every((cell) => typeof cell === "string" && cell !== "");
      if (firstRowAllText) headerScore += 30;

      // 2. 第一行没有数字，第二行有数字
      const firstRowNoNumbers = firstRow.every((cell) => typeof cell !== "number");
      const secondRowHasNumbers = secondRow?.some((cell) => typeof cell === "number");
      if (firstRowNoNumbers && secondRowHasNumbers) headerScore += 30;

      // 3. 第一行值唯一（没有重复）
      const uniqueValues = new Set(firstRow.map(String));
      if (uniqueValues.size === firstRow.length) headerScore += 20;

      // 4. 第一行值看起来像标题（没有特殊字符，不太长）
      const lookLikeHeaders = firstRow.every((cell) => {
        const str = String(cell);
        return str.length < 50 && !/^\d+$/.test(str);
      });
      if (lookLikeHeaders) headerScore += 20;

      const hasHeaders = headerScore >= 50;
      const headers = hasHeaders ? firstRow.map(String) : [];

      return {
        hasHeaders,
        headers,
        confidence: headerScore / 100,
      };
    } catch (error) {
      console.error("检测表头失败:", error);
      return { hasHeaders: false, headers: [], confidence: 0 };
    }
  }

  /**
   * 获取指定单元格的公式依赖
   */
  async getCellDependencies(cellAddress: string): Promise<FormulaDependency | null> {
    try {
      const sheet = this.context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress);
      cell.load("formulas, address");
      await this.context.sync();

      // eslint-disable-next-line office-addins/call-sync-before-read
      const formula = cell.formulas[0][0];
      if (!formula || !String(formula).startsWith("=")) {
        return null;
      }

      // 解析公式中的单元格引用
      const precedents = this.extractCellReferences(String(formula));

      return {
        // eslint-disable-next-line office-addins/call-sync-before-read
        cell: cell.address,
        formula: String(formula),
        precedents,
        dependents: [], // 需要反向分析
        isVolatile: this.isVolatileFormula(String(formula)),
        hasExternalRef: formula.includes("[") || formula.includes("!"),
      };
    } catch (error) {
      console.error("获取公式依赖失败:", error);
      return null;
    }
  }

  /**
   * 记录变更
   */
  recordChange(change: Omit<RecentChange, "id" | "timestamp">): void {
    const record: RecentChange = {
      ...change,
      id: `change_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
      timestamp: new Date(),
    };

    this.recentChanges.unshift(record);

    if (this.recentChanges.length > this.maxRecentChanges) {
      this.recentChanges = this.recentChanges.slice(0, this.maxRecentChanges);
    }
  }

  /**
   * 清除缓存
   */
  invalidateCache(): void {
    this.cachedContext = null;
    this.lastCaptureTime = null;
  }

  // ==================== 私有方法 ====================

  private isCacheValid(): boolean {
    if (!this.cachedContext || !this.lastCaptureTime) {
      return false;
    }
    const elapsed = Date.now() - this.lastCaptureTime.getTime();
    return elapsed < this.cacheValidityMs;
  }

  private async processSheets(
    sheets: Excel.Worksheet[],
    depth: "shallow" | "medium" | "deep"
  ): Promise<SheetInfo[]> {
    const result: SheetInfo[] = [];

    for (const sheet of sheets) {
      const info: SheetInfo = {
        name: sheet.name,
        id: sheet.id,
        position: sheet.position,
        visibility: sheet.visibility as SheetInfo["visibility"],
        usedRange: null,
        rowCount: 0,
        columnCount: 0,
        hasData: false,
        tables: [],
        charts: 0,
      };

      if (depth !== "shallow") {
        try {
          // 获取已使用范围
          const usedRange = sheet.getUsedRange();
          usedRange.load("address, rowCount, columnCount");

          // 获取表格
          const tables = sheet.tables;
          tables.load("items/name");

          // 获取图表
          const charts = sheet.charts;
          charts.load("count");

          await this.context.sync();

          /* eslint-disable office-addins/call-sync-before-read */
          info.usedRange = {
            address: usedRange.address,
            addressLocal: usedRange.address,
            rowCount: usedRange.rowCount,
            columnCount: usedRange.columnCount,
            cellCount: usedRange.rowCount * usedRange.columnCount,
            isEmpty: false,
            hasFormulas: false,
            hasValues: true,
          };
          info.rowCount = usedRange.rowCount;
          info.columnCount = usedRange.columnCount;
          info.hasData = usedRange.rowCount > 0 && usedRange.columnCount > 0;
          info.tables = tables.items.map((t) => t.name);
          info.charts = charts.count;
          /* eslint-enable office-addins/call-sync-before-read */
        } catch {
          // 工作表可能为空
          info.hasData = false;
        }
      }

      result.push(info);
    }

    return result;
  }

  private async getAllTables(): Promise<TableInfo[]> {
    const result: TableInfo[] = [];

    try {
      const sheets = this.context.workbook.worksheets;
      sheets.load("items");
      await this.context.sync();

      for (const sheet of sheets.items) {
        const tables = sheet.tables;
        tables.load("items");
        await this.context.sync();

        for (const table of tables.items) {
          table.load("id, name, showHeaders, showTotals, style");

          const headerRange = table.getHeaderRowRange();
          headerRange.load("address, values");

          const dataRange = table.getDataBodyRange();
          dataRange.load("address, rowCount");

          const columns = table.columns;
          columns.load("items/id, items/name, items/index");

          await this.context.sync();

          /* eslint-disable office-addins/call-sync-before-read */
          const tableInfo: TableInfo = {
            id: table.id,
            name: table.name,
            sheetName: sheet.name,
            range: `${headerRange.address}:${dataRange.address}`,
            headerRowRange: headerRange.address,
            dataBodyRange: dataRange.address,
            totalRowRange: null,
            columns: columns.items.map((col) => ({
              id: col.id,
              name: col.name,
              index: col.index,
              dataType: "unknown" as const,
              hasFormula: false,
              sampleValues: [],
            })),
            rowCount: dataRange.rowCount,
            showHeaders: table.showHeaders,
            showTotals: table.showTotals,
            style: table.style,
          };
          /* eslint-enable office-addins/call-sync-before-read */

          result.push(tableInfo);
        }
      }
    } catch (error) {
      console.error("获取表格失败:", error);
    }

    return result;
  }

  private async detectDataRegions(sheets: SheetInfo[]): Promise<DataRegion[]> {
    const regions: DataRegion[] = [];

    for (const sheet of sheets) {
      if (!sheet.hasData || !sheet.usedRange) continue;

      // 简化实现：将整个已使用范围作为一个数据区域
      const headerDetection = await this.detectHeaders(sheet.usedRange.address);

      regions.push({
        id: `region_${sheet.name}_1`,
        sheetName: sheet.name,
        range: sheet.usedRange.address,
        type: this.classifyDataRegion(sheet),
        hasHeaders: headerDetection.hasHeaders,
        headerRow: headerDetection.hasHeaders ? 1 : undefined,
        headers: headerDetection.headers,
        dataRows: sheet.rowCount - (headerDetection.hasHeaders ? 1 : 0),
        dataColumns: sheet.columnCount,
        confidence: headerDetection.confidence,
      });
    }

    return regions;
  }

  private classifyDataRegion(sheet: SheetInfo): DataRegion["type"] {
    if (sheet.tables.length > 0) return "table";
    if (sheet.columnCount === 1) return "list";
    if (sheet.rowCount === 1) return "single";
    if (sheet.rowCount > 1 && sheet.columnCount > 1) return "matrix";
    return "sparse";
  }

  private async analyzeFormulaDependencies(): Promise<FormulaDependency[]> {
    // 简化实现：只分析当前选区的公式
    const selection = await this.getCurrentSelection();
    if (!selection) return [];

    const dependencies: FormulaDependency[] = [];

    for (let row = 0; row < selection.formulas.length; row++) {
      for (let col = 0; col < selection.formulas[row].length; col++) {
        const formula = selection.formulas[row][col];
        if (formula && String(formula).startsWith("=")) {
          const cellAddress = this.getCellAddress(selection.range.address, row, col);
          dependencies.push({
            cell: cellAddress,
            formula: String(formula),
            precedents: this.extractCellReferences(String(formula)),
            dependents: [],
            isVolatile: this.isVolatileFormula(String(formula)),
            hasExternalRef: String(formula).includes("["),
          });
        }
      }
    }

    return dependencies;
  }

  private extractCellReferences(formula: string): string[] {
    // 提取公式中的单元格引用
    const cellPattern = /\$?[A-Z]+\$?\d+/g;
    const rangePattern = /\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+/g;

    const references: string[] = [];

    let match: RegExpExecArray | null;
    while ((match = rangePattern.exec(formula)) !== null) {
      references.push(match[0]);
    }
    while ((match = cellPattern.exec(formula)) !== null) {
      if (!references.some((r) => r.includes(match![0]))) {
        references.push(match[0]);
      }
    }

    return references;
  }

  private isVolatileFormula(formula: string): boolean {
    const volatileFunctions = ["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"];
    const upperFormula = formula.toUpperCase();
    return volatileFunctions.some((fn) => upperFormula.includes(fn + "("));
  }

  private getCellAddress(baseAddress: string, rowOffset: number, colOffset: number): string {
    // 简化：从基地址计算偏移后的地址
    const match = baseAddress.match(/([A-Z]+)(\d+)/);
    if (!match) return baseAddress;

    const colLetter = match[1];
    const rowNum = parseInt(match[2], 10);

    const newCol = this.offsetColumn(colLetter, colOffset);
    const newRow = rowNum + rowOffset;

    return `${newCol}${newRow}`;
  }

  private offsetColumn(col: string, offset: number): string {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.charCodeAt(i) - 64);
    }
    num += offset;

    let result = "";
    while (num > 0) {
      const remainder = (num - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      num = Math.floor((num - 1) / 26);
    }
    return result || "A";
  }
}

/**
 * 创建工作簿上下文的工厂函数
 */
export function createWorkbookContext(context: Excel.RequestContext): WorkbookContext {
  return new WorkbookContext(context);
}
