/**
 * SpecCompiler - 规格编译器 v4.0
 * 
 * Layer 2: 将高层意图规格编译成工具调用序列
 * 
 * ★★★ 核心设计原则 ★★★
 * 1. 纯 TypeScript 规则，不调用 LLM
 * 2. 零 Token 消耗
 * 3. 自动处理依赖顺序
 * 4. 自动补充感知步骤（写之前必须读）
 * 5. 使用 step.id 作为依赖，不是工具名
 * 
 * @module agent/SpecCompiler
 */

import { IntentSpec, IntentType, CreateTableSpec, WriteDataSpec, FormatSpec, FormulaSpec, ChartSpec, SheetSpec, DataOperationSpec, QuerySpec, ClarifySpec, RespondSpec } from './types/intent';
import { ExecutionPlan, PlanStep } from './TaskPlanner';

// ========== 编译上下文 ==========

/**
 * 编译上下文 (SpecCompiler 专用)
 */
export interface SpecCompileContext {
  /** 当前选区地址 */
  currentSelection?: string;
  
  /** 当前工作表 */
  activeSheet?: string;
  
  /** 可用的表格名 */
  tableNames?: string[];
}

// ========== 编译结果 ==========

/**
 * 编译结果 (SpecCompiler 专用)
 */
export interface SpecCompileResult {
  /** 是否成功 */
  success: boolean;
  
  /** 执行计划 */
  plan?: ExecutionPlan;
  
  /** 错误信息 */
  error?: string;
  
  /** 需要澄清（如果规格不完整） */
  needsClarification?: boolean;
  
  /** 澄清问题 */
  clarificationQuestion?: string;
}

// 向后兼容别名
export type CompileContext = SpecCompileContext;
export type CompileResult = SpecCompileResult;

// ========== SpecCompiler 类 ==========

/**
 * 规格编译器 - 将意图编译成工具调用
 */
export class SpecCompiler {
  private idCounter = 0;
  
  /**
   * 编译意图规格为执行计划
   * 
   * @param spec 意图规格
   * @param context 编译上下文
   * @returns 编译结果
   */
  compile(spec: IntentSpec, context: CompileContext = {}): CompileResult {
    console.log('[SpecCompiler] 编译意图:', spec.intent);
    
    // 如果需要澄清，直接返回澄清步骤
    if (spec.needsClarification) {
      return this.compileClarify(spec);
    }
    
    // 根据意图类型分发到具体编译器
    try {
      switch (spec.intent) {
        case 'create_table':
          return this.compileCreateTable(spec.spec as CreateTableSpec, context);
          
        case 'write_data':
          return this.compileWriteData(spec.spec as WriteDataSpec, context);
          
        case 'format_range':
          return this.compileFormat(spec.spec as FormatSpec, context);
          
        case 'create_formula':
        case 'batch_formula':
        case 'calculate_summary':
          return this.compileFormula(spec.spec as FormulaSpec, context);
          
        case 'create_chart':
          return this.compileChart(spec.spec as ChartSpec, context);
          
        case 'create_sheet':
        case 'switch_sheet':
          return this.compileSheet(spec.spec as SheetSpec, context);
          
        case 'sort_data':
        case 'filter_data':
        case 'remove_duplicates':
        case 'clean_data':
          return this.compileDataOperation(spec.spec as DataOperationSpec, context);
          
        case 'query_data':
        case 'analyze_data':
        case 'lookup_value':
          return this.compileQuery(spec.spec as QuerySpec, context);
          
        case 'clarify':
          return this.compileClarify(spec);
          
        case 'respond_only':
          return this.compileRespond(spec.spec as RespondSpec);
          
        default:
          console.warn('[SpecCompiler] 未知意图类型:', spec.intent);
          return {
            success: false,
            error: `不支持的意图类型: ${spec.intent}`,
          };
      }
    } catch (error) {
      console.error('[SpecCompiler] 编译失败:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }
  
  // ========== 具体编译器 ==========
  
  /**
   * 编译创建表格
   */
  private compileCreateTable(spec: CreateTableSpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    
    // 步骤1: 读取当前选区（感知）
    const readStepId = this.generateId();
    steps.push({
      id: readStepId,
      order: 1,
      phase: 'sensing',
      description: '读取当前选区位置',
      action: 'excel_read_selection',
      parameters: {},
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    // 步骤2: 写入表头
    const writeHeaderId = this.generateId();
    const startCell = spec.startCell || 'A1';
    const headers = spec.columns.map(c => c.name);
    
    steps.push({
      id: writeHeaderId,
      order: 2,
      phase: 'execution',
      description: '写入表头',
      action: 'excel_write_range',
      parameters: {
        address: startCell,
        values: [headers],
        sheet: spec.targetSheet,
      },
      dependsOn: [readStepId], // 使用步骤 ID，不是工具名！
      successCondition: { type: 'tool_success' },
      isWriteOperation: true,
      status: 'pending',
    });
    
    // 步骤3: 格式化表头
    const formatId = this.generateId();
    const headerRange = this.calculateHeaderRange(startCell, headers.length);
    
    steps.push({
      id: formatId,
      order: 3,
      phase: 'execution',
      description: '格式化表头',
      action: 'excel_format_range',
      parameters: {
        address: headerRange,
        bold: true,
        backgroundColor: '#4472C4',
        textColor: '#FFFFFF',
        horizontalAlignment: 'center',
      },
      dependsOn: [writeHeaderId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: true,
      status: 'pending',
    });
    
    // 步骤4: 自动列宽
    const autoFitId = this.generateId();
    steps.push({
      id: autoFitId,
      order: 4,
      phase: 'execution',
      description: '自动调整列宽',
      action: 'excel_auto_fit',
      parameters: {
        address: headerRange,
      },
      dependsOn: [formatId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: true,
      status: 'pending',
    });
    
    // 步骤5: 回复用户
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 5,
      phase: 'response',
      description: '通知用户完成',
      action: 'respond_to_user',
      parameters: {
        message: `已创建包含 ${headers.length} 列的表格: ${headers.join('、')}`,
      },
      dependsOn: [autoFitId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('创建表格', steps),
    };
  }
  
  /**
   * 编译写入数据
   */
  private compileWriteData(spec: WriteDataSpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    
    // 解析目标位置
    let targetAddress: string;
    let targetSheet: string | undefined;
    
    if (typeof spec.target === 'string') {
      targetAddress = spec.target;
    } else {
      targetAddress = spec.target.range || context.currentSelection || 'A1';
      targetSheet = spec.target.sheet;
    }
    
    // 步骤1: 写入数据
    const writeId = this.generateId();
    steps.push({
      id: writeId,
      order: 1,
      phase: 'execution',
      description: '写入数据',
      action: 'excel_write_range',
      parameters: {
        address: targetAddress,
        values: spec.data,
        sheet: targetSheet,
      },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: true,
      status: 'pending',
    });
    
    // 步骤2: 回复用户
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 2,
      phase: 'response',
      description: '通知用户完成',
      action: 'respond_to_user',
      parameters: {
        message: `数据已写入 ${targetAddress}`,
      },
      dependsOn: [writeId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('写入数据', steps),
    };
  }
  
  /**
   * 编译格式化
   */
  private compileFormat(spec: FormatSpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    
    const targetRange = spec.range || context.currentSelection || 'A1';
    
    // 步骤1: 应用格式
    const formatId = this.generateId();
    steps.push({
      id: formatId,
      order: 1,
      phase: 'execution',
      description: '应用格式',
      action: 'excel_format_range',
      parameters: {
        address: targetRange,
        ...spec.format,
      },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: true,
      status: 'pending',
    });
    
    // 步骤2: 回复用户
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 2,
      phase: 'response',
      description: '通知用户完成',
      action: 'respond_to_user',
      parameters: {
        message: `已格式化 ${targetRange}`,
      },
      dependsOn: [formatId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('格式化', steps),
    };
  }
  
  /**
   * 编译公式
   */
  private compileFormula(spec: FormulaSpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    
    // 步骤1: 感知数据范围
    const readId = this.generateId();
    steps.push({
      id: readId,
      order: 1,
      phase: 'sensing',
      description: '读取数据范围',
      action: 'excel_read_range',
      parameters: {
        address: spec.sourceRange || context.currentSelection || 'A1:A10',
      },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    // 步骤2: 设置公式
    const formulaId = this.generateId();
    let formula = spec.customFormula || this.generateFormula(spec);
    
    steps.push({
      id: formulaId,
      order: 2,
      phase: 'execution',
      description: '设置公式',
      action: 'excel_set_formula',
      parameters: {
        cell: spec.targetCell,
        formula,
      },
      dependsOn: [readId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: true,
      status: 'pending',
    });
    
    // 步骤3: 回复用户
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 3,
      phase: 'response',
      description: '通知用户完成',
      action: 'respond_to_user',
      parameters: {
        message: `已在 ${spec.targetCell} 设置公式: ${formula}`,
      },
      dependsOn: [formulaId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('设置公式', steps),
    };
  }
  
  /**
   * 编译图表
   */
  private compileChart(spec: ChartSpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    
    // 步骤1: 创建图表
    const chartId = this.generateId();
    steps.push({
      id: chartId,
      order: 1,
      phase: 'execution',
      description: '创建图表',
      action: 'excel_create_chart',
      parameters: {
        dataRange: spec.dataRange,
        chartType: this.mapChartType(spec.chartType),
        title: spec.title,
      },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: true,
      status: 'pending',
    });
    
    // 步骤2: 回复用户
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 2,
      phase: 'response',
      description: '通知用户完成',
      action: 'respond_to_user',
      parameters: {
        message: `已创建 ${spec.chartType} 图表`,
      },
      dependsOn: [chartId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('创建图表', steps),
    };
  }
  
  /**
   * 编译工作表操作
   */
  private compileSheet(spec: SheetSpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    
    if (spec.operation === 'create') {
      const createId = this.generateId();
      steps.push({
        id: createId,
        order: 1,
        phase: 'execution',
        description: '创建工作表',
        action: 'excel_create_sheet',
        parameters: {
          name: spec.sheetName,
        },
        dependsOn: [],
        successCondition: { type: 'tool_success' },
        isWriteOperation: true,
        status: 'pending',
      });
      
      const respondId = this.generateId();
      steps.push({
        id: respondId,
        order: 2,
        phase: 'response',
        description: '通知用户完成',
        action: 'respond_to_user',
        parameters: {
          message: `已创建工作表: ${spec.sheetName}`,
        },
        dependsOn: [createId],
        successCondition: { type: 'tool_success' },
        isWriteOperation: false,
        status: 'pending',
      });
    } else if (spec.operation === 'switch') {
      const switchId = this.generateId();
      steps.push({
        id: switchId,
        order: 1,
        phase: 'execution',
        description: '切换工作表',
        action: 'excel_switch_sheet',
        parameters: {
          name: spec.sheetName,
        },
        dependsOn: [],
        successCondition: { type: 'tool_success' },
        isWriteOperation: false,
        status: 'pending',
      });
      
      const respondId = this.generateId();
      steps.push({
        id: respondId,
        order: 2,
        phase: 'response',
        description: '通知用户完成',
        action: 'respond_to_user',
        parameters: {
          message: `已切换到工作表: ${spec.sheetName}`,
        },
        dependsOn: [switchId],
        successCondition: { type: 'tool_success' },
        isWriteOperation: false,
        status: 'pending',
      });
    }
    
    return {
      success: true,
      plan: this.buildExecutionPlan('工作表操作', steps),
    };
  }
  
  /**
   * 编译数据操作
   */
  private compileDataOperation(spec: DataOperationSpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    const targetRange = spec.range || context.currentSelection || 'A1:Z100';
    
    // 步骤1: 感知数据
    const readId = this.generateId();
    steps.push({
      id: readId,
      order: 1,
      phase: 'sensing',
      description: '读取数据范围',
      action: 'excel_read_range',
      parameters: { address: targetRange },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    // 步骤2: 执行操作
    const operationId = this.generateId();
    let action: string;
    let params: Record<string, unknown> = { address: targetRange };
    
    switch (spec.operation) {
      case 'sort':
        action = 'excel_sort_range';
        params.column = spec.sortColumn;
        params.ascending = spec.sortDirection !== 'desc';
        break;
      case 'filter':
        action = 'excel_filter';
        params.condition = spec.filterCondition;
        break;
      case 'dedupe':
        action = 'excel_remove_duplicates';
        break;
      default:
        action = 'excel_read_range'; // fallback
    }
    
    steps.push({
      id: operationId,
      order: 2,
      phase: 'execution',
      description: `执行${spec.operation}操作`,
      action,
      parameters: params,
      dependsOn: [readId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: spec.operation !== 'filter',
      status: 'pending',
    });
    
    // 步骤3: 回复用户
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 3,
      phase: 'response',
      description: '通知用户完成',
      action: 'respond_to_user',
      parameters: {
        message: `已完成${spec.operation}操作`,
      },
      dependsOn: [operationId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('数据操作', steps),
    };
  }
  
  /**
   * 编译查询
   */
  private compileQuery(spec: QuerySpec, context: CompileContext): CompileResult {
    const steps: PlanStep[] = [];
    
    // 容错: 如果 spec 为空或缺少必要字段，使用默认值
    const safeSpec = spec || {};
    const target = safeSpec.target || 'selection';
    const targetRange = safeSpec.range || context.currentSelection || 'A1:Z100';
    
    // 步骤1: 读取数据
    const readId = this.generateId();

    steps.push({
      id: readId,
      order: 1,
      phase: 'sensing',
      description: '读取数据',
      action: target === 'selection' ? 'excel_read_selection' : 'excel_read_range',
      parameters: target === 'selection' ? {} : { address: targetRange },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    // 步骤2: 分析并回复
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 2,
      phase: 'response',
      description: '分析数据并回复',
      action: 'respond_to_user',
      parameters: {
        message: '{{ANALYZE_AND_REPLY}}', // 特殊标记，执行器会处理
      },
      dependsOn: [readId],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('查询数据', steps),
    };
  }
  
  /**
   * 编译澄清
   */
  private compileClarify(spec: IntentSpec): CompileResult {
    const steps: PlanStep[] = [];
    
    const clarifyId = this.generateId();
    const clarifySpec = spec.spec as ClarifySpec;
    
    steps.push({
      id: clarifyId,
      order: 1,
      phase: 'response',
      description: '请求澄清',
      action: 'clarify_request',
      parameters: {
        question: spec.clarificationQuestion || clarifySpec?.question || '请提供更多信息',
        options: spec.clarificationOptions || clarifySpec?.options,
      },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('请求澄清', steps),
    };
  }
  
  /**
   * 编译回复
   */
  private compileRespond(spec: RespondSpec): CompileResult {
    const steps: PlanStep[] = [];
    
    const respondId = this.generateId();
    steps.push({
      id: respondId,
      order: 1,
      phase: 'response',
      description: '回复用户',
      action: 'respond_to_user',
      parameters: {
        message: spec.message,
      },
      dependsOn: [],
      successCondition: { type: 'tool_success' },
      isWriteOperation: false,
      status: 'pending',
    });
    
    return {
      success: true,
      plan: this.buildExecutionPlan('回复用户', steps),
    };
  }
  
  // ========== 辅助方法 ==========
  
  /**
   * 生成唯一 ID
   */
  private generateId(): string {
    return `step_${Date.now()}_${++this.idCounter}`;
  }
  
  /**
   * 计算表头范围
   */
  private calculateHeaderRange(startCell: string, columnCount: number): string {
    const match = startCell.match(/([A-Z]+)(\d+)/);
    if (!match) return startCell;
    
    const startCol = match[1];
    const row = match[2];
    const endCol = this.getColumnLetter(this.getColumnNumber(startCol) + columnCount - 1);
    
    return `${startCell}:${endCol}${row}`;
  }
  
  /**
   * 获取列号
   */
  private getColumnNumber(col: string): number {
    let num = 0;
    for (let i = 0; i < col.length; i++) {
      num = num * 26 + (col.charCodeAt(i) - 64);
    }
    return num;
  }
  
  /**
   * 获取列字母
   */
  private getColumnLetter(num: number): string {
    let letter = '';
    while (num > 0) {
      const mod = (num - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      num = Math.floor((num - 1) / 26);
    }
    return letter;
  }
  
  /**
   * 生成公式
   */
  private generateFormula(spec: FormulaSpec): string {
    const range = spec.sourceRange || 'A1:A10';
    
    switch (spec.formulaType) {
      case 'sum':
        return `=SUM(${range})`;
      case 'average':
        return `=AVERAGE(${range})`;
      case 'count':
        return `=COUNT(${range})`;
      case 'if':
        return `=IF(${spec.condition || 'A1>0'}, "是", "否")`;
      default:
        return spec.customFormula || `=SUM(${range})`;
    }
  }
  
  /**
   * 映射图表类型
   */
  private mapChartType(type: string): string {
    const mapping: Record<string, string> = {
      line: 'Line',
      bar: 'BarClustered',
      pie: 'Pie',
      column: 'ColumnClustered',
      area: 'Area',
      scatter: 'XYScatter',
    };
    return mapping[type] || 'ColumnClustered';
  }
  
  /**
   * 构建执行计划
   */
  private buildExecutionPlan(description: string, steps: PlanStep[]): ExecutionPlan {
    return {
      id: this.generateId(),
      taskDescription: description,
      taskType: 'mixed',
      steps,
      taskSuccessConditions: [
        {
          id: this.generateId(),
          description: '所有步骤完成',
          type: 'all_steps_complete',
          priority: 1,
        },
      ],
      completionMessage: `${description}完成`,
      dependencyCheck: {
        passed: true,
        missingDependencies: [],
        circularDependencies: [],
        warnings: [],
        semanticDependencies: [],
        unresolvedSemanticDeps: [],
      },
      estimatedDuration: steps.length * 1000,
      estimatedSteps: steps.length,
      risks: [],
      phase: 'planning',
      currentStep: 0,
      completedSteps: 0,
      failedSteps: 0,
      fieldStepIdMap: {},
    };
  }
}

// ========== 单例导出 ==========

export const specCompiler = new SpecCompiler();

export default SpecCompiler;
