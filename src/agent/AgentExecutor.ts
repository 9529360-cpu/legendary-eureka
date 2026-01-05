/**
 * AgentExecutor - 执行引擎 v4.0
 * 
 * Layer 3: 整合 IntentParser + SpecCompiler + ToolRegistry
 * 
 * 执行流程:
 * 1. IntentParser: 用户消息 → IntentSpec (调用 LLM，不含工具名)
 * 2. SpecCompiler: IntentSpec → ExecutionPlan (纯规则，零 Token)
 * 3. AgentExecutor: ExecutionPlan → 执行结果 (调用工具)
 * 
 * @module agent/AgentExecutor
 */

import { IntentParser, ParseContext } from './IntentParser';
import { SpecCompiler, CompileContext, CompileResult } from './SpecCompiler';
import { ToolRegistry } from './registry';
import { ExecutionPlan, PlanStep } from './TaskPlanner';
import { Tool, ToolResult } from './types/tool';
import { IntentSpec } from './types/intent';
import createExcelTools from './ExcelAdapter';

// ========== 执行结果 ==========

/**
 * v4.0 执行结果 (AgentExecutor 专用)
 */
export interface AgentExecutionResult {
  /** 是否成功 */
  success: boolean;
  
  /** 回复消息 */
  message: string;
  
  /** 执行的步骤 */
  executedSteps: StepResult[];
  
  /** 错误信息 */
  error?: string;
  
  /** 是否需要用户确认 */
  needsConfirmation?: boolean;
  
  /** 确认问题 */
  confirmationQuestion?: string;
}

// 向后兼容别名
export type ExecutionResult = AgentExecutionResult;

/**
 * 步骤执行结果
 */
export interface StepResult {
  stepId: string;
  action: string;
  success: boolean;
  output?: string;
  error?: string;
  duration: number;
}

// ========== 事件类型 ==========

export type ExecutorEventType = 
  | 'intent:parsed'
  | 'plan:compiled'
  | 'step:start'
  | 'step:complete'
  | 'step:error'
  | 'execution:complete';

export interface ExecutorEvent {
  type: ExecutorEventType;
  data: unknown;
  timestamp: Date;
}

// ========== AgentExecutor 类 ==========

/**
 * 执行引擎 - v4.0 架构核心
 */
export class AgentExecutor {
  private intentParser: IntentParser;
  private specCompiler: SpecCompiler;
  private toolRegistry: ToolRegistry;
  private eventHandlers: Map<ExecutorEventType, Array<(event: ExecutorEvent) => void>>;
  
  constructor(toolRegistry: ToolRegistry) {
    this.intentParser = new IntentParser();
    this.specCompiler = new SpecCompiler();
    this.toolRegistry = toolRegistry;
    this.eventHandlers = new Map();
  }
  
  /**
   * 执行用户请求
   * 
   * 这是新架构的主入口：
   * 1. 意图解析 (LLM)
   * 2. 规格编译 (规则)
   * 3. 执行计划 (工具)
   */
  async execute(context: ParseContext): Promise<ExecutionResult> {
    const startTime = Date.now();
    const stepResults: StepResult[] = [];
    
    try {
      console.log('[AgentExecutor] === 开始执行 ===');
      console.log('[AgentExecutor] 用户消息:', context.userMessage);
      
      // ===== Layer 1: 意图解析 =====
      console.log('[AgentExecutor] Step 1: 解析意图...');
      const intentSpec = await this.intentParser.parse(context);
      
      this.emit('intent:parsed', {
        intent: intentSpec,  // 发送完整的 IntentSpec 对象
        confidence: intentSpec.confidence,
        needsClarification: intentSpec.needsClarification,
      });
      
      console.log('[AgentExecutor] 意图:', intentSpec.intent, '置信度:', intentSpec.confidence);
      
      // 如果需要澄清，直接返回
      if (intentSpec.needsClarification) {
        console.log('[AgentExecutor] 需要澄清，返回问题');
        return {
          success: true,
          message: intentSpec.clarificationQuestion || '请提供更多信息',
          executedSteps: [],
          needsConfirmation: true,
          confirmationQuestion: intentSpec.clarificationQuestion,
        };
      }
      
      // ===== Layer 2: 规格编译 =====
      console.log('[AgentExecutor] Step 2: 编译规格...');
      const compileContext: CompileContext = {
        currentSelection: context.selection?.address,
        activeSheet: context.activeSheet,
      };
      
      const compileResult = this.specCompiler.compile(intentSpec, compileContext);
      
      if (!compileResult.success || !compileResult.plan) {
        console.error('[AgentExecutor] 编译失败:', compileResult.error);
        return {
          success: false,
          message: compileResult.error || '无法生成执行计划',
          executedSteps: [],
          error: compileResult.error,
        };
      }
      
      this.emit('plan:compiled', {
        stepCount: compileResult.plan.steps.length,
        description: compileResult.plan.taskDescription,
        plan: compileResult.plan,
      });
      
      console.log('[AgentExecutor] 编译完成，步骤数:', compileResult.plan.steps.length);
      
      // ===== Layer 3: 执行计划 =====
      console.log('[AgentExecutor] Step 3: 执行计划...');
      const plan = compileResult.plan;
      const totalSteps = plan.steps.length;
      let lastOutput = '';
      
      for (let i = 0; i < plan.steps.length; i++) {
        const step = plan.steps[i];
        const stepStart = Date.now();
        
        this.emit('step:start', {
          step: { description: step.description || step.action, id: step.id, action: step.action },
          index: i,
          total: totalSteps,
          stepId: step.id,
          action: step.action,
          description: step.description || step.action,
        });
        
        console.log(`[AgentExecutor] 执行步骤: ${step.action}`);
        
        try {
          const result = await this.executeStep(step, lastOutput);
          
          stepResults.push({
            stepId: step.id,
            action: step.action,
            success: result.success,
            output: result.output,
            error: result.error,
            duration: Date.now() - stepStart,
          });
          
          if (result.success) {
            lastOutput = result.output;
            this.emit('step:complete', {
              step: { description: step.description || step.action, id: step.id, action: step.action },
              result: { success: true, output: result.output },
              index: i,
              total: totalSteps,
              stepId: step.id,
              success: true,
              output: result.output,
            });
          } else {
            this.emit('step:error', {
              step: { description: step.description || step.action, id: step.id, action: step.action },
              index: i,
              total: totalSteps,
              stepId: step.id,
              error: result.error,
            });
            
            // 如果步骤失败，根据策略决定是否继续
            if (step.isWriteOperation) {
              console.error(`[AgentExecutor] 写操作失败，停止执行: ${result.error}`);
              return {
                success: false,
                message: `操作失败: ${result.error}`,
                executedSteps: stepResults,
                error: result.error,
              };
            }
          }
          
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : String(error);
          console.error(`[AgentExecutor] 步骤执行异常:`, error);
          
          stepResults.push({
            stepId: step.id,
            action: step.action,
            success: false,
            error: errorMsg,
            duration: Date.now() - stepStart,
          });
          
          this.emit('step:error', {
            stepId: step.id,
            error: errorMsg,
          });
        }
      }
      
      // ===== 生成回复 =====
      const message = this.generateResponseMessage(stepResults, plan);
      
      this.emit('execution:complete', {
        success: true,
        duration: Date.now() - startTime,
        stepCount: stepResults.length,
      });
      
      console.log('[AgentExecutor] === 执行完成 ===');
      
      return {
        success: true,
        message,
        executedSteps: stepResults,
      };
      
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      console.error('[AgentExecutor] 执行失败:', error);
      
      this.emit('execution:complete', {
        success: false,
        error: errorMsg,
        duration: Date.now() - startTime,
      });
      
      return {
        success: false,
        message: `执行失败: ${errorMsg}`,
        executedSteps: stepResults,
        error: errorMsg,
      };
    }
  }
  
  /**
   * 执行单个步骤
   */
  private async executeStep(step: PlanStep, previousOutput: string): Promise<ToolResult> {
    const toolName = step.action;
    
    // 特殊处理 respond_to_user
    if (toolName === 'respond_to_user') {
      const message = step.parameters?.message as string || '';
      
      // 如果是分析回复，需要基于之前的数据生成
      if (message === '{{ANALYZE_AND_REPLY}}') {
        return {
          success: true,
          output: previousOutput || '数据已读取',
        };
      }
      
      return {
        success: true,
        output: message,
      };
    }
    
    // 特殊处理 clarify_request
    if (toolName === 'clarify_request') {
      const question = step.parameters?.question as string || '请提供更多信息';
      return {
        success: true,
        output: question,
      };
    }
    
    // 查找工具
    const tool = this.toolRegistry.get(toolName);
    if (!tool) {
      console.warn(`[AgentExecutor] 工具不存在: ${toolName}`);
      return {
        success: false,
        output: '',
        error: `工具不存在: ${toolName}`,
      };
    }
    
    // 执行工具
    try {
      const result = await tool.execute(step.parameters || {});
      return result;
    } catch (error) {
      return {
        success: false,
        output: '',
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }
  
  /**
   * 生成回复消息
   */
  private generateResponseMessage(stepResults: StepResult[], plan: ExecutionPlan): string {
    // 查找最后一个 respond_to_user 步骤的输出
    for (let i = stepResults.length - 1; i >= 0; i--) {
      const result = stepResults[i];
      if (result.action === 'respond_to_user' && result.success && result.output) {
        return result.output;
      }
      if (result.action === 'clarify_request' && result.success && result.output) {
        return result.output;
      }
    }
    
    // 如果没有显式回复，生成默认消息
    const successCount = stepResults.filter(r => r.success).length;
    const totalCount = stepResults.length;
    
    if (successCount === totalCount) {
      return plan.completionMessage || '操作已完成';
    } else {
      return `完成 ${successCount}/${totalCount} 个步骤`;
    }
  }
  
  // ========== 事件系统 ==========
  
  /**
   * 订阅事件
   */
  on(eventType: ExecutorEventType, handler: (event: ExecutorEvent) => void): void {
    if (!this.eventHandlers.has(eventType)) {
      this.eventHandlers.set(eventType, []);
    }
    this.eventHandlers.get(eventType)!.push(handler);
  }
  
  /**
   * 发送事件
   */
  private emit(type: ExecutorEventType, data: unknown): void {
    const event: ExecutorEvent = {
      type,
      data,
      timestamp: new Date(),
    };
    
    const handlers = this.eventHandlers.get(type);
    if (handlers) {
      for (const handler of handlers) {
        try {
          handler(event);
        } catch (error) {
          console.error('[AgentExecutor] 事件处理错误:', error);
        }
      }
    }
  }
}

// ========== 工厂函数 ==========

/**
 * 创建默认的 ToolRegistry（包含 Excel 工具）
 */
function createDefaultToolRegistry(): ToolRegistry {
  const registry = new ToolRegistry();
  
  // 注册 Excel 工具（61 个）
  const excelTools = createExcelTools();
  registry.registerMany(excelTools);
  
  console.log(`[AgentExecutor] 已注册 ${excelTools.length} 个 Excel 工具`);
  return registry;
}

/**
 * 创建 AgentExecutor 实例
 * 
 * @param toolRegistry 工具注册中心（可选，默认自动注册 Excel 工具）
 */
export function createAgentExecutor(toolRegistry?: ToolRegistry): AgentExecutor {
  const registry = toolRegistry ?? createDefaultToolRegistry();
  return new AgentExecutor(registry);
}

export default AgentExecutor;
