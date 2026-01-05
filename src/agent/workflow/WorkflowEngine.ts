/**
 * 工作流系统实现
 *
 * 借鉴 LlamaIndex Workflows 设计，提供事件驱动的工作流编排能力
 *
 * @packageDocumentation
 */

import type {
  WorkflowEvent,
  WorkflowState,
  WorkflowEventHandler,
  SimpleWorkflow,
  WorkflowContextInterface,
  WorkflowEventStreamInterface,
  WorkflowEventRegistryInterface,
  ToolCallInfo,
  ToolResult,
} from "../types";
import type { ExecutionPlan, PlanStep } from "../types/workflow";

// ========== 工厂函数 ==========

/**
 * 创建类型化工作流事件的工厂函数
 *
 * 用法: const myEvent = createWorkflowEvent<MyDataType>("my_event");
 *       workflow.emit(myEvent.with({ ... }));
 */
export function createWorkflowEvent<T>(eventType: string) {
  return {
    type: eventType,
    with: (data: T): WorkflowEvent<T> => ({
      type: eventType,
      data,
      timestamp: new Date(),
    }),
    is: (event: WorkflowEvent<unknown>): event is WorkflowEvent<T> => event.type === eventType,
  };
}

/**
 * 创建初始工作流状态
 */
export function createInitialWorkflowState(): WorkflowState {
  return {
    currentTaskId: null,
    currentRequest: null,
    isRunning: false,
    isPaused: false,
    collectedData: [],
    toolCallHistory: [],
    stepCounter: 0,
    maxSteps: 30,
    awaitingConfirmation: false,
    awaitingFollowUp: false,
    // v2.9.60: 补充缺失的属性
    toolsCalled: [],
    currentResponse: "",
    errors: [],
    custom: {},
  };
}

// ========== 预定义工作流事件 ==========

/**
 * 预定义的工作流事件集合
 */
export const WorkflowEvents = {
  // 任务生命周期事件
  taskStart: createWorkflowEvent<{ taskId: string; request: string }>("task:start"),
  taskComplete: createWorkflowEvent<{ taskId: string; result: string }>("task:complete"),
  taskError: createWorkflowEvent<{ taskId: string; error: string }>("task:error"),
  taskPending: createWorkflowEvent<{ taskId: string; reason: string }>("task:pending"),

  // Agent 流式事件
  agentStream: createWorkflowEvent<{
    delta: string;
    response: string;
    currentAgentName: string;
    toolCalls: ToolCallInfo[];
    thinkingDelta?: string;
  }>("agent:stream"),
  agentOutput: createWorkflowEvent<{
    response: string;
    structuredResponse?: Record<string, unknown>;
    currentAgentName: string;
    toolCalls: ToolCallInfo[];
    retryMessages: string[];
  }>("agent:output"),

  // 工具调用事件
  toolCall: createWorkflowEvent<ToolCallInfo>("tool:call"),
  toolCallResult: createWorkflowEvent<{
    toolName: string;
    toolKwargs: Record<string, unknown>;
    toolId: string;
    toolOutput: ToolResult;
    returnDirect: boolean;
  }>("tool:result"),

  // 计划事件
  planGenerated: createWorkflowEvent<{ plan: ExecutionPlan }>("plan:generated"),
  planStepStart: createWorkflowEvent<{ stepIndex: number; step: PlanStep }>("plan:step_start"),
  planStepComplete: createWorkflowEvent<{ stepIndex: number; result: ToolResult }>(
    "plan:step_complete"
  ),
  planConfirmationRequired: createWorkflowEvent<{ taskId: string; preview: string }>(
    "plan:confirmation_required"
  ),

  // 进度事件
  progressUpdate: createWorkflowEvent<{ current: number; total: number; message: string }>(
    "progress:update"
  ),

  // 验证事件
  validationStart: createWorkflowEvent<{ type: string }>("validation:start"),
  validationComplete: createWorkflowEvent<{ passed: boolean; errors: string[] }>(
    "validation:complete"
  ),

  // 思考链事件 (CoT)
  cotStepStart: createWorkflowEvent<{ step: string }>("cot:step_start"),
  cotStepComplete: createWorkflowEvent<{ step: string; result: string }>("cot:step_complete"),

  // 跟进上下文事件
  followUpContextSet: createWorkflowEvent<{
    originalRequest: string;
    suggestedAction: string;
  }>("followup:context_set"),
  followUpHandled: createWorkflowEvent<{ userReply: string; action: string }>("followup:handled"),
};

// ========== 工作流上下文实现 ==========

/**
 * 工作流上下文 (借鉴 LlamaIndex Workflows 的 Context)
 *
 * 提供事件发送、状态访问和取消控制能力
 */
export class WorkflowContext implements WorkflowContextInterface {
  private state: WorkflowState;
  private eventQueue: Array<WorkflowEvent<unknown>> = [];
  private abortController: AbortController;
  private _isCancelled = false;

  constructor(initialState?: Partial<WorkflowState>) {
    this.state = { ...createInitialWorkflowState(), ...initialState };
    this.abortController = new AbortController();
  }

  /**
   * 发送事件到工作流 (LlamaIndex: context.sendEvent())
   */
  sendEvent<T>(event: WorkflowEvent<T> | { type: string; payload?: T; data?: T }): void {
    let normalizedEvent: WorkflowEvent<T>;

    if ("data" in event && event.data !== undefined) {
      normalizedEvent = event as WorkflowEvent<T>;
    } else if ("payload" in event && event.payload !== undefined) {
      normalizedEvent = {
        type: event.type,
        data: event.payload,
        timestamp: new Date(),
      };
    } else {
      normalizedEvent = {
        type: event.type,
        data: undefined as T,
        timestamp: new Date(),
      };
    }

    this.eventQueue.push(normalizedEvent as WorkflowEvent<unknown>);
  }

  getEvents(): Array<WorkflowEvent<unknown>> {
    return [...this.eventQueue];
  }

  consumeEvent(): WorkflowEvent<unknown> | undefined {
    return this.eventQueue.shift();
  }

  clearEvents(): void {
    this.eventQueue = [];
  }

  getState(): WorkflowState {
    return this.state;
  }

  updateState(updates: Partial<WorkflowState>): void {
    this.state = { ...this.state, ...updates };
  }

  setCustom<T>(key: string, value: T): void {
    this.state.custom[key] = value;
  }

  getCustom<T>(key: string): T | undefined {
    return this.state.custom[key] as T | undefined;
  }

  cancel(): void {
    this._isCancelled = true;
    this.abortController.abort();
  }

  get isCancelled(): boolean {
    return this._isCancelled;
  }

  get signal(): AbortSignal {
    return this.abortController.signal;
  }
}

// ========== 工作流事件注册表实现 ==========

/**
 * 工作流事件处理器注册表
 */
export class WorkflowEventRegistry implements WorkflowEventRegistryInterface {
  private handlers = new Map<string, WorkflowEventHandler[]>();

  handle<T>(event: WorkflowEvent<T>, handler: WorkflowEventHandler<T>): this {
    const eventType = event.type || (event as unknown as { name: string }).name || String(event);
    if (!this.handlers.has(eventType)) {
      this.handlers.set(eventType, []);
    }
    this.handlers.get(eventType)!.push(handler as WorkflowEventHandler);
    return this;
  }

  async dispatch<T>(
    eventType: string,
    context: WorkflowContextInterface,
    data: T
  ): Promise<Array<{ type: string; payload: unknown } | void>> {
    const eventHandlers = this.handlers.get(eventType) || [];
    const results: Array<{ type: string; payload: unknown } | void> = [];

    for (const handler of eventHandlers) {
      if (context.isCancelled) break;
      const result = await handler(context, data);
      results.push(result);
    }

    return results;
  }

  getRegisteredEvents(): string[] {
    return Array.from(this.handlers.keys());
  }
}

// ========== 工作流事件流实现 ==========

/**
 * 工作流事件流 (借鉴 LlamaIndex stream.until().toArray())
 */
export class WorkflowEventStream implements WorkflowEventStreamInterface {
  private events: Array<{ type: string; payload: unknown }> = [];
  private stopCondition: ((event: { type: string; payload: unknown }) => boolean) | null = null;
  private isStopped = false;

  push(event: WorkflowEvent<unknown> | { type: string; payload: unknown }): void {
    if (this.isStopped) return;

    const normalized: { type: string; payload: unknown } =
      "data" in event ? { type: event.type, payload: event.data } : event;

    this.events.push(normalized);

    if (this.stopCondition && this.stopCondition(normalized)) {
      this.isStopped = true;
    }
  }

  until<T>(stopEvent: WorkflowEvent<T>): this {
    this.stopCondition = (event) => event.type === stopEvent.type;
    return this;
  }

  toArray(): Array<{ type: string; payload: unknown }> {
    return [...this.events];
  }

  last(): { type: string; payload: unknown } | undefined {
    return this.events[this.events.length - 1];
  }

  filter<T>(eventType: WorkflowEvent<T>): Array<{ type: string; payload: T }> {
    return this.events.filter((e) => e.type === eventType.type) as Array<{
      type: string;
      payload: T;
    }>;
  }

  *[Symbol.iterator](): IterableIterator<{ type: string; payload: unknown }> {
    for (const event of this.events) {
      yield event;
    }
  }
}

// ========== 工作流执行器 ==========

/**
 * 创建工作流执行器 (简化版 LlamaIndex createWorkflow)
 */
export function createSimpleWorkflow(): SimpleWorkflow {
  const registry = new WorkflowEventRegistry();
  const stream = new WorkflowEventStream();

  const workflow: SimpleWorkflow = {
    on<T>(eventType: string | WorkflowEvent<T>, handler: WorkflowEventHandler<T>): SimpleWorkflow {
      const type = typeof eventType === "string" ? eventType : eventType.type;
      const wrappedEvent = { type } as WorkflowEvent<T>;
      registry.handle(wrappedEvent, handler);
      return workflow;
    },

    async run(context: WorkflowContextInterface): Promise<WorkflowEventStreamInterface> {
      while (!context.isCancelled) {
        const event = context.consumeEvent();
        if (!event) break;

        stream.push(event);
        const results = await registry.dispatch(event.type, context, event.data);

        for (const result of results) {
          if (result && "type" in result) {
            context.sendEvent(result);
            stream.push(result);
          }
        }
      }
      return stream;
    },

    getStream(): WorkflowEventStreamInterface {
      return stream;
    },

    getRegistry(): WorkflowEventRegistryInterface {
      return registry;
    },
  };

  return workflow;
}
