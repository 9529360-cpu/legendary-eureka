# Activepieces 学习笔记 - Agent 架构参考

> 来源: activepieces-main 项目
> 日期: 2026-01-04
> 用途: 提取有价值的设计模式供 Excel Agent 项目参考

---

## 1. 核心架构亮点

### 1.1 Flow Executor 执行引擎

**路径**: `packages/engine/src/lib/handler/flow-executor.ts`

```typescript
// 关键设计: 类型到执行器的映射
function getExecuteFunction(): Record<FlowActionType, BaseExecutor<FlowAction>> {
    return {
        [FlowActionType.CODE]: codeExecutor,
        [FlowActionType.LOOP_ON_ITEMS]: loopExecutor,
        [FlowActionType.PIECE]: pieceExecutor,
        [FlowActionType.ROUTER]: routerExecuter,
    }
}

// 执行循环 - 类似我们的 ReAct 循环
async execute({ action, constants, executionState }) {
    let currentAction = action
    while (!isNil(currentAction)) {
        if (currentAction.skip && !testSingleStepMode) {
            currentAction = currentAction.nextAction
            continue
        }
        
        const handler = this.getExecutorForAction(currentAction.type)
        
        // 发送进度更新
        progressService.sendUpdate({...})
        
        // 执行并获取新状态
        flowExecutionContext = await handler.handle({...})
        
        // 检查是否应该中断
        const shouldBreakExecution = flowExecutionContext.verdict.status !== FlowRunStatus.RUNNING
        if (shouldBreakExecution) break
        
        currentAction = currentAction.nextAction
    }
    return flowExecutionContext.setDuration(flowEndTime - flowStartTime)
}
```

**学习点**:
- ✅ 状态机设计: `FlowRunStatus.RUNNING | PAUSED | SUCCEEDED | FAILED`
- ✅ 不可变状态更新: `executionState = await handler.handle({...})`
- ✅ 进度服务分离: `progressService.sendUpdate()`
- ✅ 执行时间追踪: `setDuration()`

---

### 1.2 执行上下文 (Immutable State)

**路径**: `packages/engine/src/lib/handler/context/flow-execution-context.ts`

```typescript
export class FlowExecutorContext {
    tags: readonly string[]
    steps: Readonly<Record<string, StepOutput>>
    verdict: FlowVerdict
    currentPath: StepExecutionPath
    duration: number

    // 不可变更新模式
    public upsertStep(stepName: string, stepOutput: StepOutput): FlowExecutorContext {
        const steps = { ...this.steps }
        const targetMap = getStateAtPath({ currentPath: this.currentPath, steps })
        targetMap[stepName] = stepOutput
        
        return new FlowExecutorContext({
            ...this,
            steps,
        })
    }

    public setVerdict(verdict: FlowVerdict): FlowExecutorContext {
        return new FlowExecutorContext({
            ...this,
            verdict,
        })
    }
}
```

**学习点**:
- ✅ 使用 `Readonly<>` 强制不可变
- ✅ 每次状态变更返回新实例
- ✅ `verdict` 判决机制 (RUNNING/PAUSED/FAILED/SUCCEEDED)

---

### 1.3 错误处理与重试

**路径**: `packages/engine/src/lib/helper/error-handling.ts`

```typescript
export async function runWithExponentialBackoff<T>(
    executionState: FlowExecutorContext,
    action: T,
    constants: EngineConstants,
    requestFunction: RequestFunction<T>,
    attemptCount = 1,
): Promise<FlowExecutorContext> {
    const resultExecutionState = await requestFunction({ action, executionState, constants })
    const retryEnabled = action.settings.errorHandlingOptions?.retryOnFailure?.value
    
    if (
        executionFailedWithRetryableError(resultExecutionState) &&
        attemptCount < constants.retryConstants.maxAttempts &&
        retryEnabled
    ) {
        const backoffTime = Math.pow(
            constants.retryConstants.retryExponential, 
            attemptCount
        ) * constants.retryConstants.retryInterval
        
        await new Promise(resolve => setTimeout(resolve, backoffTime))
        return runWithExponentialBackoff(executionState, action, constants, requestFunction, attemptCount + 1)
    }

    return resultExecutionState
}

export async function continueIfFailureHandler(
    executionState: FlowExecutorContext,
    action: CodeAction | PieceAction,
): Promise<FlowExecutorContext> {
    const continueOnFailure = action.settings.errorHandlingOptions?.continueOnFailure?.value

    if (executionState.verdict.status === FlowRunStatus.FAILED && continueOnFailure) {
        return executionState.setVerdict({ status: FlowRunStatus.RUNNING })
    }
    return executionState
}
```

**学习点**:
- ✅ 指数退避重试
- ✅ `continueOnFailure` 选项让流程继续
- ✅ 可配置的错误处理策略

---

### 1.4 MCP Server 集成

**路径**: `packages/server/api/src/app/mcp/mcp-service.ts`

```typescript
export const mcpServerService = (log: FastifyBaseLogger) => {
    return {
        buildServer: async ({ mcp }): Promise<McpServer> => {
            const server = new McpServer({
                name: 'Activepieces',
                version: '1.0.0',
            })
            
            const enabledFlows = mcp.flows.filter((flow) => flow.status === FlowStatus.ENABLED)
            
            for (const flow of enabledFlows) {
                const mcpTrigger = flow.version.trigger.settings as McpTrigger
                const mcpInputs = mcpTrigger.input?.inputSchema ?? []
                
                // 将 McpProperty 转换为 Zod Schema
                const zodFromInputSchema = Object.fromEntries(
                    mcpInputs.map((property) => [property.name, mcpPropertyToZod(property)])
                )
                
                const toolName = (mcpTrigger.input?.toolName ?? flow.version.displayName)
                    .toLowerCase()
                    .replace(/[^a-zA-Z0-9_]/g, '_') + '_' + flow.id.substring(0, 4)
                
                // 注册工具
                server.tool(toolName, toolDescription, zodFromInputSchema, { title: toolName }, async (args) => {
                    const response = await webhookService.handleWebhook({...})
                    
                    if (isOkay) {
                        return {
                            content: [{
                                type: 'text',
                                text: `✅ Successfully executed flow ${flow.version.displayName}\n\n` +
                                    `Output:\n\`\`\`json\n${JSON.stringify(response, null, 2)}\n\`\`\``,
                            }],
                        }
                    }
                    return { content: [{ type: 'text', text: `❌ Error...` }] }
                })
            }
            return server
        },
    }
}

function mcpPropertyToZod(property: McpProperty): z.ZodTypeAny {
    let schema: z.ZodTypeAny
    switch (property.type) {
        case McpPropertyType.TEXT:
        case McpPropertyType.DATE:
            schema = z.string()
            break
        case McpPropertyType.NUMBER:
            schema = z.number()
            break
        case McpPropertyType.BOOLEAN:
            schema = z.boolean()
            break
        case McpPropertyType.ARRAY:
            schema = z.array(z.string())
            break
        case McpPropertyType.OBJECT:
            schema = z.record(z.string(), z.string())
            break
        default:
            schema = z.unknown()
    }
    if (property.description) {
        schema = schema.describe(property.description)
    }
    return property.required ? schema : schema.nullish()
}
```

**学习点**:
- ✅ 动态工具注册模式
- ✅ 属性类型到 Zod Schema 的映射
- ✅ 统一的工具输出格式

---

### 1.5 Agent Tools 动态解析

**路径**: `packages/engine/src/lib/tools/index.ts`

```typescript
export const agentTools = {
    async tools({ engineConstants, tools, model }: ConstructToolParams): Promise<ToolSet> {
        const piecesTools = await Promise.all(tools.map(async (tool) => {
            const { pieceAction } = await pieceLoader.getPieceAndActionOrThrow({...})
            
            return {
                name: tool.toolName,
                description: pieceAction.description,
                inputSchema: z.object({
                    instruction: z.string().describe('The instruction to the tool'),
                }),
                execute: async ({ instruction }) => execute({
                    ...engineConstants,
                    instruction,
                    pieceName: tool.pieceMetadata.pieceName,
                    pieceVersion: tool.pieceMetadata.pieceVersion,
                    actionName: tool.pieceMetadata.actionName,
                    predefinedInput: tool.pieceMetadata.predefinedInput,
                    model,
                }),
            }
        }))
        
        return Object.fromEntries(piecesTools.map((tool) => [tool.name, tool]))
    },
}

// LLM 驱动的属性解析
async function resolveProperties(
    depthToPropertyMap: Record<number, string[]>, 
    instruction: string, 
    action: Action, 
    model: LanguageModel, 
    operation: ExecuteToolOperation,
): Promise<Record<string, unknown>> {
    let result: Record<string, unknown> = { ...operation.predefinedInput }
    
    for (const [_, properties] of Object.entries(depthToPropertyMap)) {
        // 构建 Zod schema
        const propertyToFill: Record<string, z.ZodTypeAny> = {}
        const propertyPrompts: string[] = []
        
        for (const property of properties) {
            const propertySchema = await propertyToSchema(property, ...)
            propertyToFill[property] = propertySchema
        }
        
        // 使用 LLM 填充属性
        const { object } = await generateObject({
            model,
            schema: z.object(propertyToFill),
            prompt: constructExtractionPrompt(instruction, propertyToFill, propertyPrompts, result),
            mode: 'json',
            output: 'object',
        })
        
        result = { ...result, ...(object as Record<string, unknown>) }
    }
    return result
}
```

**学习点**:
- ✅ 使用 LLM 动态解析工具参数 (`generateObject`)
- ✅ 拓扑排序处理属性依赖 (`tsort`)
- ✅ 已填充值作为上下文传递

---

### 1.6 进度服务 (Debounce + 锁)

**路径**: `packages/engine/src/lib/services/progress.service.ts`

```typescript
let lastScheduledUpdateId: NodeJS.Timeout | null = null
let lastActionExecutionTime: number | undefined = undefined
const MAXIMUM_UPDATE_THRESHOLD = 15000
const DEBOUNCE_THRESHOLD = 5000
const lock = new Mutex()
const updateLock = new Mutex()

export const progressService = {
    sendUpdate: async (params: UpdateStepProgressParams): Promise<void> => {
        return updateLock.runExclusive(async () => {
            if (lastScheduledUpdateId) {
                clearTimeout(lastScheduledUpdateId)
            }

            const shouldUpdateNow = isNil(lastActionExecutionTime) || 
                (Date.now() - lastActionExecutionTime > MAXIMUM_UPDATE_THRESHOLD)
            
            if (shouldUpdateNow || params.updateImmediate) {
                await sendUpdateRunRequest(params)
                return
            }

            // 防抖延迟发送
            lastScheduledUpdateId = setTimeout(async () => {
                await sendUpdateRunRequest(params)
            }, DEBOUNCE_THRESHOLD)
        })
    },
}
```

**学习点**:
- ✅ 使用 Mutex 防止并发问题
- ✅ 防抖 + 最大阈值策略
- ✅ 优雅关闭信号处理 (`SIGTERM`, `SIGINT`)

---

## 2. 可复用到 Excel Agent 的模式

### 2.1 执行状态管理

```typescript
// 借鉴 FlowExecutorContext 的不可变模式
export class AgentExecutionContext {
    readonly steps: Record<string, StepOutput>
    readonly verdict: AgentVerdict // running | paused | succeeded | failed | awaiting_approval
    readonly duration: number
    readonly approvalState?: ApprovalState
    
    upsertStep(name: string, output: StepOutput): AgentExecutionContext {
        return new AgentExecutionContext({
            ...this,
            steps: { ...this.steps, [name]: output },
        })
    }
    
    setVerdict(verdict: AgentVerdict): AgentExecutionContext {
        return new AgentExecutionContext({ ...this, verdict })
    }
}
```

### 2.2 错误处理策略

```typescript
// 借鉴 errorHandlingOptions 模式
interface ToolErrorHandling {
    retryOnFailure: boolean
    maxRetries: number
    continueOnFailure: boolean
    exponentialBackoff: boolean
}
```

### 2.3 MCP 工具类型映射

```typescript
// 借鉴 mcpPropertyToZod 模式
function excelPropertyToSchema(type: ExcelPropertyType): z.ZodTypeAny {
    const mapping: Record<ExcelPropertyType, () => z.ZodTypeAny> = {
        'range': () => z.string().regex(/^[A-Z]+\d+(:[A-Z]+\d+)?$/),
        'cell': () => z.string().regex(/^[A-Z]+\d+$/),
        'sheetName': () => z.string(),
        'values': () => z.array(z.array(z.unknown())),
        'formula': () => z.string().startsWith('='),
    }
    return mapping[type]?.() ?? z.unknown()
}
```

---

## 3. 总结

| 模块 | 借鉴点 | 优先级 |
|------|--------|--------|
| FlowExecutorContext | 不可变状态 + verdict 判决 | ⭐⭐⭐ |
| error-handling | 指数退避 + continueOnFailure | ⭐⭐⭐ |
| progress.service | 防抖 + Mutex 锁 | ⭐⭐ |
| MCP Service | 动态工具注册 + Zod 映射 | ⭐⭐ |
| Agent Tools | LLM 驱动参数解析 | ⭐ |

---

**文档生成时间**: 2026-01-04
**用途**: 架构参考，可删除原始文件夹
