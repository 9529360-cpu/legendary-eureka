# Agent 架构重构备忘录

> 创建日期: 2026-01-04
> 版本: v3.0.0 架构重构

## 核心设计原则

### 黄金法则

**用户是跟 LLM 对话的，不是 Agent。**

```
User ←→ LLM (决策层) ←→ Agent (执行层) ←→ Excel API
```

### 职责划分

| 职责 | 归属 | 说明 |
|-----|------|------|
| 理解用户意图 | LLM | 通过对话历史理解 |
| 决定是否需要澄清 | LLM | 在回复中询问 |
| 决定做什么操作 | LLM | 生成执行计划 JSON |
| 执行 Excel 操作 | Agent | 调用工具，无脑执行 |
| 失败后重新规划 | LLM | Agent 把错误反馈给 LLM |

### 护城河分析

**❌ 不是护城河的东西:**
- `classifyUserIntent()` - LLM 比硬编码规则更准
- `ClarifyGate` - 基于规则不够智能
- 16000 行 Agent 代码 - 复杂度 ≠ 价值

**✅ 真正的护城河:**

1. **工具层 (ExcelAdapter)** - 90+ Excel 工具，领域知识积累
2. **Prompt Engineering** - 让 LLM 选对工具的技巧
3. **闭环执行能力** - 失败 → 反馈 LLM → 重新规划
4. **上下文压缩** - 高效利用 Token

---

## 重构目标

### Before (v2.x)
```
AgentCore.ts: 16000+ 行
- classifyUserIntent() ❌ 删除
- ClarifyGate.decide() ❌ 删除
- detectUserFeedbackType() ❌ 删除
- checkAndSetFollowUpContext() ❌ 删除
- handleFollowUpReply() ❌ 删除
- ... 大量"思考"代码
```

### After (v3.0)
```
AgentCore.ts: ~500 行
- buildPrompt() ✅ 保留并优化
- callLLM() ✅ 保留
- executePlan() ✅ 保留
- executeTool() ✅ 保留
- replan() ✅ 保留 (闭环核心)

ExcelAdapter.ts: 不变
- 90+ 工具 ✅ 核心资产
```

---

## 理想的 Agent 核心逻辑

```typescript
class AgentCore {
  private tools: Tool[];
  
  async run(request: string, context: TaskContext): Promise<AgentTask> {
    // 1. 构建 Prompt (对话历史 + 工具列表 + 当前请求)
    const prompt = this.buildPrompt(request, context);
    
    // 2. 调用 LLM - LLM 决定一切
    const llmResponse = await this.callLLM(prompt);
    
    // 3. 解析 LLM 返回的计划
    const plan = this.parseLLMResponse(llmResponse);
    
    // 4. 如果是纯对话回复，直接返回
    if (plan.isDirectResponse) {
      return { result: plan.message, status: "completed" };
    }
    
    // 5. 执行计划中的每个步骤
    for (const step of plan.steps) {
      const result = await this.executeTool(step);
      
      // 6. 闭环: 失败则反馈给 LLM 重新规划
      if (!result.success) {
        const newPlan = await this.replan(step, result.error, context);
        // 继续执行新计划...
      }
    }
    
    return { result: plan.completionMessage, status: "completed" };
  }
  
  // 核心方法: 构建高质量 Prompt
  private buildPrompt(request: string, context: TaskContext): string {
    // - 对话历史 (裁剪到最近 N 条)
    // - 工作簿上下文 (压缩大表格)
    // - 工具列表 (精简描述)
    // - 用户请求
  }
  
  // 核心方法: 闭环重新规划
  private async replan(failedStep: Step, error: string, context: TaskContext): Promise<Plan> {
    const prompt = this.buildReplanPrompt(failedStep, error, context);
    return await this.callLLM(prompt);
  }
}
```

---

## 重构步骤

### Phase 1: 清理"思考"代码
- [x] 删除 UI 层的关键词检测 (App.tsx pendingFollowUp)
- [ ] 删除 `classifyUserIntent()`
- [ ] 删除 `ClarifyGate` 相关代码
- [ ] 删除 `detectUserFeedbackType()`
- [ ] 删除 `checkAndSetFollowUpContext()`
- [ ] 删除 `handleFollowUpReply()`

### Phase 2: 简化 run() 方法
- [ ] 提取核心流程到新方法
- [ ] 移除分支判断，统一走 LLM
- [ ] 保留闭环重规划逻辑

### Phase 3: 优化 Prompt
- [ ] `buildPrompt()` 包含完整上下文
- [ ] 工具描述精简化
- [ ] 对话历史智能裁剪

### Phase 4: 测试验证
- [ ] 基本对话流程
- [ ] 多轮对话上下文
- [ ] 工具执行闭环
- [ ] 错误重试机制

---

## 注意事项

1. **保留 ExcelAdapter.ts** - 这是核心资产，不动
2. **保留闭环机制** - `triggerReplanForStep()` 是价值所在
3. **保留事件系统** - UI 需要监听执行状态
4. **渐进式重构** - 每步验证，不要一次性大改

---

## 版本记录

| 版本 | 日期 | 变更 |
|-----|------|------|
| v2.9.75 | 2026-01-04 | 移除 UI 关键词检测，添加对话历史到 Prompt |
| v3.0.0 | 2026-01-04 | Agent 架构重构完成 |

---

## v3.0.0 重构完成记录

### 已完成的改动

1. **run() 方法简化** ([AgentCore.ts](src/agent/AgentCore.ts))
   - 删除 `classifyUserIntent()` 调用
   - 删除 `clarifyGate.decide()` 调用  
   - 删除 `detectUserFeedbackType()` 调用
   - 删除 switch/case 分支，统一调用 `executeComplexTask()`
   - 删除 `checkAndSetFollowUpContext()` 调用

2. **buildResponseContext() 简化**
   - 不再调用 `classifyUserIntent()`
   - 直接用 `inferTaskTypeFromPlan()` 从执行计划推断任务类型

3. **UI 层清理** ([App.tsx](src/taskpane/components/App.tsx))
   - 移除 `pendingFollowUp` 关键词检测
   - 保留 `pendingPlanConfirmation` 场景的简单确认/取消检测（这是优化路径）

4. **对话历史集成**
   - `buildPlanGenerationPrompt()` 已包含对话历史
   - LLM 通过历史理解 "好的开始吧" 等确认语

### 保留的代码（暂不删除）

- `classifyUserIntent()` 方法本身 - 不再被调用，但保留备用
- `ClarifyGate` 类 - 不再使用，但保留避免引入更多改动
- `handleFollowUpReply()` 和相关方法 - 已废弃

### 核心流程（v3.0）

```
用户输入
   ↓
App.tsx: 检查是否有 pendingPlanConfirmation？
   ├── 是 → 简单关键词判断确认/取消 → confirmAndExecutePlan()
   └── 否 → 传给 Agent.run()
              ↓
         Agent.run(): 
           1. resolveContextualReferences() // 解析"这里"等指代
           2. executeComplexTask()          // 统一交给 LLM
              ↓
         LLM 决定:
           - 需要更多信息？ → respond_to_user 询问
           - 可以执行？ → 生成工具调用计划
           - 闲聊？ → 直接回复
              ↓
         Agent 执行计划，闭环重规划
```

---

## v3.0.1 工具调用链审计报告

### 1. 工具注册 ✅
- **位置**: `useAgent.ts` 第 149 行
- **机制**: `agent.registerTools(createExcelTools())`
- **工具数量**: 70+ 个 Excel 工具

### 2. 工具列表传递给 LLM ✅ (已修复)
- **问题**: 之前 `buildPlannerSystemPrompt` 只硬编码了 5 个工具
- **修复**: v3.0.1 改为动态生成工具列表
  - 19 个核心工具详细说明
  - 其他 50+ 工具简略列出

### 3. 工具调用流程 ✅
```
executePlanDriven()
   ↓
for each step in plan.steps:
   1. toolRegistry.get(step.action)     // 获取工具
   2. tool.execute(step.parameters)     // 执行
   3. checkStepSuccess(step, result)    // 验证结果
   4. 失败 → triggerReplanForStep()    // 闭环重规划
```

### 4. 执行验证机制 ✅
| 阶段 | 方法 | 验证内容 |
|-----|------|---------|
| 步骤级 | `checkStepSuccess()` | 工具返回 success + 值检查 |
| 读操作 | `verifyReadOperation()` | 返回数据是否有意义 |
| 写操作 | `verifyWriteOperation()` | 重新读取验证数据 |
| 任务级 | `executeVerificationPhase()` | Goal + 抽样 + 错误检查 |

### 5. 闭环重规划 ✅
- **触发**: 工具执行失败或验证失败
- **方法**: `triggerReplanForStep()`
- **上下文**:
  - 失败步骤的参数和错误
  - 已完成步骤的结果
  - 剩余步骤
- **限制**: 最多 3 次重规划

### 6. 潜在风险点
1. ⚠️ **Token 限制**: 工具列表过长可能超出 LLM 上下文
   - 缓解: 只详细说明 19 个核心工具
2. ⚠️ **工具参数兼容性**: LLM 可能生成不兼容的参数格式
   - 缓解: `executePlanDriven` 中有参数转换逻辑 (range → address)
3. ⚠️ **验证延迟**: 写入后验证需要额外 API 调用
   - 可接受: 确保数据正确性比速度更重要

---

## v3.0.3 更新 (2026-01-04)

### Agent 层增强：工具调用控制力

核心原则：**不依赖 LLM 遵守规则，Agent 层主动保障**

#### 1. 强制感知机制 ✅
```typescript
ensurePerceptionBeforeWrite(task, plan)
```
- **时机**: 在 `executePlanDriven` 开始时
- **条件**: 计划包含写操作但没有感知步骤
- **动作**: Agent 自动执行 `excel_read_range` 获取目标区域状态
- **存储**: 感知结果存入 `task.context.perceivedData`

**解决问题**: LLM 不主动调用感知工具时，Agent 层兜底

#### 2. 参数预验证和自动修正 ✅
```typescript
preValidateAndFixParams(toolName, params)
```
在工具执行**前**主动检查和修正：
- 地址格式：中文冒号→英文冒号，自动大写
- values 格式：单值→二维数组，一维→二维
- 公式格式：确保 `=` 开头，中文括号→英文
- 颜色格式：颜色名称→十六进制

**解决问题**: 避免执行失败后才触发 replan，节省时间

#### 3. 增强的 Schema 工具 ✅
`get_table_schema` 现在返回：
```
A列「姓名」: text, 格式=General, 示例=[张三, 李四, 王五]
B列「年龄」: number, 格式=#, 示例=[25, 30, 28]
C列「日期」: date (YYYY-MM-DD), 格式=yyyy-mm-dd, 示例=[2024-01-01, ...]
```
包含：列名、数据类型推断、格式样例、前 3 条样本值

**解决问题**: LLM 能真正理解表格结构，而不是只知道列名

#### 4. 工具调用流程（增强版）
```
executePlanDriven(task)
    │
    ├── ensurePerceptionBeforeWrite()  // 🆕 强制感知
    │
    └── for each step:
            ├── preValidateAndFixParams()  // 🆕 预验证修正
            ├── tool.execute()
            ├── checkStepSuccess()
            └── 失败 → smartRetry() → triggerReplanForStep()
```

### TaskContext 类型扩展
```typescript
interface TaskContext {
  // ... existing fields
  perceivedData?: {
    address: string;
    values: unknown;
    output: string;
    timestamp: Date;
  };
}
```
