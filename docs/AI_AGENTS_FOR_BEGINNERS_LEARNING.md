# 🎓 AI Agents for Beginners - 学习笔记

> **来源**: [microsoft/ai-agents-for-beginners](https://github.com/microsoft/ai-agents-for-beginners)
> **学习日期**: 2025年
> **目的**: 从微软官方教程中提取最佳实践，应用到 Excel Copilot 项目

## 📚 课程概览

这是一个 15 课的全面教程，涵盖 AI Agent 的从基础到生产的完整知识体系：

| 课程 | 主题 | 与我们项目的相关性 |
|------|------|------------------|
| 03 | Agentic Design Patterns | ⭐⭐⭐⭐⭐ 核心设计原则 |
| 04 | Tool Use | ⭐⭐⭐⭐⭐ 工具调用模式 |
| 06 | Building Trustworthy Agents | ⭐⭐⭐⭐⭐ 安全与人工介入 |
| 07 | Planning Design | ⭐⭐⭐⭐⭐ 任务分解与规划 |
| 08 | Multi-Agent | ⭐⭐⭐⭐ 多智能体协作 |
| 09 | Metacognition | ⭐⭐⭐⭐⭐ 自我反思与改进 |
| 12 | Context Engineering | ⭐⭐⭐⭐⭐ 上下文管理 |
| 13 | Agent Memory | ⭐⭐⭐⭐ 记忆系统 |

---

## 🎯 核心概念提取

### 1. Agentic Design Principles（第 3 课）

#### Agent 的三个维度
```
Agent Space（空间）: 连接性、可访问性
Agent Time（时间）: 过去（学习）、现在（执行）、未来（规划）
Agent Core（核心）: 信任与可靠性
```

#### 设计指导原则

| 原则 | 定义 | 应用到 Excel Copilot |
|------|------|---------------------|
| **Transparency（透明性）** | Agent 应该清晰地解释其能力和局限 | 告知用户工具的边界 |
| **Control（控制）** | 用户应能够控制 Agent 的行为 | 高风险操作需要确认 |
| **Consistency（一致性）** | Agent 应有一致的行为模式 | 相同输入应有相似输出 |

---

### 2. Tool Use Design Pattern（第 4 课）

#### 工具定义最佳结构

```typescript
// 推荐的工具 Schema 结构
interface ToolDefinition {
  name: string;                    // 唯一标识符
  description: string;             // LLM 用于理解的描述
  parameters: {
    type: "object";
    properties: {
      [key: string]: {
        type: string;
        description: string;       // 每个参数也需要描述！
      }
    };
    required: string[];
  };
}
```

#### Semantic Kernel Plugin 模式

```python
# 使用装饰器定义工具
class DestinationsPlugin:
    @kernel_function(description="Provides a list of vacation destinations.")
    def get_destinations(self) -> Annotated[str, "Returns the destinations."]:
        return "..."
    
    @kernel_function(description="Provides the availability of a destination.")
    def get_availability(
        self, 
        destination: Annotated[str, "The destination to check."]
    ) -> Annotated[str, "Returns the availability."]:
        return "..."
```

**💡 启发**: 我们的 `ExcelAdapter.ts` 中的工具可以增加更丰富的 description

---

### 3. Building Trustworthy Agents（第 6 课）

#### System Message Framework

```
┌─────────────────────────────────────────┐
│           Meta Prompt                    │
│  (公司政策、安全规则、角色定义)          │
├─────────────────────────────────────────┤
│           Basic Prompt                   │
│  (任务指令、输出格式、约束条件)          │
├─────────────────────────────────────────┤
│        LLM Optimization                  │
│  (CoT、Few-shot、结构化输出)            │
└─────────────────────────────────────────┘
```

#### Agent 安全威胁分类

| 威胁类型 | 描述 | 缓解策略 |
|----------|------|----------|
| **Task Manipulation** | 用户试图操纵 Agent 执行非预期任务 | 输入验证、范围限制 |
| **System Access** | 通过 Agent 获取系统权限 | 最小权限原则 |
| **Resource Overloading** | 消耗过多计算资源 | 速率限制、超时 |
| **Knowledge Poisoning** | 注入错误信息影响未来决策 | 信息验证、隔离 |
| **Cascading Errors** | 错误在多轮对话中累积 | 上下文重置、验证点 |

#### Human-in-the-Loop 模式

```python
# AutoGen 示例：人工确认中断
termination = TextMentionTermination("APPROVE") | MaxMessageTermination(10)

async def handle_approval():
    while not approved:
        user_input = await get_user_input("Do you approve? (APPROVE/REJECT)")
        if user_input == "APPROVE":
            return True
        elif user_input == "REJECT":
            return False
```

**💡 我们已实现**: `ApprovalManager.ts` + `ApprovalDialog.tsx`

---

### 4. Planning Design（第 7 课）

#### 结构化输出（Pydantic 模式）

```python
from pydantic import BaseModel, Field
from typing import List

class SubTask(BaseModel):
    assigned_agent: str = Field(
        description="The specific agent assigned to handle this subtask"
    )
    task_details: str = Field(
        description="Detailed description of what needs to be done"
    )

class TravelPlan(BaseModel):
    main_task: str = Field(
        description="The overall travel request from the user"
    )
    subtasks: List[SubTask] = Field(
        description="List of subtasks broken down from the main task"
    )

# 使用结构化输出
settings = OpenAIChatPromptExecutionSettings(response_format=TravelPlan)
```

**💡 启发**: 我们可以用 Zod 实现类似的类型安全输出解析

```typescript
// TypeScript 版本
import { z } from 'zod';

const SubTaskSchema = z.object({
  toolName: z.string().describe("The tool to execute"),
  parameters: z.record(z.unknown()).describe("Tool parameters"),
  rationale: z.string().describe("Why this step is needed")
});

const ExecutionPlanSchema = z.object({
  mainTask: z.string(),
  subtasks: z.array(SubTaskSchema),
  estimatedRisk: z.enum(["low", "medium", "high"])
});
```

#### Semantic Router Agent 模式

根据用户意图动态路由到不同的专业 Agent：
```
用户输入 → Router Agent → [FlightAgent | HotelAgent | CarAgent | ...]
```

---

### 5. Multi-Agent Design（第 8 课）

#### 何时使用多 Agent

| 场景 | 单 Agent | 多 Agent |
|------|----------|----------|
| 简单任务 | ✅ | ❌ |
| 大工作量 | ❌ | ✅ 并行处理 |
| 需要专业知识 | ❌ | ✅ 专业化分工 |
| 需要容错 | ❌ | ✅ 故障隔离 |

#### 多 Agent 模式

```
1. Group Chat（群聊）
   Agent A ←→ Agent B ←→ Agent C
   用于：团队协作、头脑风暴

2. Hand-off（接力）
   Agent A → Agent B → Agent C
   用于：工作流程、审批链

3. Collaborative Filtering（协同过滤）
   User Query
      ↓
   ┌──────────┐
   │ Agent A  │ (行业专家)
   │ Agent B  │ (技术分析)
   │ Agent C  │ (基本面分析)
   └──────────┘
      ↓
   综合推荐
```

#### 可见性与监控

```typescript
interface AgentInteractionLog {
  timestamp: Date;
  fromAgent: string;
  toAgent: string;
  messageType: 'query' | 'response' | 'handoff';
  content: string;
  metrics: {
    latencyMs: number;
    tokensUsed: number;
  };
}
```

---

### 6. Metacognition（第 9 课）

#### 元认知定义

> "Thinking about thinking" - 让 Agent 具备自我反思能力

#### 三大组成部分

```
┌─────────────────────────────────────────┐
│               Persona                    │
│  Agent 的角色定位和行为风格              │
├─────────────────────────────────────────┤
│               Tools                      │
│  可用的外部能力（API、函数）             │
├─────────────────────────────────────────┤
│               Skills                     │
│  内化的知识和推理能力                    │
└─────────────────────────────────────────┘
```

#### Corrective RAG 模式

```
用户查询 → 检索文档 → 验证相关性 → 
  ├── 相关 → 生成回答
  └── 不相关 → 修正查询 → 重新检索
```

#### 自适应 Agent 指令模板

```typescript
const ADAPTIVE_AGENT_INSTRUCTIONS = `
Your process for assisting users:

1. MAINTAIN a customer_preferences object throughout the conversation
2. RECORD their choices in the preferences object
3. For subsequent inquiries, AUTOMATICALLY apply existing preferences
4. Explicitly say: "Based on your previous preference for [X], I recommend..."
5. After each action, UPDATE the preferences object
6. ALWAYS mention which preference you used when making suggestions

Guidelines:
- Always seek feedback to ensure suggestions meet expectations
- Acknowledge when a request falls outside your capabilities
- When giving suggestions, reflect if they are reasonable. Respond again if not.
`;
```

**💡 启发**: 增强我们的 `ResponseTemplates.ts` 包含偏好记忆

---

### 7. Context Engineering（第 12 课）

#### Prompt Engineering vs Context Engineering

| 方面 | Prompt Engineering | Context Engineering |
|------|-------------------|---------------------|
| 范围 | 单次静态指令 | 动态信息管理 |
| 时间跨度 | 单轮对话 | 多轮、多会话 |
| 关注点 | 如何表达指令 | 如何管理信息流 |

#### 上下文类型

```typescript
interface AgentContext {
  // 1. Instructions - 规则和指令
  instructions: {
    systemPrompt: string;
    fewShotExamples: Example[];
    toolDescriptions: ToolDescription[];
  };
  
  // 2. Knowledge - 知识库
  knowledge: {
    factDatabase: Fact[];
    ragResults: Document[];
    longTermMemory: Memory[];
  };
  
  // 3. Tools - 工具定义和结果
  tools: {
    definitions: ToolDefinition[];
    callHistory: ToolCallResult[];
  };
  
  // 4. Conversation - 对话历史
  conversation: {
    messages: Message[];
    summary?: string;  // 压缩后的摘要
  };
  
  // 5. User Preferences - 用户偏好
  userPreferences: {
    settings: Record<string, unknown>;
    pastInteractions: Interaction[];
  };
}
```

#### 上下文管理策略

| 策略 | 描述 | 实现方式 |
|------|------|----------|
| **Agent Scratchpad** | 单次会话的临时笔记 | 运行时对象/文件 |
| **Memories** | 跨会话的持久记忆 | 数据库/向量存储 |
| **Compressing** | 压缩过长的上下文 | 摘要/裁剪 |
| **Multi-Agent** | 分散到多个 Agent | 每个 Agent 独立上下文 |
| **Sandbox** | 隔离代码执行 | 仅返回结果 |
| **Runtime State** | 子任务状态容器 | 结构化状态对象 |

#### ⚠️ 常见上下文失败模式

| 失败类型 | 症状 | 解决方案 |
|----------|------|----------|
| **Context Poisoning** | 幻觉进入上下文并被反复引用 | 验证 + 隔离 |
| **Context Distraction** | 上下文过大导致模型分心 | 定期摘要 |
| **Context Confusion** | 工具太多导致选择错误 | RAG 动态加载工具（<30个）|
| **Context Clash** | 上下文中存在矛盾信息 | 修剪 + 覆盖旧信息 |

**💡 启发**: 我们的 `ConversationMemory.ts` 需要实现摘要压缩

```typescript
// 建议添加到 ConversationMemory.ts
class ConversationMemory {
  private static readonly MAX_MESSAGES = 20;
  private static readonly COMPRESSION_THRESHOLD = 15;
  
  async addMessage(message: Message): Promise<void> {
    this.messages.push(message);
    
    if (this.messages.length > this.COMPRESSION_THRESHOLD) {
      await this.compressOldMessages();
    }
  }
  
  private async compressOldMessages(): Promise<void> {
    const oldMessages = this.messages.slice(0, -5);  // 保留最近5条
    const summary = await this.summarize(oldMessages);
    
    this.messages = [
      { role: 'system', content: `[Previous conversation summary: ${summary}]` },
      ...this.messages.slice(-5)
    ];
  }
}
```

---

### 8. Agent Memory（第 13 课）

#### 记忆类型体系

```
┌─────────────────────────────────────────────────────────┐
│                    Memory Types                          │
├─────────────────────────────────────────────────────────┤
│  Working Memory      单任务过程中的即时信息              │
│  ├── 当前需求、决策、行动                               │
│                                                          │
│  Short-Term Memory   单会话上下文                        │
│  ├── 对话历史、当前状态                                 │
│                                                          │
│  Long-Term Memory    跨会话持久信息                      │
│  ├── 用户偏好、历史交互                                 │
│                                                          │
│  Persona Memory      Agent 角色一致性                    │
│  ├── 专家身份、语气风格                                 │
│                                                          │
│  Episodic Memory     工作流程记录                        │
│  ├── 成功/失败的步骤序列                                │
│                                                          │
│  Entity Memory       提取的实体信息                      │
│  ├── 人名、地点、事件                                   │
└─────────────────────────────────────────────────────────┘
```

#### Self-Improving Agent 模式

```
┌─────────────────────────────────────────┐
│           Main Agent                     │
│      (执行用户任务)                      │
└────────────┬────────────────────────────┘
             │ 观察
             ▼
┌─────────────────────────────────────────┐
│         Knowledge Agent                  │
│  1. 识别有价值的信息                    │
│  2. 提取并摘要                          │
│  3. 存储到知识库                        │
│  4. 增强未来查询                        │
└─────────────────────────────────────────┘
             │
             ▼
┌─────────────────────────────────────────┐
│         Vector Database                  │
│     (存储提取的知识)                     │
└─────────────────────────────────────────┘
```

#### 优化策略

```typescript
// 延迟管理：使用轻量模型快速判断
async function shouldStoreMemory(content: string): Promise<boolean> {
  // 用便宜快速的模型判断
  const importance = await lightweightModel.classify(content);
  return importance > 0.7;
}

// 冷热存储分层
interface MemoryStorage {
  hot: InMemoryCache;      // 高频访问
  warm: Redis;             // 中频访问
  cold: Blob;              // 低频归档
}
```

---

## 🔧 应用到 Excel Copilot 的行动项

### 立即可实施

1. **✅ 已完成 - 人工确认机制**
   - `ApprovalManager.ts` - Agent 层风险评估
   - `ApprovalDialog.tsx` - UI 确认对话框

2. **📝 待实施 - 结构化输出解析**
   ```typescript
   // 使用 Zod 验证 LLM 输出
   const ExecutionPlanSchema = z.object({
     operation: z.enum(["execute", "ask", "clarify"]),
     steps: z.array(StepSchema),
     estimatedRisk: z.enum(["low", "medium", "high"])
   });
   ```

3. **📝 待实施 - 上下文压缩**
   ```typescript
   // ConversationMemory 添加摘要功能
   async compressContext(): Promise<void>;
   ```

### 中期改进

4. **📝 System Message Framework**
   - 分层 prompt 结构
   - Meta prompt (安全规则) + Basic prompt (任务) + Optimization (CoT)

5. **📝 工具描述增强**
   - 每个参数添加 description
   - 添加使用示例

6. **📝 Episodic Memory**
   - 记录成功/失败的操作序列
   - 用于改进未来执行

### 长期架构

7. **📝 Multi-Agent 支持**
   - 规划 Agent + 执行 Agent 分离
   - 专业化工具 Agent

8. **📝 Self-Improving 机制**
   - Knowledge Agent 观察主 Agent
   - 自动提取有价值信息

---

## 📖 关键代码模式速查

### Tool Definition (最佳实践)

```typescript
{
  name: "excel_write_cell",
  description: "Write a value to a specific cell. Use this for single-cell updates.",
  parameters: {
    type: "object",
    properties: {
      cell: {
        type: "string",
        description: "Cell address in A1 notation (e.g., 'A1', 'B5')"
      },
      value: {
        type: "string",
        description: "The value to write. Can be text, number, or formula starting with '='"
      }
    },
    required: ["cell", "value"]
  }
}
```

### Adaptive Instructions

```typescript
const AGENT_INSTRUCTIONS = `
You are an Excel assistant that helps users with spreadsheet tasks.

Your process:
1. MAINTAIN a task_context object throughout the conversation
2. RECORD user preferences and past decisions
3. For subsequent requests, APPLY learned preferences automatically
4. Explicitly say: "Based on your previous preference for [X], I'll..."
5. After each action, UPDATE the context with new learnings
6. ALWAYS explain which preference influenced your decision

When uncertain:
- Ask ONE clarifying question at a time
- Offer 2-3 specific options when possible
- Acknowledge limitations honestly

Self-reflection:
- After generating a plan, verify each step is achievable
- If a step seems risky, flag it for user confirmation
- Learn from errors and adjust approach
`;
```

### Context Management

```typescript
class ContextManager {
  private maxTokens = 8000;
  
  async buildContext(session: Session): Promise<Context> {
    return {
      instructions: await this.getInstructions(),
      knowledge: await this.retrieveRelevantKnowledge(session.query),
      tools: this.selectRelevantTools(session.query, { maxTools: 25 }),
      conversation: await this.getCompressedHistory(session.id),
      preferences: await this.getUserPreferences(session.userId)
    };
  }
  
  selectRelevantTools(query: string, options: { maxTools: number }): Tool[] {
    // RAG over tool descriptions
    const ranked = this.rankToolsByRelevance(query, this.allTools);
    return ranked.slice(0, options.maxTools);
  }
}
```

---

## 🔗 参考资源

- [课程 GitHub 仓库](https://github.com/microsoft/ai-agents-for-beginners)
- [Azure AI Foundry Discord](https://aka.ms/ai-agents/discord)
- [Semantic Kernel 文档](https://learn.microsoft.com/semantic-kernel/)
- [AutoGen 设计模式](https://microsoft.github.io/autogen/stable/user-guide/core-user-guide/design-patterns/intro.html)

---

## ✨ 总结

这个教程是目前看到的**最有价值**的学习资源，因为它：

1. **系统性** - 涵盖从基础到生产的完整知识体系
2. **实践性** - 每课都有可运行的代码示例
3. **权威性** - 来自微软官方，与 Azure AI 生态深度集成
4. **前沿性** - 涵盖 Context Engineering、Metacognition 等最新概念

对于我们的 Excel Copilot 项目，最重要的收获是：

| 概念 | 启发 | 优先级 |
|------|------|--------|
| **Context Engineering** | 上下文管理远比 Prompt Engineering 重要 | ⭐⭐⭐⭐⭐ |
| **Structured Output** | 使用 Zod 验证 LLM 输出结构 | ⭐⭐⭐⭐⭐ |
| **Self-Reflection** | Agent 应该反思自己的输出是否合理 | ⭐⭐⭐⭐ |
| **Tool Selection** | 动态选择相关工具（<30个） | ⭐⭐⭐⭐ |
| **Memory Hierarchy** | 分层记忆系统 | ⭐⭐⭐ |

---

*📌 学习完成后，可以删除 `ai-agents-for-beginners-main` 文件夹*
