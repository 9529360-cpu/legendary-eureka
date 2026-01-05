# Excel 智能助手 - 系统架构设计

> **版本**: v4.0 (待实现)  
> **当前状态**: v3.1.1 架构有严重问题，需要重构  
> **最后更新**: 2026-01-05

---

## 一、目标架构 (v4.0)

```
┌──────────────────────────────────────────────────────────────────────────────┐
│                              系统架构图                                       │
├──────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  用户输入: "请给我一个标准的销售表格结构"                                       │
│                  │                                                           │
│                  ▼                                                           │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │              Layer 1: IntentParser (意图理解层)                         │ │
│  │                                                                        │ │
│  │  职责: 调用 LLM 理解用户意图，输出高层规格                                │ │
│  │  输入: 用户自然语言 + 上下文                                             │ │
│  │  输出: IntentSpec (不含任何工具名!)                                      │ │
│  │                                                                        │ │
│  │  示例输出:                                                              │ │
│  │  {                                                                     │ │
│  │    "intent": "create_table",                                           │ │
│  │    "confidence": 0.95,                                                 │ │
│  │    "spec": {                                                           │ │
│  │      "tableType": "sales",                                             │ │
│  │      "columns": ["日期", "产品", "数量", "单价", "金额"],                 │ │
│  │      "location": "current_sheet",                                      │ │
│  │      "options": { "hasTotalRow": true, "hasHeader": true }             │ │
│  │    },                                                                  │ │
│  │    "needsClarification": false                                         │ │
│  │  }                                                                     │ │
│  │                                                                        │ │
│  │  ★ LLM System Prompt 只包含业务概念，不包含任何工具名                     │ │
│  │  ★ Token 消耗少，出错率低                                               │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                  │                                                           │
│                  ▼                                                           │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │              Layer 2: SpecCompiler (规格编译层)                         │ │
│  │                                                                        │ │
│  │  职责: 将高层规格编译成工具调用序列                                       │ │
│  │  输入: IntentSpec                                                       │ │
│  │  输出: ExecutionPlan (合法的工具调用序列)                                 │ │
│  │                                                                        │ │
│  │  核心规则 (硬编码，不调LLM):                                             │ │
│  │  - create_table → [read_current, write_headers, format, set_formulas]  │ │
│  │  - analyze_data → [read_range, compute_stats, respond]                 │ │
│  │  - 自动处理依赖顺序 (step.id 引用，不是工具名!)                           │ │
│  │  - 自动补充感知步骤 (写之前必须读)                                        │ │
│  │  - 参数完整性验证                                                       │ │
│  │                                                                        │ │
│  │  ★ 纯 TypeScript 逻辑，零 Token 消耗                                    │ │
│  │  ★ 规则可测试、可调试                                                   │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                  │                                                           │
│                  ▼                                                           │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │              Layer 3: AgentExecutor (执行引擎层)                        │ │
│  │                                                                        │ │
│  │  职责: 执行计划、管理状态、处理异常、结果校验                              │ │
│  │  输入: ExecutionPlan                                                    │ │
│  │  输出: ExecutionResult                                                  │ │
│  │                                                                        │ │
│  │  核心能力:                                                              │ │
│  │  1. 状态管理: pending → running → success/failed                        │ │
│  │  2. 依赖检查: 按 step.id 检查，非工具名                                  │ │
│  │  3. 工具执行: 调用 ToolRegistry                                         │ │
│  │  4. 结果校验: 验证写入是否成功                                           │ │
│  │  5. 错误恢复: 失败时回调 LLM 修复                                        │ │
│  │  6. 进度通知: emit 事件给 UI                                            │ │
│  │                                                                        │ │
│  │  ★ 有智能的执行器，不是傻代理                                            │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                  │                                                           │
│                  ▼                                                           │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │              Layer 4: ToolRegistry (工具注册层)                         │ │
│  │                                                                        │ │
│  │  职责: 管理所有 Excel 工具的注册和执行                                    │ │
│  │  工具分类:                                                              │ │
│  │  - read: excel_read_range, excel_read_selection, get_table_schema      │ │
│  │  - write: excel_write_range, excel_write_cell                          │ │
│  │  - format: excel_format_range, excel_auto_fit, excel_merge_cells       │ │
│  │  - formula: excel_set_formula, excel_batch_formula                     │ │
│  │  - data: excel_sort, excel_filter, excel_remove_duplicates             │ │
│  │  - chart: excel_create_chart, excel_chart_trendline                    │ │
│  │  - sheet: excel_create_sheet, excel_switch_sheet                       │ │
│  │                                                                        │ │
│  │  ★ 工具只被 AgentExecutor 调用，LLM 不知道这些名字                       │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                  │                                                           │
│                  ▼                                                           │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │              Layer 5: Excel JavaScript API                              │ │
│  │                                                                        │ │
│  │  Office.Excel.run() → 实际操作 Excel                                    │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                                                                              │
└──────────────────────────────────────────────────────────────────────────────┘
```

---

## 二、当前架构 (v3.1.1) 的问题

```
┌──────────────────────────────────────────────────────────────────────────────┐
│                         当前架构 (有严重问题)                                  │
├──────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  用户输入                                                                    │
│       │                                                                      │
│       ▼                                                                      │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │                    LLM (DeepSeek)                                       │ │
│  │                                                                        │ │
│  │  ❌ 问题1: System Prompt 包含 75 个工具名和参数                          │ │
│  │  ❌ 问题2: LLM 直接生成工具调用序列                                       │ │
│  │  ❌ 问题3: Token 浪费严重                                                │ │
│  │  ❌ 问题4: LLM 不懂 Excel API 约束，生成的序列常常无效                     │ │
│  │                                                                        │ │
│  │  输出: {"steps": [{"action": "excel_xxx", ...}]}                        │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│       │                                                                      │
│       ▼                                                                      │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │                    Agent (AgentCore.ts - 16000行)                       │ │
│  │                                                                        │ │
│  │  ❌ 问题5: 变成了 JSON 执行器，没有智能                                   │ │
│  │  ❌ 问题6: dependsOn 用工具名做ID，设计混乱                               │ │
│  │  ❌ 问题7: 代码膨胀，职责不清                                             │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│       │                                                                      │
│       ▼                                                                      │
│  ┌────────────────────────────────────────────────────────────────────────┐ │
│  │                    Tool Layer (75个工具)                                │ │
│  └────────────────────────────────────────────────────────────────────────┘ │
│                                                                              │
└──────────────────────────────────────────────────────────────────────────────┘

错误根因:
1. 缺少 SpecCompiler 层 - 没有规格到操作的编译
2. LLM 职责越界 - 不应该知道工具细节
3. Agent 职责缺失 - 变成傻执行器
```

---

## 三、已发现问题记录

| 日期 | 问题 | 根因 | 状态 |
|------|------|------|------|
| 2026-01-05 | "依赖不存在的步骤 excel_create_sheet" | dependsOn 用工具名而非 step.id | 临时修复 |
| 2026-01-05 | LLM 生成的计划常失败 | LLM 不懂 Excel API 约束 | 未解决 |
| 2026-01-05 | Token 消耗过高 | System Prompt 包含 75 个工具 | 未解决 |
| 2026-01-05 | Agent 没有智能 | 架构设计错误，缺少编译层 | 未解决 |

---

## 四、重构计划

### Phase 1: 创建 IntentParser (意图解析器)

**文件**: `src/agent/IntentParser.ts`

```typescript
interface IntentSpec {
  intent: IntentType;
  confidence: number;
  spec: Record<string, unknown>;
  needsClarification: boolean;
  clarificationQuestion?: string;
}

type IntentType = 
  | 'create_table'      // 创建表格
  | 'write_data'        // 写入数据
  | 'format_range'      // 格式化
  | 'analyze_data'      // 分析数据
  | 'create_formula'    // 创建公式
  | 'create_chart'      // 创建图表
  | 'query_data'        // 查询数据
  | 'clarify';          // 需要澄清

class IntentParser {
  // LLM 只理解意图，不知道工具
  async parse(userMessage: string, context: Context): Promise<IntentSpec>;
}
```

### Phase 2: 创建 SpecCompiler (规格编译器)

**文件**: `src/agent/SpecCompiler.ts`

```typescript
class SpecCompiler {
  // 纯规则，不调 LLM
  compile(spec: IntentSpec): ExecutionPlan {
    switch (spec.intent) {
      case 'create_table':
        return this.compileCreateTable(spec);
      case 'analyze_data':
        return this.compileAnalyzeData(spec);
      // ...
    }
  }
  
  private compileCreateTable(spec: IntentSpec): ExecutionPlan {
    const steps: PlanStep[] = [];
    const id1 = generateId();
    const id2 = generateId();
    
    // 自动补充感知步骤
    steps.push({ id: id1, action: 'excel_read_selection', dependsOn: [] });
    
    // 写入表头
    steps.push({ id: id2, action: 'excel_write_range', dependsOn: [id1], ... });
    
    // 正确的依赖：用 step.id，不是工具名
    return { steps };
  }
}
```

### Phase 3: 重构 AgentExecutor

**文件**: `src/agent/AgentExecutor.ts`

```typescript
class AgentExecutor {
  private intentParser: IntentParser;
  private specCompiler: SpecCompiler;
  private toolRegistry: ToolRegistry;
  
  async execute(userMessage: string): Promise<string> {
    // 1. 意图理解
    const intent = await this.intentParser.parse(userMessage, context);
    
    if (intent.needsClarification) {
      return intent.clarificationQuestion;
    }
    
    // 2. 编译成操作序列
    const plan = this.specCompiler.compile(intent);
    
    // 3. 执行并校验
    for (const step of plan.steps) {
      const result = await this.executeStep(step);
      if (!result.success) {
        // 错误恢复逻辑
      }
    }
    
    return this.generateResponse(results);
  }
}
```

---

## 五、目录结构 (目标)

```
src/agent/
├── IntentParser.ts       # Layer 1: 意图理解 (调LLM)
├── SpecCompiler.ts       # Layer 2: 规格编译 (纯规则)
├── AgentExecutor.ts      # Layer 3: 执行引擎
├── types/
│   ├── intent.ts         # IntentSpec, IntentType
│   ├── plan.ts           # ExecutionPlan, PlanStep
│   └── ...
├── registry/
│   └── ToolRegistry.ts   # Layer 4: 工具注册
├── tools/excel/
│   └── ...               # 具体工具实现
└── compiler/
    ├── rules/            # 编译规则
    │   ├── create-table.ts
    │   ├── analyze-data.ts
    │   └── ...
    └── validators/       # 参数验证器
```

---

## 六、迁移策略

1. **创建新层，不删旧代码** - 新旧并行
2. **逐步迁移意图类型** - 先迁移 create_table，验证后再迁移其他
3. **保持向后兼容** - 旧路径作为 fallback
4. **完整测试后切换** - 确认新架构稳定再废弃旧代码

---

## 七、文档维护规则

1. **架构变更必须先更新此文档**
2. **发现问题必须记录到"已发现问题记录"表**
3. **每次重构必须更新目录结构图**
4. **禁止直接写代码不更新文档**
