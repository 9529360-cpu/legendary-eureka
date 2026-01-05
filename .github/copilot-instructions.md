# Excel 智能助手 Add-in - Copilot 开发指南

---

## 🚨 AI 代码规范（写代码前必读）

> **以下是 AI 写代码时最容易犯的错误，每次写代码前必须检查！**

### ❌ 绝对禁止

| 错误 | 后果 | 正确做法 |
|------|------|----------|
| PowerShell 重定向写文件 `echo "内容" > file` | 中文乱码 | 用 `create_file` 工具或 Node.js |
| 不看现有代码就写新代码 | 风格不一致、重复造轮子 | 先 `read_file` 查看相关文件 |
| 凭空猜测 API 或函数签名 | 运行时错误 | 先搜索现有实现 |
| 修改文件不验证 | 引入 bug 不知道 | 改完运行 `npm run build:dev` |
| 硬编码中文字符串到多处 | 维护困难 | 使用常量或配置 |
| 在 UI 层直接调用 Excel API | 架构违规 | 通过 Agent 工具层调用 |
| **在单文件写超过 500 行代码** | 臃肿难维护 | **拆分到多个模块** |
| **把不同职责混在一起** | 耦合严重 | **层次分离** |

### ✅ 必须遵守

| 规范 | 说明 |
|------|------|
| **先读后写** | 修改任何文件前，先读取该文件了解上下文 |
| **写完验证** | 每次改代码后运行 `npm run build:dev` 检查编译 |
| **UTF-8 编码** | 所有文件必须是 UTF-8 无 BOM |
| **遵循现有模式** | 新代码必须与现有代码风格一致 |
| **中文注释** | 用户可见文本和注释使用中文 |
| **层次分离** | 每层只做一件事，不跨层调用 |
| **单一职责** | 每个文件/函数只负责一个功能 |

---

## 🏛️ v4.0 架构设计原则（必须遵守！）

> **教训：旧版 AgentCore.ts 膨胀到 16000+ 行，无法维护！**

### 核心原则

| 原则 | 说明 | 反例 |
|------|------|------|
| **层次分离** | 每层独立，只通过接口通信 | ❌ Parser 直接调用工具 |
| **单一职责** | 每个模块只做一件事 | ❌ 一个文件又解析又执行 |
| **零工具名** | LLM 只理解意图，不知道工具 | ❌ Prompt 里写 75 个工具名 |
| **规则优先** | 能用规则就不用 LLM | ❌ 简单映射也调 LLM |
| **文件上限** | 单文件不超过 500 行 | ❌ 16000 行的 AgentCore.ts |

### 三层架构职责

```
┌─────────────────────────────────────────────────────────────┐
│  Layer 1: IntentParser (意图解析)                            │
│  - 输入: 用户自然语言                                        │
│  - 输出: IntentSpec (意图规格)                               │
│  - 方法: 调用 LLM，但 Prompt 无工具名                         │
│  - 文件: src/agent/IntentParser.ts (~300行)                  │
├─────────────────────────────────────────────────────────────┤
│  Layer 2: SpecCompiler (规格编译)                            │
│  - 输入: IntentSpec                                          │
│  - 输出: ExecutionPlan (执行计划)                            │
│  - 方法: 纯规则，零 Token 消耗                                │
│  - 文件: src/agent/SpecCompiler.ts (~400行)                  │
├─────────────────────────────────────────────────────────────┤
│  Layer 3: AgentExecutor (执行引擎)                           │
│  - 输入: ExecutionPlan                                       │
│  - 输出: ExecutionResult                                     │
│  - 方法: 调用工具，处理结果                                   │
│  - 文件: src/agent/AgentExecutor.ts (~500行)                 │
└─────────────────────────────────────────────────────────────┘
```

### 添加新功能时的检查清单

1. [ ] 这个功能属于哪一层？（解析/编译/执行）
2. [ ] 会不会让文件超过 500 行？需要拆分吗？
3. [ ] 是否可以用规则实现，而不是调 LLM？
4. [ ] 有没有跨层调用？（禁止！）
5. [ ] 类型定义放在 `types/` 目录了吗？

### 🔴 绝对不要

- ❌ 在 IntentParser 里直接调用 Excel 工具
- ❌ 在 SpecCompiler 里调用 LLM
- ❌ 在 AgentExecutor 里解析用户意图
- ❌ 把新功能随手加到 AgentCore.ts（已废弃）
- ❌ 定义重复的类型（先查 `src/agent/types/`）

---

### 📋 写代码前检查清单

1. [ ] 我读过要修改的文件了吗？
2. [ ] 我知道现有的代码风格吗？
3. [ ] 我用的 API/函数确实存在吗？
4. [ ] 写完后我会验证构建吗？
5. [ ] 我用的是正确的编码方式吗？
6. [ ] **这个改动会让文件超过 500 行吗？**
7. [ ] **这个功能放对层了吗？**

### 🔴 本项目历史踩坑记录

| 错误场景 | 具体问题 | 教训 |
|----------|----------|------|
| dependsOn 字段 | 用工具名 `excel_create_sheet` 做依赖 | 必须用步骤 ID，不是工具名 |
| App.tsx 乱码 | PowerShell 写文件导致中文变 `浣犲ソ` | 用 Node.js 或 VS Code 写文件 |
| README.md 损坏 | 编辑器自动转换编码 | 验证文件编码 |
| 重复定义类型 | 不知道 `src/agent/types/` 已有定义 | 先搜索再定义 |
| Excel API 报错 | 忘记 `ctx.sync()` 或 `load()` | 看 ESLint 警告 |
| 工具不存在 | 调用了不存在的工具函数 | 查 `ExcelAdapter.ts` 工具列表 |
| **AgentCore 膨胀** | **16000+ 行无法维护** | **拆分成三层架构** |
| **降级只写一个单元格** | **excel_write_range 降级丢数据** | **v4.0.1 修复循环写入** |

### 🛠️ 常见修复命令

```bash
# 编码问题
node scripts/clean_encoding.cjs <file>    # 清理不可见字符
node scripts/fix_encoding.js              # 修复 App.tsx

# 构建验证
npm run build:dev                          # 检查 TypeScript 编译
npm run lint:fix                           # 修复 lint 问题

# 缓存问题
.\scripts\clear_wef.ps1                    # 清除 Office 缓存
```

---

## 项目概述

这是一个 AI 驱动的 Excel Office Add-in，使用 **v4.0 意图驱动架构** 将自然语言指令转换为 Excel 操作。

## 核心架构 (v4.0)

```
用户输入 → IntentParser (LLM) → SpecCompiler (规则) → AgentExecutor (执行) → Excel API
               无工具名            零Token消耗           调用工具
```

### v4.0 架构优势

| 特性 | v3.x (旧) | v4.0 (新) |
|------|-----------|-----------|
| LLM System Prompt | 75 个工具名 | **0 个工具名** |
| Token 消耗 | ~3000+ | **~500** |
| 计划生成 | LLM 直接生成 | **规则编译** |
| 依赖管理 | 工具名 (错误) | **step.id (正确)** |
| 代码规模 | 16000+ 行单体 | **~500 行/层** |
| 可维护性 | 差 | **好** |

### v4.0 核心文件

| 文件 | 层级 | 职责 | 行数上限 |
|------|------|------|----------|
| `src/agent/IntentParser.ts` | Layer 1 | 调用 LLM 理解意图，无工具名 | 300 |
| `src/agent/SpecCompiler.ts` | Layer 2 | 纯规则编译，零 Token | 400 |
| `src/agent/AgentExecutor.ts` | Layer 3 | 执行计划，调用工具 | 500 |
| `src/agent/types/intent.ts` | 类型 | 意图类型定义 | 200 |
| `src/taskpane/hooks/useAgentV4.ts` | UI 集成 | 新架构 React Hook | 400 |

### 模块化目录结构 (v4.0+)

```
src/agent/
├── IntentParser.ts           # Layer 1: 意图解析（调LLM，无工具名）
├── SpecCompiler.ts           # Layer 2: 规格编译（纯规则，零Token）
├── AgentExecutor.ts          # Layer 3: 执行引擎（调用工具）
├── AgentCore.ts              # Agent 核心逻辑（旧版，保留兼容）
├── ExcelAdapter.ts           # Excel 工具入口
├── types/                    # 类型定义
│   ├── intent.ts             # IntentSpec, IntentType 等
│   ├── tool.ts               # Tool, ToolResult 等
│   ├── task.ts               # AgentTask, AgentStep 等
│   └── ...
├── workflow/                 # 工作流系统
│   └── WorkflowEngine.ts
├── constants/                # 常量定义
├── registry/                 # 工具注册中心
│   └── ToolRegistry.ts
└── tools/excel/              # Excel 工具分类
    ├── read.ts               # 读取类 (6个)
    ├── write.ts              # 写入类 (2个)
    ├── formula.ts            # 公式类 (5个)
    ├── format.ts             # 格式化类 (6个)
    ├── chart.ts              # 图表类 (2个)
    ├── data.ts               # 数据操作类 (13个)
    ├── sheet.ts              # 工作表类 (6个)
    ├── analysis.ts           # 分析类 (8个)
    ├── advanced.ts           # 高级工具 (11个)
    └── misc.ts               # 其他 (2个)
```

### 关键设计原则

1. **UI 只负责展示** - 所有 Excel 操作通过 Agent 工具层执行，UI 层不直接调用 Excel API
2. **Agent 是核心，工具是外接模块** - AgentCore 不依赖任何特定工具，Excel 只是一个 Adapter
3. **任务驱动而非工具驱动** - Agent 基于用户意图规划任务，再选择合适的工具

### 分层职责

| 层 | 目录 | 职责 |
|---|---|---|
| UI 层 | `src/taskpane/components/` | React 组件、用户交互、状态展示 |
| Agent 层 | `src/agent/` | ReAct 循环、任务规划、决策引擎 |
| Core 层 | `src/core/` | 工具注册、执行监控、错误处理 |
| Service 层 | `src/services/` | API 调用、数据扫描 |

## 开发命令

```bash
npm run dev:full      # 同时启动前端(3000)、AI后端(3001)、Mock后端(3002)
npm run build:dev     # 开发构建
npm run start         # 在 Excel Desktop 中调试
npm run lint:fix      # 修复 lint 问题
npm run test          # 运行 Jest 测试
```

## 代码约定

### Excel API 调用

所有 Excel 操作必须在 `Excel.run()` 上下文中执行：

```typescript
// ✅ 正确: 在工具文件中使用 excelRun
import { excelRun } from './common';

async execute() {
  return await excelRun(async (ctx) => {
    const range = ctx.workbook.getActiveCell();
    range.load("address");
    await ctx.sync();
    return { success: true, output: range.address };
  });
}

// ❌ 错误: UI 组件直接调用 Excel API
```

### 添加新 Agent 工具

**推荐方式**: 在对应类别文件中添加 (如 `src/agent/tools/excel/data.ts`)

```typescript
export function createMyNewTool(): Tool {
  return {
    name: "excel_my_tool",           // 必须以 excel_ 前缀
    description: "描述用于 LLM 理解",  // 中文描述
    category: "excel",
    parameters: [/* ToolParameter[] */],
    execute: async (input) => { /* return ToolResult */ }
  };
}
```

然后在该文件的 `createXxxTools()` 函数中注册。

**兼容方式**: 仍可在 `ExcelAdapter.ts` 的 `createExcelTools()` 中直接添加。

### React 组件模式

- 使用 Fluent UI 9 组件 (`@fluentui/react-components`)
- 自定义 hooks 放在 `src/taskpane/hooks/`
- 工具函数放在 `src/taskpane/utils/`
- 组件内使用 `makeStyles()` 定义样式

### 类型定义

- 核心类型: `src/types/index.ts`
- Agent 类型: `src/agent/types/` 模块 (Tool, ToolResult, AgentTask, AgentStep 等)
- 工作流类型: `src/agent/workflow/` (WorkflowEngine, WorkflowState)
- 工具定义: `src/agent/registry/ToolRegistry.ts`
- Excel 工具: `src/agent/tools/excel/` (分类工具实现)

## 关键文件

| 文件 | 用途 |
|---|---|
| [src/agent/IntentParser.ts](src/agent/IntentParser.ts) | v4.0 意图解析器，LLM 无工具名 |
| [src/agent/SpecCompiler.ts](src/agent/SpecCompiler.ts) | v4.0 规格编译器，纯规则零 Token |
| [src/agent/AgentExecutor.ts](src/agent/AgentExecutor.ts) | v4.0 执行引擎 |
| [src/agent/AgentCore.ts](src/agent/AgentCore.ts) | ReAct Agent 核心（旧版兼容） |
| [src/agent/ExcelAdapter.ts](src/agent/ExcelAdapter.ts) | Excel 工具入口，兼容层 |
| [src/agent/types/](src/agent/types/) | Agent 类型定义 (Tool, ToolResult, AgentTask 等) |
| [src/agent/workflow/](src/agent/workflow/) | 工作流引擎 (WorkflowEngine) |
| [src/agent/registry/](src/agent/registry/) | 工具注册中心 |
| [src/agent/tools/excel/](src/agent/tools/excel/) | 分类 Excel 工具 (61个工具) |
| [src/taskpane/hooks/useAgent.ts](src/taskpane/hooks/useAgent.ts) | UI 调用 Agent 的 hook (旧版) |
| [src/taskpane/hooks/useAgentV4.ts](src/taskpane/hooks/useAgentV4.ts) | UI 调用 Agent 的 hook (v4.0) |
| [src/services/ApiService.ts](src/services/ApiService.ts) | AI 后端 API 调用 |
| [ai-backend.cjs](ai-backend.cjs) | Express 后端，对接 DeepSeek API |
| [manifest.xml](manifest.xml) | Office Add-in 清单 |

## AI 后端

- 使用 CommonJS (`.cjs`) 以兼容旧版 Node 环境
- DeepSeek API 配置通过 `.env` 文件
- 开发时通过 webpack devServer proxy 避免 CORS 问题

## 测试

测试文件在 `src/__tests__/`，使用 Jest + React Testing Library：

```bash
npm run test:watch    # 监视模式
npm run test:coverage # 覆盖率报告
```

## 注意事项

1. **⚠️ 编码问题（最常见！）**: 中文乱码问题频繁出现，详见 `docs/TROUBLESHOOTING.md` 顶部
2. **Office Add-in 环境**: 必须使用 HTTPS (`https://localhost:3000`)
3. **ESLint 规则**: 启用了 `eslint-plugin-office-addins`，确保正确使用 `ctx.sync()`
4. **中文支持**: 所有用户可见文本使用中文

## 编码问题速查 ⚠️

> **这是本项目最频繁出现的问题！每次写文件前必须注意！**

```powershell
# ✅ 正确：用 Node.js 或 VS Code 创建文件
node -e "require('fs').writeFileSync('file.md', '内容', 'utf8')"

# ❌ 错误：PowerShell 重定向（会产生乱码！）
echo "内容" > file.md

# 修复乱码文件
node scripts/clean_encoding.cjs <file>
node scripts/fix_encoding.js  # 专门修复 App.tsx
```

详细说明见 [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)

## 调试技巧

### 清除 WEF 缓存

Office Add-in 开发时，旧代码可能被缓存导致更新不生效。运行以下脚本清除：

```powershell
# 方法1: 使用项目脚本
.\scripts\clear_wef.ps1

# 方法2: 手动删除
Remove-Item -Recurse -Force "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
```

清除后需重启 Excel。

### 常用调试脚本

| 脚本 | 用途 |
|---|---|
| `scripts/clear_wef.ps1` | 清除 Office Add-in 缓存 |
| `scripts/stop_port_3000.ps1` | 释放被占用的 3000 端口 |
| `scripts/fix_encoding.js` | 修复 App.tsx 中文乱码 |
| `scripts/clean_encoding.cjs` | 清理任意文件的不可见字符 |
| `npm run stop` | 停止 Office Add-in 调试会话 |

### 调试流程

1. `npm run dev:full` 启动所有服务
2. `npm run start` 在 Excel Desktop 中加载 Add-in
3. 在 Excel 中按 `F12` 打开开发者工具
4. 如遇缓存问题，关闭 Excel → 运行 `clear_wef.ps1` → 重新启动
