# Excel 智能助手 Add-in - Copilot 开发指南

## 项目概述

这是一个 AI 驱动的 Excel Office Add-in，使用 **ReAct Agent 架构** 将自然语言指令转换为 Excel 操作。

## 核心架构

```
UI Layer (React)  →  Agent Layer  →  Tool Layer  →  Excel JavaScript API
   App.tsx           AgentCore.ts    ExcelAdapter.ts    Office.Excel.run()
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
// ✅ 正确: 在 ExcelAdapter 中封装
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

在 [src/agent/ExcelAdapter.ts](src/agent/ExcelAdapter.ts) 中添加：

```typescript
function createMyNewTool(): Tool {
  return {
    name: "excel_my_tool",           // 必须以 excel_ 前缀
    description: "描述用于 LLM 理解",  // 中文描述
    category: "excel",
    parameters: [/* ToolParameter[] */],
    execute: async (input) => { /* return ToolResult */ }
  };
}
```

然后在 `createExcelTools()` 函数中注册。

### React 组件模式

- 使用 Fluent UI 9 组件 (`@fluentui/react-components`)
- 自定义 hooks 放在 `src/taskpane/hooks/`
- 工具函数放在 `src/taskpane/utils/`
- 组件内使用 `makeStyles()` 定义样式

### 类型定义

- 核心类型: `src/types/index.ts`
- Agent 类型: `src/agent/AgentCore.ts` 中导出
- 工具定义: `src/core/ToolRegistry.ts`

## 关键文件

| 文件 | 用途 |
|---|---|
| [src/agent/AgentCore.ts](src/agent/AgentCore.ts) | ReAct Agent 核心，任务执行循环 |
| [src/agent/ExcelAdapter.ts](src/agent/ExcelAdapter.ts) | 所有 Excel 工具实现 (~90个工具) |
| [src/taskpane/hooks/useAgent.ts](src/taskpane/hooks/useAgent.ts) | UI 调用 Agent 的 hook |
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

1. **Office Add-in 环境**: 必须使用 HTTPS (`https://localhost:3000`)
2. **ESLint 规则**: 启用了 `eslint-plugin-office-addins`，确保正确使用 `ctx.sync()`
3. **中文支持**: 所有用户可见文本使用中文
4. **版本约定**: 遵循 `PROJECT_DOCUMENTATION.md` 中的架构演进记录

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
| `npm run stop` | 停止 Office Add-in 调试会话 |

### 调试流程

1. `npm run dev:full` 启动所有服务
2. `npm run start` 在 Excel Desktop 中加载 Add-in
3. 在 Excel 中按 `F12` 打开开发者工具
4. 如遇缓存问题，关闭 Excel → 运行 `clear_wef.ps1` → 重新启动
