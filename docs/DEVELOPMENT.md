# 开发指南

## 环境要求

- Node.js >= 18
- npm >= 9
- Excel Desktop (Windows)

## 快速开始

```bash
# 1. 安装依赖
npm install

# 2. 启动所有服务 (前端3000 + AI后端3001 + Mock后端3002)
npm run dev:full

# 3. 在 Excel Desktop 中加载 Add-in
npm run start
```

## 常用命令

| 命令 | 说明 |
|------|------|
| `npm run dev:full` | 启动所有服务 |
| `npm run dev-server` | 只启动前端 (https://localhost:3000) |
| `npm run build:dev` | 开发构建 |
| `npm run build` | 生产构建 |
| `npm run start` | 在 Excel Desktop 中调试 |
| `npm run stop` | 停止调试 |
| `npm run lint` | 检查代码规范 |
| `npm run lint:fix` | 自动修复 |
| `npm run test` | 运行测试 |

## 调试技巧

### 清除 Office Add-in 缓存

```powershell
.\scripts\clear_wef.ps1
```

### 释放被占用的端口

```powershell
.\scripts\stop_port_3000.ps1
```

### 查看 Add-in 日志

在 Excel 中按 `F12` 打开开发者工具。

## 代码规范

### Excel API 调用

所有 Excel 操作必须在 `Excel.run()` 上下文中：

```typescript
import { excelRun } from './common';

async execute() {
  return await excelRun(async (ctx) => {
    const range = ctx.workbook.getActiveCell();
    range.load("address");
    await ctx.sync();
    return { success: true, output: range.address };
  });
}
```

### 添加新工具

在 `src/agent/tools/excel/` 对应类别文件中添加：

```typescript
export function createMyTool(): Tool {
  return {
    name: "excel_my_tool",
    description: "工具描述",
    category: "excel",
    parameters: [...],
    execute: async (input) => { ... }
  };
}
```

## 目录结构

```
src/
├── agent/           # Agent 核心
│   ├── AgentCore.ts
│   ├── IntentParser.ts    # 意图解析 (TODO)
│   ├── SpecCompiler.ts    # 规格编译 (TODO)
│   ├── types/
│   ├── workflow/
│   ├── registry/
│   └── tools/excel/       # Excel 工具
├── taskpane/        # UI 层
│   ├── components/
│   └── hooks/
├── services/        # API 服务
└── core/            # 核心工具类
```

详见 [ARCHITECTURE.md](../ARCHITECTURE.md)
