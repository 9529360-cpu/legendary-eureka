# API 文档

## 后端接口

### 健康检查

```
GET http://localhost:3001/api/health
```

**响应**:
```json
{ "status": "ok", "timestamp": "2026-01-05T..." }
```

### 聊天接口

```
POST http://localhost:3001/chat
Content-Type: application/json

{
  "message": "用户消息",
  "systemPrompt": "系统提示词（可选）"
}
```

**响应**:
```json
{
  "message": "AI 回复内容"
}
```

### 流式聊天

```
POST http://localhost:3001/chat/stream
Content-Type: application/json

{
  "message": "用户消息",
  "systemPrompt": "系统提示词（可选）"
}
```

**响应**: Server-Sent Events 流

### API Key 配置

```
POST http://localhost:3001/api/config/key
Content-Type: application/json

{ "key": "sk-xxx" }
```

```
GET http://localhost:3001/api/config/status
```

```
DELETE http://localhost:3001/api/config/key
```

---

## Agent 工具接口

### 工具定义

```typescript
interface Tool {
  name: string;           // 工具名称，如 "excel_write_range"
  description: string;    // 工具描述
  category: string;       // 分类，如 "excel"
  parameters: ToolParameter[];
  execute: (input: Record<string, unknown>) => Promise<ToolResult>;
}

interface ToolParameter {
  name: string;
  type: "string" | "number" | "boolean" | "array" | "object";
  description: string;
  required: boolean;
}

interface ToolResult {
  success: boolean;
  output: string;
  error?: string;
  data?: unknown;
}
```

### 工具分类

| 分类 | 文件 | 数量 |
|------|------|------|
| 读取 | tools/excel/read.ts | 6 |
| 写入 | tools/excel/write.ts | 2 |
| 公式 | tools/excel/formula.ts | 5 |
| 格式化 | tools/excel/format.ts | 6 |
| 图表 | tools/excel/chart.ts | 2 |
| 数据操作 | tools/excel/data.ts | 13 |
| 工作表 | tools/excel/sheet.ts | 6 |
| 分析 | tools/excel/analysis.ts | 8 |
| 高级 | tools/excel/advanced.ts | 11 |
| 其他 | tools/excel/misc.ts | 2 |

### 常用工具

#### excel_read_range

读取指定范围数据。

```typescript
{
  address: "A1:D10",  // 必填
  sheet?: "Sheet1"    // 可选
}
```

#### excel_write_range

写入数据到指定范围。

```typescript
{
  address: "A1",      // 必填
  values: [["a", "b"], ["c", "d"]],  // 必填，二维数组
  sheet?: "Sheet1"
}
```

#### excel_set_formula

设置公式。

```typescript
{
  cell: "A10",        // 必填
  formula: "=SUM(A1:A9)",  // 必填
  sheet?: "Sheet1"
}
```

#### excel_format_range

格式化范围。

```typescript
{
  address: "A1:D10",
  bold?: true,
  fontSize?: 12,
  backgroundColor?: "#FFFF00",
  textColor?: "#000000",
  horizontalAlignment?: "center"
}
```

完整工具列表见 `src/agent/tools/excel/index.ts`
