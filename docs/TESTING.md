# 测试说明

## 运行测试

```bash
# 运行所有测试
npm run test

# 监视模式
npm run test:watch

# 覆盖率报告
npm run test:coverage
```

## 测试文件位置

```
src/__tests__/
├── App.test.tsx                    # UI 组件测试
├── config-manager.test.ts          # 配置管理测试
├── core-integration.test.ts        # 核心集成测试
├── dynamic-registry.test.ts        # 动态注册测试
├── excel-tools.test.ts             # Excel 工具测试
├── formula-validator-enhanced.test.ts
├── task-monitor.test.ts
├── tool-executor.test.ts
├── trace-context.test.ts
└── validation.test.ts
```

## Agent 端到端测试

```bash
# 运行 Agent 测试
node tests/agent/test-runner.cjs --help

# 运行指定用例
node tests/agent/test-runner.cjs --verbose --case=A1

# 查看测试报告
Get-ChildItem tests/agent/reports/*.md
```

## 手动测试

### 测试 AI 后端

```powershell
# 1. 启动后端
node ai-backend.cjs

# 2. 健康检查
Invoke-RestMethod -Uri "http://localhost:3001/api/health"

# 3. 聊天测试
$body = @{message="你好"; systemPrompt="你是助手"} | ConvertTo-Json
Invoke-RestMethod -Uri "http://localhost:3001/chat" -Method Post -ContentType "application/json" -Body $body
```

### 测试 Excel Add-in

1. 启动服务: `npm run dev:full`
2. 加载 Add-in: `npm run start`
3. 在 Excel 中测试指令:
   - "读取当前选中的内容"
   - "帮我把 A1 到 A10 求和放到 A11"
   - "给这个表格加个标题样式"

## 测试用例设计

| 类别 | 用例 | 预期 |
|------|------|------|
| 读取 | "读取 A1:D10" | 返回数据 |
| 写入 | "在 A1 写入 Hello" | A1 显示 Hello |
| 公式 | "A10 = SUM(A1:A9)" | 公式正确 |
| 格式 | "A1 加粗" | A1 变粗体 |
| 澄清 | "删除没用的" | 应该先澄清 |

## Mock 后端

用于离线测试，不消耗 API 额度：

```bash
# 启动 Mock 后端 (3002)
node mock-backend.cjs
```
