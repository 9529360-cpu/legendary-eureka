# 常见问题排查

## 启动问题

### 端口被占用

**症状**: `EADDRINUSE: address already in use :::3000`

**解决**:
```powershell
.\scripts\stop_port_3000.ps1
# 或手动
netstat -ano | findstr :3000
taskkill /PID <PID> /F
```

### Add-in 不加载

**症状**: Excel 中看不到 Add-in 面板

**解决**:
1. 确保 dev-server 在运行 (`npm run dev-server`)
2. 清除缓存: `.\scripts\clear_wef.ps1`
3. 重启 Excel
4. 以管理员身份运行: `npm run start`

### 证书问题

**症状**: `CERT_UNTRUSTED` 或 HTTPS 警告

**解决**:
```bash
npx office-addin-dev-certs install
```

---

## 运行时问题

### "依赖不存在的步骤"

**症状**: 
```
操作计划检查发现问题:
步骤 "xxx" 依赖不存在的步骤 excel_create_sheet
```

**根因**: 架构设计问题，dependsOn 使用工具名而非步骤ID

**临时修复**: v3.1.1 已修复此 bug

**根本解决**: 需要重构架构，见 ARCHITECTURE.md

### LLM 返回无效计划

**症状**: 计划验证失败，工具参数缺失

**根因**: LLM 不懂 Excel API 约束

**根本解决**: 需要添加 SpecCompiler 层，见 ARCHITECTURE.md

### Token 消耗过高

**根因**: System Prompt 包含 75 个工具描述

**根本解决**: 需要重构，LLM 只做意图理解

---

## 构建问题

### 构建失败

```bash
# 清理重建
rm -rf dist node_modules
npm install
npm run build:dev
```

### Lint 错误过多

```bash
# 自动修复
npm run lint:fix
```

---

## 调试技巧

### 查看 Agent 日志

在 Excel 按 F12，Console 中查看 `[Agent]` 开头的日志。

### 查看 LLM 请求

AI 后端日志会输出请求和响应。

### 手动测试 API

```powershell
# 健康检查
Invoke-RestMethod -Uri "http://localhost:3001/api/health"

# 聊天测试
$body = @{message="测试"} | ConvertTo-Json
Invoke-RestMethod -Uri "http://localhost:3001/chat" -Method Post -ContentType "application/json" -Body $body
```

---

## 已知问题

| 问题 | 状态 | 版本 |
|------|------|------|
| dependsOn 用工具名做ID | 临时修复 | v3.1.1 |
| LLM 生成计划常失败 | 未解决 | - |
| Token 消耗过高 | 未解决 | - |
| Agent 变成傻执行器 | 未解决 | - |

详见 [ARCHITECTURE.md](../ARCHITECTURE.md) 问题记录表
