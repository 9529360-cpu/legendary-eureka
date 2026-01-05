# 常见问题排查

---

## ⚠️ 最常见问题：编码/乱码

> **这是本项目最频繁出现的问题，必须首先检查！**

### 症状

- 中文字符显示为乱码（如 `浣犲ソ` 或 `???`）
- 文件内容出现不可见字符
- Git diff 显示大量莫名其妙的改动
- README.md 或其他 MD 文件内容损坏

### 根本原因

1. **PowerShell 默认编码不是 UTF-8** - 使用 `Out-File` 或重定向时会产生乱码
2. **编辑器保存编码不一致** - 某些工具默认使用 GBK 或 UTF-16
3. **Git 自动转换** - 可能在提交时转换换行符或编码

### 预防措施 ⭐

```powershell
# 在 PowerShell 中设置 UTF-8 编码（每次会话前运行）
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$PSDefaultParameterValues['Set-Content:Encoding'] = 'utf8'

# 或添加到 PowerShell 配置文件
notepad $PROFILE
```

### 修复脚本

项目提供了多个编码修复脚本：

| 脚本 | 用途 |
|---|---|
| `scripts/fix_encoding.js` | 修复 App.tsx 中的乱码 |
| `scripts/fix_encoding.ps1` | PowerShell 版本的编码修复 |
| `scripts/fix_app_encoding.ps1` | 专门修复 App.tsx |
| `scripts/clean_encoding.cjs` | 清理文件中的不可见字符 |

```powershell
# 修复特定文件
node scripts/clean_encoding.cjs src/taskpane/components/App.tsx

# 修复 App.tsx
node scripts/fix_encoding.js
```

### 创建新文件的正确方式

```powershell
# ✅ 正确：使用 Node.js 写文件
node -e "require('fs').writeFileSync('file.md', '内容', 'utf8')"

# ✅ 正确：使用 VS Code 创建
# 直接在 VS Code 中新建文件

# ❌ 错误：PowerShell 重定向
echo "内容" > file.md  # 可能产生乱码！

# ❌ 错误：Set-Content 不指定编码
Set-Content -Path file.md -Value "内容"  # 可能产生乱码！
```

### 验证文件编码

```powershell
# 检查文件头部字节（UTF-8 BOM: EF BB BF）
Format-Hex -Path file.md | Select-Object -First 1

# 用 Node.js 验证
node -e "console.log(require('fs').readFileSync('file.md', 'utf8').substring(0, 100))"
```

---

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

## 已知问题（待解决）

| 问题 | 影响 | 状态 | 记录版本 |
|------|------|------|----------|
| LLM 生成计划常失败 | 用户体验差 | 设计问题，待架构重构 | v3.1 |
| Token 消耗过高 | 成本高 | System Prompt 含 75 工具，待重构 | v3.1 |
| Agent 变成傻执行器 | 不够智能 | 缺少 SpecCompiler 层，待重构 | v3.1 |

详见 [ARCHITECTURE.md](../ARCHITECTURE.md) 问题记录表

---

## 已修复 BUG 历史

| 问题 | 根因 | 修复版本 | 修复方式 |
|------|------|----------|----------|
| "依赖不存在的步骤 excel_xxx" | `dependsOn` 使用工具名而非步骤ID | v3.1.1 | 改用两遍处理，先生成 stepId |
| App.tsx 中文乱码 | PowerShell 编码不是 UTF-8 | 多次修复 | 添加 `scripts/fix_encoding.js` |
| README.md 损坏 | 编辑器/终端编码不一致 | v3.1.1 | 删除重写 |

---

## 故障排查清单

遇到问题时按以下顺序检查：

1. **编码问题** - 文件是否有乱码？运行 `node scripts/clean_encoding.cjs <file>`
2. **缓存问题** - Add-in 没更新？运行 `.\scripts\clear_wef.ps1` 并重启 Excel
3. **端口问题** - 服务启动失败？运行 `.\scripts\stop_port_3000.ps1`
4. **后端问题** - API 返回错误？检查 `ai-backend.cjs` 日志
5. **LLM 问题** - 计划生成失败？查看 ARCHITECTURE.md 已知问题
