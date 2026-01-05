# Excel 智能助手 Add-in

AI 驱动的 Excel Office Add-in，将自然语言指令转换为 Excel 操作。

## ⚠️ 开发须知

**写代码前必读 [.github/copilot-instructions.md](.github/copilot-instructions.md)**

特别注意：
- ❌ **禁止用 PowerShell 重定向写文件**（会产生中文乱码）
- ✅ 用 `create_file` 工具或 Node.js 写文件
- ✅ 改代码后运行 `npm run build:dev` 验证

## 快速开始

```bash
npm install
npm run dev:full    # 启动所有服务
npm run start       # 在 Excel 中加载
```

## 文档索引

| 文档 | 说明 |
|------|------|
| [.github/copilot-instructions.md](.github/copilot-instructions.md) | **AI 开发规范（必读）** |
| [ARCHITECTURE.md](ARCHITECTURE.md) | 系统架构设计（目标架构 v4.0） |
| [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) | 开发指南 |
| [docs/API.md](docs/API.md) | API 文档 |
| [docs/TESTING.md](docs/TESTING.md) | 测试说明 |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 问题排查（编码问题置顶） |
| [CHANGELOG.md](CHANGELOG.md) | 版本记录 |

## 目录结构

```
├── .github/
│   └── copilot-instructions.md  # AI 开发规范
├── src/
│   ├── agent/                   # Agent 核心
│   │   ├── AgentCore.ts
│   │   ├── types/
│   │   ├── workflow/
│   │   ├── registry/
│   │   └── tools/excel/         # 61 个 Excel 工具
│   ├── taskpane/                # UI 层 (React)
│   └── services/                # API 服务
├── docs/                        # 文档
│   ├── archive/                 # 归档旧文档
│   └── learning/                # 学习笔记
├── scripts/                     # 工具脚本
├── tests/                       # 测试
├── ai-backend.cjs               # AI 后端
└── manifest.xml                 # Office Add-in 清单
```

## 技术栈

- **前端**: React 18, Fluent UI 9, TypeScript
- **后端**: Express, DeepSeek API
- **构建**: Webpack 5
- **测试**: Jest, React Testing Library

## 常用脚本

| 脚本 | 用途 |
|------|------|
| `scripts/clear_wef.ps1` | 清除 Office 缓存 |
| `scripts/stop_port_3000.ps1` | 释放端口 |
| `scripts/fix_encoding.js` | 修复 App.tsx 乱码 |
| `scripts/clean_encoding.cjs` | 清理文件不可见字符 |

## 版本

当前版本: **v3.1.1**

## License

MIT


