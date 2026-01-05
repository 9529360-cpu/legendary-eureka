# Excel 智能助手 Add-in

AI 驱动的 Excel Office Add-in，将自然语言指令转换为 Excel 操作。

## 快速开始

```bash
npm install
npm run dev:full    # 启动所有服务
npm run start       # 在 Excel 中加载
```

## 文档

| 文档 | 说明 |
|------|------|
| [ARCHITECTURE.md](ARCHITECTURE.md) | 系统架构设计 |
| [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) | 开发指南 |
| [docs/API.md](docs/API.md) | API 文档 |
| [docs/TESTING.md](docs/TESTING.md) | 测试说明 |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 问题排查 |
| [CHANGELOG.md](CHANGELOG.md) | 版本记录 |

## 架构概览

```
用户输入 → IntentParser(LLM) → SpecCompiler(规则) → AgentExecutor → Tools → Excel
```

详见 [ARCHITECTURE.md](ARCHITECTURE.md)

## 目录结构

```
├── src/
│   ├── agent/           # Agent 核心
│   ├── taskpane/        # UI 层
│   └── services/        # API 服务
├── docs/                # 文档
├── tests/               # 测试
├── ai-backend.cjs       # AI 后端
└── manifest.xml         # Office Add-in 清单
```

## 技术栈

- **前端**: React 18, Fluent UI 9, TypeScript
- **后端**: Express, DeepSeek API
- **构建**: Webpack 5

## 版本

当前版本: v3.1.1

## License

MIT


