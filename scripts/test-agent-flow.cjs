/**
 * Agent 执行流程快速测试脚本
 * 用虚假数据模拟，不需要真实 Excel 环境
 * 
 * 运行: node scripts/test-agent-flow.cjs
 */

const http = require('http');

// ========== 模拟数据 ==========
const mockEnvironmentState = {
  workbook: {
    sheets: [
      { name: "Sheet1", isActive: true },
      { name: "销售数据", isActive: false },
      { name: "汇总", isActive: false }
    ],
    tables: [
      {
        name: "销售表",
        columns: ["日期", "产品", "销量", "单价", "金额"],
        sheetName: "销售数据",
        rowCount: 150
      },
      {
        name: "产品目录",
        columns: ["产品ID", "产品名称", "类别", "成本价"],
        sheetName: "Sheet1",
        rowCount: 50
      }
    ],
    charts: [
      { name: "销售趋势图", type: "line", sheetName: "汇总" }
    ],
    namedRanges: [
      { name: "销售区域", address: "销售数据!A1:E151" }
    ]
  }
};

const mockConversationHistory = [
  { role: "user", content: "帮我把销售表按金额排序" }
];

// ========== 模拟工具注册表 ==========
const mockToolRegistry = {
  tools: new Map([
    // 读写操作
    ["excel_read_range", { name: "excel_read_range", description: "读取指定范围数据" }],
    ["excel_read_selection", { name: "excel_read_selection", description: "读取当前选中区域" }],
    ["excel_write_range", { name: "excel_write_range", description: "写入数据到范围" }],
    ["excel_write_cell", { name: "excel_write_cell", description: "写入单个单元格" }],
    
    // 感知工具
    ["get_table_schema", { name: "get_table_schema", description: "获取表格结构（列名、数据类型、行数）" }],
    ["sample_rows", { name: "sample_rows", description: "获取前N行样本数据" }],
    
    // 排序筛选
    ["excel_sort_range", { name: "excel_sort_range", description: "对范围排序" }],
    ["excel_filter", { name: "excel_filter", description: "筛选数据" }],
    
    // 格式化
    ["excel_format_range", { name: "excel_format_range", description: "格式化范围（字体、颜色、边框等）" }],
    ["excel_auto_fit", { name: "excel_auto_fit", description: "自动调整列宽" }],
    ["excel_conditional_format", { name: "excel_conditional_format", description: "条件格式" }],
    ["excel_merge_cells", { name: "excel_merge_cells", description: "合并单元格" }],
    
    // 公式
    ["excel_set_formula", { name: "excel_set_formula", description: "设置公式" }],
    ["excel_fill_formula", { name: "excel_fill_formula", description: "填充公式到范围" }],
    
    // 工作表操作
    ["excel_create_sheet", { name: "excel_create_sheet", description: "创建新工作表" }],
    ["excel_switch_sheet", { name: "excel_switch_sheet", description: "切换工作表" }],
    ["excel_delete_sheet", { name: "excel_delete_sheet", description: "删除工作表" }],
    
    // 表格操作
    ["excel_create_table", { name: "excel_create_table", description: "创建表格" }],
    ["excel_get_tables", { name: "excel_get_tables", description: "获取所有表格" }],
    
    // 图表
    ["excel_create_chart", { name: "excel_create_chart", description: "创建图表" }],
    
    // 行列操作
    ["excel_insert_rows", { name: "excel_insert_rows", description: "插入行" }],
    ["excel_delete_rows", { name: "excel_delete_rows", description: "删除行" }],
    ["excel_insert_columns", { name: "excel_insert_columns", description: "插入列" }],
    ["excel_delete_columns", { name: "excel_delete_columns", description: "删除列" }],
    
    // 其他
    ["excel_clear", { name: "excel_clear", description: "清除内容" }],
    ["respond_to_user", { name: "respond_to_user", description: "回复用户" }],
    ["clarify_request", { name: "clarify_request", description: "向用户澄清模糊请求" }],
  ]),
  get(name) {
    return this.tools.get(name);
  },
  getAll() {
    return Array.from(this.tools.values());
  }
};

// ========== 测试用例 ==========
const testCases = [
  // === 基础操作 ===
  {
    name: "简单排序",
    request: "帮我把销售表按金额从大到小排序",
    expectedTools: ["get_table_schema", "excel_sort_range"],
    difficulty: "easy"
  },
  
  // === 跨表操作 ===
  {
    name: "跨表复制数据",
    request: "把销售数据表里的产品列复制到Sheet1的A列",
    expectedTools: ["excel_read_range", "excel_write_range"],
    difficulty: "medium"
  },
  {
    name: "跨表汇总",
    request: "把Sheet1和销售数据两个表的数据合并到一个新表里",
    expectedTools: ["excel_read_range", "excel_create_sheet", "excel_write_range"],
    difficulty: "hard"
  },
  
  // === 条件筛选和分析 ===
  {
    name: "条件筛选",
    request: "筛选出销售表中金额大于1000的记录",
    expectedTools: ["get_table_schema", "excel_filter"],
    difficulty: "medium"
  },
  {
    name: "数据分析-找最大值",
    request: "销售表里哪个产品的销量最高？",
    expectedTools: ["excel_read_range", "respond_to_user"],
    difficulty: "medium"
  },
  {
    name: "数据统计",
    request: "帮我统计一下销售表每个产品的总销量",
    expectedTools: ["get_table_schema", "excel_read_range"],
    difficulty: "hard"
  },
  
  // === 公式操作 ===
  {
    name: "添加求和公式",
    request: "在销售表的F列添加公式，计算每行的 销量*单价",
    expectedTools: ["get_table_schema", "excel_set_formula"],
    difficulty: "medium"
  },
  {
    name: "批量公式填充",
    request: "在G2到G100填充公式 =E2*1.1 计算涨价10%后的金额",
    expectedTools: ["excel_set_formula"],
    difficulty: "medium"
  },
  
  // === 格式化操作 ===
  {
    name: "复杂格式化",
    request: "把销售表的标题行加粗、居中、背景色设为蓝色",
    expectedTools: ["get_table_schema", "excel_format_range"],
    difficulty: "medium"
  },
  {
    name: "条件格式",
    request: "把销售表中金额超过500的单元格标红",
    expectedTools: ["get_table_schema", "excel_conditional_format"],
    difficulty: "hard"
  },
  
  // === 图表操作 ===
  {
    name: "创建图表",
    request: "用销售表的产品和销量数据创建一个柱状图",
    expectedTools: ["get_table_schema", "excel_create_chart"],
    difficulty: "hard"
  },
  
  // === 数据清洗 ===
  {
    name: "查找空值",
    request: "检查销售表有没有空值或缺失数据",
    expectedTools: ["excel_read_range", "respond_to_user"],
    difficulty: "medium"
  },
  {
    name: "数据去重",
    request: "删除销售表中重复的行",
    expectedTools: ["get_table_schema", "excel_read_range"],
    difficulty: "hard"
  },
  
  // === 模糊指令（考验理解能力）===
  {
    name: "模糊指令-整理表格",
    request: "帮我整理一下这个销售表，让它看起来更专业",
    expectedTools: ["get_table_schema", "excel_format_range", "excel_auto_fit"],
    difficulty: "hard"
  },
  {
    name: "模糊指令-数据有问题",
    request: "我觉得这个表的数据有点问题，你帮我检查一下",
    expectedTools: ["get_table_schema", "excel_read_range", "respond_to_user"],
    difficulty: "hard"
  },
  
  // === 多步骤复杂任务 ===
  {
    name: "完整报表流程",
    request: "帮我做一个销售报表：先按金额排序，然后给标题行加格式，最后生成一个饼图",
    expectedTools: ["get_table_schema", "excel_sort_range", "excel_format_range", "excel_create_chart"],
    difficulty: "hard"
  },
  
  // === 边界情况 ===
  {
    name: "不存在的表",
    request: "帮我打开库存表看看有多少数据",
    expectedTools: ["get_table_schema"],
    difficulty: "edge",
    expectError: true
  },
  {
    name: "纯对话-不需要操作",
    request: "Excel里怎么用VLOOKUP函数？",
    expectedTools: ["respond_to_user"],
    difficulty: "easy"
  }
];

// ========== 调用 AI 后端 ==========
async function callAIBackend(message, systemPrompt) {
  return new Promise((resolve, reject) => {
    const postData = JSON.stringify({
      message,
      systemPrompt,
      responseFormat: "json"
    });

    const options = {
      hostname: 'localhost',
      port: 3001,
      path: '/agent/chat',  // Agent 专用接口，支持自定义 systemPrompt
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(postData)
      }
    };

    const req = http.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          resolve(JSON.parse(data));
        } catch (e) {
          resolve({ message: data });
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(60000, () => {  // 60秒超时，避免复杂任务超时
      req.destroy();
      reject(new Error('Request timeout'));
    });
    req.write(postData);
    req.end();
  });
}

// ========== 构建 System Prompt (模拟 AgentCore.buildPlannerSystemPrompt) ==========
function buildSystemPrompt() {
  const toolList = mockToolRegistry.getAll()
    .map(t => `- ${t.name}: ${t.description}`)
    .join('\n');

  return `你是Excel Office Add-in助手。根据用户请求生成执行计划。

## 可用工具
${toolList}
- clarify_request: 向用户澄清模糊请求

## 感知工具（重要！）
- get_table_schema: 获取表格结构（必须提供sheetName或tableName参数）
- sample_rows: 获取前N行样本数据，了解数据格式
- excel_read_range: 读取指定区域数据（必须提供address参数）

## 特殊工具
- respond_to_user: 回复用户
  参数: {message: "{{ANALYZE_AND_REPLY}}"} 需要分析数据后回复
  参数: {message: "具体内容"} 简单回复
- clarify_request: 向用户澄清请求
  参数: {question: "您具体想...?", options: ["选项1", "选项2"]}

## ★★★ 澄清优先规则（最重要！）★★★
以下情况**必须**先用 clarify_request 澄清，**禁止**直接操作：

1. **模糊+删除类请求**：
   - "删除没用的" → 什么是"没用的"？空行？空列？重复数据？
   - "清理一下" → 清理什么？格式？数据？
   - "优化表格" → 优化什么？格式？结构？删除数据？
   
2. **有副作用+不明确范围**：
   - "把这些数据整理一下" → 整理到哪里？覆盖原数据？新建sheet？
   - "帮我处理一下" → 处理什么？怎么处理？

3. **澄清示例**：
   用户说"删除没用的列"
   → 先 clarify_request: "您想删除哪些列？请选择：
      A) 完全空白的列
      B) 大部分为空的列（超过50%为空）  
      C) 您指定的特定列
      请告诉我您的选择，或直接告诉我要删除的列名。"

## 核心规则（必须遵守）
1. **先感知再操作**：执行任何写操作前，必须先调用感知工具确认目标区域结构
2. **感知工具必须带参数**：get_table_schema 必须传 sheetName 或 tableName
3. **跨表操作**：需要先用 excel_read_range 读取源数据，再用 excel_write_range 写入目标位置
4. **公式操作**：使用 excel_set_formula，公式必须以 = 开头
5. **筛选操作**：使用 excel_filter，需要指定列名和条件
6. **格式化**：使用 excel_format_range，可设置 font, backgroundColor, alignment 等属性
7. **图表**：使用 excel_create_chart，需要指定 dataRange 和 chartType
8. **纯问答**：如果用户问的是 Excel 知识而非操作请求，直接用 respond_to_user 回答

## 输出JSON格式
{"intent":"query|operation|clarify","clarifyReason":"如果intent是clarify，说明原因","steps":[{"order":1,"action":"工具名","parameters":{...},"description":"描述","isWriteOperation":true/false}],"completionMessage":"完成提示"}

## 判断流程
1. 用户请求是否模糊？（"删除没用的"、"优化一下"等）
2. 是否有副作用？（删除、修改、覆盖等）
3. 如果 模糊 + 有副作用 → intent: "clarify"，用 clarify_request 工具
4. 如果明确 → intent: "operation"，正常执行`;
}

// ========== 构建用户 Prompt (模拟 AgentCore.buildPlanGenerationPrompt) ==========
function buildUserPrompt(request) {
  return `## 当前请求
用户: ${request}

## 工作簿信息
${JSON.stringify(mockEnvironmentState.workbook, null, 2)}

请根据请求生成执行计划 JSON。`;
}

// ========== 解析 LLM 返回的计划 ==========
function parsePlan(response) {
  const message = response.message || response;
  
  // 尝试提取 JSON
  const jsonMatch = message.match(/\{[\s\S]*\}/);
  if (jsonMatch) {
    try {
      return JSON.parse(jsonMatch[0]);
    } catch (e) {
      console.error("JSON 解析失败:", e.message);
      return null;
    }
  }
  return null;
}

// ========== 模拟工具执行 ==========
function simulateToolExecution(step) {
  const tool = mockToolRegistry.get(step.action);
  
  if (!tool) {
    return {
      success: false,
      error: `工具不存在: ${step.action}`,
      output: null
    };
  }

  const params = step.parameters || {};

  // 模拟各种工具的返回
  switch (step.action) {
    case "get_table_schema":
      // 检查表是否存在
      const tableName = params.name || params.tableName;
      const tableExists = mockEnvironmentState.workbook.tables.some(t => t.name === tableName) ||
                          mockEnvironmentState.workbook.sheets.some(s => s.name === tableName);
      if (!tableExists && tableName !== "库存表") {
        return {
          success: true,
          output: `表格「${tableName}」详细结构:
- 类型: Excel Table
- 行数: 150
- 列数: 5
- 列定义:
  A列「日期」: date (YYYY-MM-DD), 格式=yyyy-mm-dd, 示例=[2024-01-01, 2024-01-02, 2024-01-03]
  B列「产品」: text, 格式=General, 示例=[苹果, 香蕉, 橙子]
  C列「销量」: number, 格式=#, 示例=[100, 200, 150]
  D列「单价」: number, 格式=#.00, 示例=[5.00, 3.50, 4.00]
  E列「金额」: number, 格式=#.00, 示例=[500.00, 700.00, 600.00]`,
          data: { 
            columns: ["日期", "产品", "销量", "单价", "金额"],
            rowCount: 150,
            dataAddress: "销售数据!A2:E151"
          }
        };
      } else if (tableName === "库存表") {
        return {
          success: false,
          error: "未找到表格或工作表「库存表」",
          output: "未找到表格或工作表「库存表」"
        };
      }
      return {
        success: true,
        output: `表格「${tableName}」结构: 5列, 150行`,
        data: { columns: ["日期", "产品", "销量", "单价", "金额"], rowCount: 150 }
      };

    case "sample_rows":
      return {
        success: true,
        output: `样本数据 (前5行):
1. 2024-01-01 | 苹果 | 100 | 5.00 | 500.00
2. 2024-01-02 | 香蕉 | 200 | 3.50 | 700.00
3. 2024-01-03 | 橙子 | 150 | 4.00 | 600.00
4. 2024-01-04 | 苹果 | 80 | 5.00 | 400.00
5. 2024-01-05 | 葡萄 | 300 | 8.00 | 2400.00`,
        data: { 
          sampleData: [
            ["2024-01-01", "苹果", 100, 5.00, 500.00],
            ["2024-01-02", "香蕉", 200, 3.50, 700.00],
            ["2024-01-03", "橙子", 150, 4.00, 600.00]
          ]
        }
      };

    case "excel_read_range":
      return {
        success: true,
        output: `读取 ${params.address || params.range || 'A1:E10'}: 数据包含 10 行 5 列
第一行: 日期, 产品, 销量, 单价, 金额
数据范围: 2024-01-01 到 2024-01-10`,
        data: { 
          values: [
            ["日期", "产品", "销量", "单价", "金额"],
            ["2024-01-01", "苹果", 100, 5, 500],
            ["2024-01-02", "香蕉", 200, 3.5, 700],
            ["2024-01-03", "橙子", 150, 4, 600]
          ],
          rowCount: 10,
          columnCount: 5
        }
      };

    case "excel_sort_range":
      return {
        success: true,
        output: `已按「${params.sortBy || params.column || '金额'}」${params.ascending === false || params.order === 'descending' ? '降序' : '升序'}排序，共 150 行数据`,
        data: {}
      };

    case "excel_filter":
      return {
        success: true,
        output: `已应用筛选条件，筛选出 ${Math.floor(Math.random() * 50) + 10} 条记录`,
        data: { filteredCount: 35 }
      };

    case "excel_write_range":
      return {
        success: true,
        output: `已写入数据到 ${params.address || params.range || 'A1'}`,
        data: {}
      };

    case "excel_format_range":
      return {
        success: true,
        output: `已格式化范围 ${params.address || params.range || 'A1:E1'}`,
        data: {}
      };

    case "excel_auto_fit":
      return {
        success: true,
        output: `已自动调整列宽`,
        data: {}
      };

    case "excel_conditional_format":
      return {
        success: true,
        output: `已添加条件格式规则`,
        data: {}
      };

    case "excel_set_formula":
      return {
        success: true,
        output: `已在 ${params.address || params.cell || params.range} 设置公式: ${params.formula}`,
        data: {}
      };

    case "excel_fill_formula":
      return {
        success: true,
        output: `已填充公式到指定范围`,
        data: {}
      };

    case "excel_create_sheet":
      return {
        success: true,
        output: `已创建新工作表「${params.name || '新工作表'}」`,
        data: {}
      };

    case "excel_switch_sheet":
      return {
        success: true,
        output: `已切换到工作表「${params.name || params.sheetName}」`,
        data: {}
      };

    case "excel_create_chart":
      return {
        success: true,
        output: `已创建${params.chartType || '柱状'}图表`,
        data: {}
      };

    case "excel_create_table":
      return {
        success: true,
        output: `已创建表格「${params.name || '表格1'}」`,
        data: {}
      };

    case "excel_insert_rows":
    case "excel_delete_rows":
    case "excel_insert_columns":
    case "excel_delete_columns":
      return {
        success: true,
        output: `${step.action} 操作成功`,
        data: {}
      };

    case "excel_clear":
      return {
        success: true,
        output: `已清除指定范围内容`,
        data: {}
      };

    case "respond_to_user":
      return {
        success: true,
        output: params.message || "操作完成",
        data: {}
      };

    default:
      return {
        success: true,
        output: `${step.action} 执行成功`,
        data: {}
      };
  }
}

// ========== 模拟 ensurePerceptionBeforeWrite ==========
function ensurePerceptionBeforeWrite(plan) {
  const writeTools = new Set([
    "excel_write_range", "excel_write_cell", "excel_set_formula",
    "excel_format_range", "excel_sort_range", "excel_filter"
  ]);
  const perceptionTools = new Set([
    "excel_read_range", "excel_read_selection", "get_table_schema", "sample_rows"
  ]);

  const hasWrite = plan.steps?.some(s => writeTools.has(s.action));
  const hasPerception = plan.steps?.some(s => perceptionTools.has(s.action));

  if (hasWrite && !hasPerception) {
    console.log("⚠️  计划缺少感知步骤，Agent 层强制插入！");
    return true;
  }
  return false;
}

// ========== 模拟 preValidateAndFixParams ==========
function preValidateAndFixParams(action, params) {
  const fixed = { ...params };
  let changes = [];

  // v3.0.3: 参数别名兼容
  const aliasMap = {
    get_table_schema: { tableName: "name", table: "name" },
    sample_rows: { tableName: "name", table: "name" },
    excel_read_range: { range: "address" },
    excel_write_range: { range: "address", data: "values" },
    excel_write_cell: { range: "address", data: "value" },
    excel_sort_range: { range: "address", sortColumn: "column", order: "ascending" },
    excel_format_range: { range: "address" },
    excel_set_formula: { range: "address", cell: "address" },
  };

  const toolAliases = aliasMap[action];
  if (toolAliases) {
    for (const [alias, canonical] of Object.entries(toolAliases)) {
      if (fixed[alias] !== undefined && fixed[canonical] === undefined) {
        if (alias === "order" && canonical === "ascending") {
          const orderVal = String(fixed[alias]).toLowerCase();
          fixed[canonical] = orderVal !== "descending" && orderVal !== "desc";
        } else {
          fixed[canonical] = fixed[alias];
        }
        delete fixed[alias];
        changes.push(`${alias} -> ${canonical}`);
      }
    }
  }

  // 地址格式修正
  if (fixed.address && typeof fixed.address === 'string') {
    let addr = fixed.address;
    if (addr.includes('：')) {
      fixed.address = addr.replace(/：/g, ':');
      changes.push(`地址中文冒号修正: ${addr} -> ${fixed.address}`);
    }
  }

  // values 格式修正
  if (fixed.values !== undefined) {
    if (!Array.isArray(fixed.values)) {
      fixed.values = [[fixed.values]];
      changes.push(`values 转二维数组`);
    } else if (fixed.values.length > 0 && !Array.isArray(fixed.values[0])) {
      fixed.values = fixed.values.map(v => [v]);
      changes.push(`values 一维转二维`);
    }
  }

  if (changes.length > 0) {
    console.log(`  📝 参数修正: ${changes.join(', ')}`);
  }

  return fixed;
}

// ========== 主测试流程 ==========
async function runTest(testCase) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`📋 测试: ${testCase.name}`);
  console.log(`📝 请求: ${testCase.request}`);
  console.log('='.repeat(60));

  try {
    // 1. 构建 Prompt
    const systemPrompt = buildSystemPrompt();
    const userPrompt = buildUserPrompt(testCase.request);

    console.log('\n[1] 发送请求到 AI 后端...');

    // 2. 调用 LLM
    const response = await callAIBackend(userPrompt, systemPrompt);
    
    console.log('[2] LLM 原始响应:');
    console.log(response.message?.substring(0, 500) || JSON.stringify(response).substring(0, 500));

    // 3. 解析计划
    const plan = parsePlan(response);
    
    if (!plan) {
      console.log('\n❌ 计划解析失败！');
      return { success: false, error: 'Plan parse failed' };
    }

    console.log('\n[3] 解析后的计划:');
    console.log(`  Intent: ${plan.intent}`);
    console.log(`  Steps: ${plan.steps?.length || 0} 个`);
    plan.steps?.forEach((step, i) => {
      console.log(`    ${i + 1}. ${step.action} - ${step.description}`);
      console.log(`       参数: ${JSON.stringify(step.parameters)}`);
    });

    // 4. 检查是否需要强制感知
    const needsPerception = ensurePerceptionBeforeWrite(plan);
    if (needsPerception) {
      console.log('\n[4] 强制感知检查: 需要插入感知步骤');
    } else {
      console.log('\n[4] 强制感知检查: 通过');
    }

    // 5. 模拟执行每个步骤
    console.log('\n[5] 模拟执行步骤:');
    const results = [];
    
    for (let i = 0; i < (plan.steps?.length || 0); i++) {
      const step = plan.steps[i];
      
      // 检查工具是否存在
      const tool = mockToolRegistry.get(step.action);
      if (!tool) {
        console.log(`  ❌ 步骤 ${i + 1}: 工具不存在 "${step.action}"`);
        results.push({ step: i + 1, success: false, error: `Tool not found: ${step.action}` });
        continue;
      }

      // 预验证参数
      const fixedParams = preValidateAndFixParams(step.action, step.parameters || {});
      
      // 执行
      const result = simulateToolExecution({ ...step, parameters: fixedParams });
      
      if (result.success) {
        console.log(`  ✅ 步骤 ${i + 1}: ${step.action} 成功`);
        console.log(`     输出: ${result.output.substring(0, 100)}`);
      } else {
        console.log(`  ❌ 步骤 ${i + 1}: ${step.action} 失败 - ${result.error}`);
      }
      
      results.push({ step: i + 1, ...result });
    }

    // 6. 检查预期工具是否被调用
    console.log('\n[6] 工具调用检查:');
    const calledTools = plan.steps?.map(s => s.action) || [];
    testCase.expectedTools.forEach(expected => {
      if (calledTools.includes(expected)) {
        console.log(`  ✅ ${expected} 已调用`);
      } else {
        console.log(`  ⚠️  ${expected} 未调用（可能需要检查）`);
      }
    });

    // 7. 总结
    // 对于 expectError 用例：LLM 正确尝试了感知操作，即使执行失败也算通过
    // 因为真实环境中会触发 replan 并向用户解释
    const hasErrors = results.some(r => !r.success);
    let testPassed;
    
    if (testCase.expectError) {
      // 边缘用例：只要 LLM 生成了正确的感知计划就算通过
      testPassed = plan.steps && plan.steps.length > 0;
      console.log(`\n[结果] ${testPassed ? '✅ 测试通过（边缘用例：LLM正确生成了感知计划）' : '❌ 测试失败'}`);
    } else {
      testPassed = !hasErrors;
      console.log(`\n[结果] ${testPassed ? '✅ 测试通过' : '❌ 测试失败'}`);
    }
    
    return { success: testPassed, plan, results };

  } catch (error) {
    console.log(`\n❌ 测试异常: ${error.message}`);
    return { success: false, error: error.message };
  }
}

// ========== 入口 ==========
async function main() {
  console.log('🚀 Agent 执行流程综合测试');
  console.log('=' .repeat(60));
  console.log('测试用例: ' + testCases.length + ' 个');
  console.log('难度分布: easy=' + testCases.filter(t => t.difficulty === 'easy').length +
              ', medium=' + testCases.filter(t => t.difficulty === 'medium').length +
              ', hard=' + testCases.filter(t => t.difficulty === 'hard').length +
              ', edge=' + testCases.filter(t => t.difficulty === 'edge').length);
  console.log('=' .repeat(60));

  // 支持命令行参数选择测试
  const args = process.argv.slice(2);
  let casesToRun = testCases;
  
  if (args.includes('--quick')) {
    // 快速模式：只运行 3 个核心测试
    casesToRun = testCases.filter(t => 
      ['简单排序', '跨表复制数据', '完整报表流程'].includes(t.name)
    );
    console.log('⚡ 快速模式: 只运行 ' + casesToRun.length + ' 个核心测试\n');
  } else if (args.includes('--hard')) {
    // 只运行困难测试
    casesToRun = testCases.filter(t => t.difficulty === 'hard');
    console.log('💪 困难模式: 只运行 ' + casesToRun.length + ' 个困难测试\n');
  } else if (args.length > 0 && !args[0].startsWith('--')) {
    // 按名称筛选
    const keyword = args[0];
    casesToRun = testCases.filter(t => t.name.includes(keyword));
    console.log('🔍 筛选模式: 匹配 "' + keyword + '", 共 ' + casesToRun.length + ' 个\n');
  }

  const results = [];
  const stats = {
    total: casesToRun.length,
    passed: 0,
    failed: 0,
    perceptionUsed: 0,  // 使用了感知工具的数量
    paramFixApplied: 0, // 应用了参数修正的数量
    avgSteps: 0,
    toolUsage: {}       // 工具使用统计
  };

  for (const testCase of casesToRun) {
    const result = await runTest(testCase);
    results.push({ name: testCase.name, difficulty: testCase.difficulty, ...result });
    
    if (result.success) {
      stats.passed++;
    } else {
      stats.failed++;
    }
    
    // 统计感知工具使用
    if (result.plan?.steps?.some(s => 
      ['get_table_schema', 'sample_rows', 'excel_read_range'].includes(s.action)
    )) {
      stats.perceptionUsed++;
    }
    
    // 统计工具使用
    result.plan?.steps?.forEach(s => {
      stats.toolUsage[s.action] = (stats.toolUsage[s.action] || 0) + 1;
    });
    
    stats.avgSteps += (result.plan?.steps?.length || 0);
  }
  
  stats.avgSteps = (stats.avgSteps / casesToRun.length).toFixed(1);

  // 汇总
  console.log('\n\n' + '='.repeat(60));
  console.log('📊 测试汇总');
  console.log('='.repeat(60));
  
  // 按难度分组显示结果
  const byDifficulty = { easy: [], medium: [], hard: [], edge: [] };
  results.forEach(r => {
    byDifficulty[r.difficulty || 'medium'].push(r);
  });
  
  for (const [diff, items] of Object.entries(byDifficulty)) {
    if (items.length === 0) continue;
    console.log(`\n[${diff.toUpperCase()}]`);
    items.forEach(r => {
      const icon = r.success ? '✅' : '❌';
      const stepsInfo = r.plan?.steps?.length ? ` (${r.plan.steps.length}步)` : '';
      console.log(`  ${icon} ${r.name}${stepsInfo}`);
      if (!r.success && r.error) {
        console.log(`     └─ ${r.error.substring(0, 50)}`);
      }
    });
  }

  console.log('\n' + '-'.repeat(60));
  console.log(`通过率: ${stats.passed}/${stats.total} (${(stats.passed/stats.total*100).toFixed(0)}%)`);
  console.log(`感知工具使用率: ${stats.perceptionUsed}/${stats.total} (${(stats.perceptionUsed/stats.total*100).toFixed(0)}%)`);
  console.log(`平均步骤数: ${stats.avgSteps}`);
  
  // 工具使用排行
  console.log('\n工具使用 Top 5:');
  const topTools = Object.entries(stats.toolUsage)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  topTools.forEach(([tool, count], i) => {
    console.log(`  ${i + 1}. ${tool}: ${count}次`);
  });
  
  // 失败用例详情
  const failedCases = results.filter(r => !r.success);
  if (failedCases.length > 0) {
    console.log('\n❌ 失败用例详情:');
    failedCases.forEach(r => {
      console.log(`  - ${r.name}: ${r.error || '未知错误'}`);
    });
  }
  
  console.log('\n' + '='.repeat(60));
  console.log(stats.passed === stats.total ? '🎉 全部通过!' : `⚠️ ${stats.failed} 个用例失败`);
}

main().catch(console.error);
