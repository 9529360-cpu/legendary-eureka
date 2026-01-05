/**
 * Agent 暴力压力测试
 * 用极端场景测试助手的理解和执行能力
 * 
 * 运行: node scripts/stress-test-agent.cjs
 */

const http = require('http');

// ========== 暴力测试用例 ==========
const stressTestCases = [
  // ========== 1. 复杂公式生成 ==========
  {
    name: "复杂嵌套公式",
    request: "帮我写一个公式：如果A列是'已完成'且B列大于100，就显示C列乘以1.1，否则如果B列小于50就显示'低'，其他情况显示C列原值",
    category: "formula",
    expectedCapabilities: ["理解多层IF嵌套", "生成正确公式语法"],
  },
  {
    name: "VLOOKUP跨表查询",
    request: "用VLOOKUP从产品目录表查找销售表里每个产品的成本价，然后计算利润率",
    category: "formula",
    expectedCapabilities: ["跨表引用", "公式组合"],
  },
  {
    name: "动态数组公式",
    request: "用UNIQUE函数提取所有不重复的产品名，然后用SUMIF统计每个产品的总销量",
    category: "formula",
    expectedCapabilities: ["现代Excel函数", "公式联动"],
  },
  {
    name: "日期计算公式",
    request: "计算每笔订单距今多少天，超过30天的标记为'逾期'，7天内的标记为'新订单'",
    category: "formula",
    expectedCapabilities: ["日期函数", "条件判断"],
  },
  
  // ========== 2. 数据分析洞察 ==========
  {
    name: "趋势分析",
    request: "分析销售数据的趋势，告诉我哪个月份销量最好，哪个产品增长最快",
    category: "analysis",
    expectedCapabilities: ["数据聚合", "趋势识别", "自然语言回复"],
  },
  {
    name: "异常值检测",
    request: "帮我找出销售表里的异常数据，比如金额特别大或特别小的记录",
    category: "analysis",
    expectedCapabilities: ["统计分析", "异常识别"],
  },
  {
    name: "数据质量检查",
    request: "全面检查这个表的数据质量：有没有空值、重复、格式错误、数据类型不一致的问题",
    category: "analysis",
    expectedCapabilities: ["多维度检查", "问题汇总"],
  },
  {
    name: "对比分析",
    request: "对比今年和去年同期的销售数据，算出增长率",
    category: "analysis",
    expectedCapabilities: ["时间对比", "增长计算"],
    note: "可能缺少去年数据"
  },
  
  // ========== 3. 模糊/不完整指令 ==========
  {
    name: "极度模糊-处理数据",
    request: "处理一下这个数据",
    category: "fuzzy",
    expectedCapabilities: ["澄清意图", "合理假设"],
  },
  {
    name: "模糊-让表格好看",
    request: "让这个表好看点",
    category: "fuzzy",
    expectedCapabilities: ["理解美化意图", "应用格式"],
  },
  {
    name: "口语化指令",
    request: "把那个啥，就是卖得最多的那几个产品给我挑出来放到新表里",
    category: "fuzzy",
    expectedCapabilities: ["理解口语", "推断意图"],
  },
  {
    name: "不完整指令-排序",
    request: "排个序",
    category: "fuzzy",
    expectedCapabilities: ["询问排序依据", "或合理默认"],
  },
  {
    name: "指代不明",
    request: "把它复制到那边去",
    category: "fuzzy",
    expectedCapabilities: ["识别指代不明", "请求澄清"],
  },
  {
    name: "错误表名",
    request: "打开销售汇总表",
    category: "fuzzy",
    note: "表名不存在，应该是'销售表'",
    expectedCapabilities: ["模糊匹配", "建议正确名称"],
  },
  
  // ========== 4. 多步骤复杂任务 ==========
  {
    name: "完整数据清洗流程",
    request: "帮我清洗这个表：去掉重复行，填充空值用0，把日期格式统一成YYYY-MM-DD，金额保留两位小数",
    category: "complex",
    expectedCapabilities: ["多步骤规划", "顺序执行"],
  },
  {
    name: "完整报表生成",
    request: "生成一份月度销售报表：按月汇总销量和金额，计算环比增长率，生成趋势图，并在表头加上'2024年销售月报'的标题",
    category: "complex",
    expectedCapabilities: ["数据聚合", "计算", "图表", "格式化"],
  },
  {
    name: "数据透视表模拟",
    request: "按产品和月份做一个交叉表，显示每个产品每个月的销量汇总",
    category: "complex",
    expectedCapabilities: ["透视逻辑", "多维聚合"],
  },
  {
    name: "条件批量操作",
    request: "找出所有金额超过1000的订单，把它们的行背景标黄，在旁边加一列备注写'大额订单'，最后统计一共有多少笔",
    category: "complex",
    expectedCapabilities: ["条件筛选", "批量格式", "批量写入", "统计"],
  },
  
  // ========== 5. 跨表复杂操作 ==========
  {
    name: "多表关联查询",
    request: "把销售表、产品目录和库存表三个表的数据合并，显示每个产品的销量、成本和库存",
    category: "cross-table",
    expectedCapabilities: ["多表关联", "数据合并"],
  },
  {
    name: "跨表计算",
    request: "根据产品目录的成本价计算销售表每笔订单的利润，利润=金额-销量*成本价",
    category: "cross-table",
    expectedCapabilities: ["跨表引用", "公式计算"],
  },
  {
    name: "跨表数据同步",
    request: "把销售数据表的产品列去重后更新到产品目录表，如果有新产品就添加",
    category: "cross-table",
    expectedCapabilities: ["去重", "跨表写入", "增量更新"],
  },
  
  // ========== 6. 边缘情况 ==========
  {
    name: "空表操作",
    request: "在汇总表里创建一个表头：日期、产品、销量、单价、金额，然后从销售表复制前10行数据过去",
    category: "edge",
    expectedCapabilities: ["空表处理", "表头创建"],
  },
  {
    name: "特殊字符处理",
    request: "搜索产品名称包含'/'或'&'的记录",
    category: "edge",
    expectedCapabilities: ["特殊字符", "正则匹配"],
  },
  {
    name: "大范围操作",
    request: "给A1到Z1000的所有单元格加上边框",
    category: "edge",
    expectedCapabilities: ["大范围", "性能考虑"],
  },
  {
    name: "负数/零值处理",
    request: "找出所有销量为0或负数的记录，标记为异常",
    category: "edge",
    expectedCapabilities: ["边界值", "条件标记"],
  },
  
  // ========== 7. 上下文理解 ==========
  {
    name: "指代上一步结果",
    request: "把刚才排序的结果导出成新表",
    category: "context",
    expectedCapabilities: ["理解'刚才'", "结果引用"],
    note: "需要对话历史"
  },
  {
    name: "修改上一步",
    request: "不对，我要的是降序不是升序",
    category: "context",
    expectedCapabilities: ["理解纠正", "撤销重做"],
    note: "需要对话历史"
  },
  {
    name: "追问细节",
    request: "为什么那个公式算出来是这个结果？",
    category: "context",
    expectedCapabilities: ["解释计算", "公式分析"],
    note: "需要知道之前的公式"
  },
  
  // ========== 8. 专业领域 ==========
  {
    name: "财务计算",
    request: "计算每个产品的毛利率和净利率，毛利率=(销售额-成本)/销售额",
    category: "domain",
    expectedCapabilities: ["财务公式", "准确计算"],
  },
  {
    name: "统计分析",
    request: "计算销售数据的平均值、中位数、标准差，并判断数据分布是否正态",
    category: "domain",
    expectedCapabilities: ["统计函数", "分布分析"],
  },
  {
    name: "时间序列",
    request: "按周汇总销售数据，计算周环比，找出销售高峰周",
    category: "domain",
    expectedCapabilities: ["时间聚合", "环比计算"],
  },
  
  // ========== 9. 错误场景 ==========
  {
    name: "无效范围",
    request: "读取Z999:AA1000的数据",
    category: "error",
    expectedCapabilities: ["范围验证", "错误提示"],
    note: "可能超出数据范围"
  },
  {
    name: "类型不匹配",
    request: "把日期列求和",
    category: "error",
    expectedCapabilities: ["类型检查", "合理处理或提示"],
  },
  {
    name: "循环引用风险",
    request: "在A1写一个公式引用B1，在B1写一个公式引用A1",
    category: "error",
    expectedCapabilities: ["循环检测", "警告用户"],
  },
  
  // ========== 10. 极限压力 ==========
  {
    name: "超长指令",
    request: "首先切换到销售数据表，然后获取表结构，接着按日期升序排序，之后按金额降序再排一次，然后给标题行加粗加背景色蓝色字体白色居中对齐，给数据区域加边框，把金额列格式化成货币格式保留两位小数，把日期列格式化成YYYY年MM月DD日格式，然后筛选出金额大于500的记录，给这些记录的行背景标黄，接着在F列添加一个公式计算每行的利润率=(金额-销量*5)/金额，然后生成一个柱状图显示每个产品的总销量，最后生成一个饼图显示各产品销售占比，把图表放在G列开始的位置",
    category: "stress",
    expectedCapabilities: ["长指令解析", "多步骤拆分"],
  },
  {
    name: "矛盾指令",
    request: "按金额从大到小排序，同时按日期从小到大排序",
    category: "stress",
    expectedCapabilities: ["识别矛盾", "请求澄清优先级"],
  },
  {
    name: "不可能任务",
    request: "预测下个月的销量会是多少",
    category: "stress",
    expectedCapabilities: ["识别能力边界", "诚实回应"],
  },
];

// ========== 模拟环境 ==========
const mockEnvironmentState = {
  workbook: {
    sheets: [
      { name: "Sheet1", isActive: true },
      { name: "销售数据", isActive: false },
      { name: "产品目录", isActive: false },
      { name: "库存表", isActive: false },
      { name: "汇总", isActive: false }
    ],
    tables: [
      {
        name: "销售表",
        columns: ["日期", "产品", "销量", "单价", "金额"],
        sheetName: "销售数据",
        rowCount: 500,
        sampleData: [
          ["2024-01-15", "苹果", 150, 5.5, 825],
          ["2024-01-16", "香蕉", 0, 3.5, 0],
          ["2024-01-17", "橙子", -10, 4.0, -40],
          ["2024-02-01", "苹果", 200, 5.5, 1100],
          ["2024-02-15", "葡萄", 80, 12.0, 960]
        ]
      },
      {
        name: "产品目录",
        columns: ["产品ID", "产品名称", "类别", "成本价", "供应商"],
        sheetName: "产品目录",
        rowCount: 30
      },
      {
        name: "库存表",
        columns: ["产品", "库存数量", "安全库存", "最后盘点日期"],
        sheetName: "库存表",
        rowCount: 30
      }
    ],
    charts: [],
    namedRanges: []
  }
};

// ========== 工具注册表 ==========
const mockToolRegistry = {
  tools: new Map([
    ["excel_read_range", { name: "excel_read_range", description: "读取指定范围数据" }],
    ["excel_write_range", { name: "excel_write_range", description: "写入数据到范围" }],
    ["excel_write_cell", { name: "excel_write_cell", description: "写入单个单元格" }],
    ["get_table_schema", { name: "get_table_schema", description: "获取表格结构（列名、数据类型、行数、样本值）" }],
    ["sample_rows", { name: "sample_rows", description: "获取前N行样本数据" }],
    ["excel_sort_range", { name: "excel_sort_range", description: "对范围排序" }],
    ["excel_filter", { name: "excel_filter", description: "筛选数据" }],
    ["excel_format_range", { name: "excel_format_range", description: "格式化范围" }],
    ["excel_set_formula", { name: "excel_set_formula", description: "设置单元格公式" }],
    ["excel_fill_formula", { name: "excel_fill_formula", description: "填充公式到范围" }],
    ["excel_create_chart", { name: "excel_create_chart", description: "创建图表" }],
    ["excel_create_table", { name: "excel_create_table", description: "创建表格" }],
    ["excel_create_sheet", { name: "excel_create_sheet", description: "创建新工作表" }],
    ["excel_switch_sheet", { name: "excel_switch_sheet", description: "切换工作表" }],
    ["excel_delete_rows", { name: "excel_delete_rows", description: "删除行" }],
    ["excel_insert_rows", { name: "excel_insert_rows", description: "插入行" }],
    ["excel_auto_fit", { name: "excel_auto_fit", description: "自动调整列宽" }],
    ["excel_conditional_format", { name: "excel_conditional_format", description: "条件格式" }],
    ["excel_clear", { name: "excel_clear", description: "清除内容" }],
    ["excel_copy_range", { name: "excel_copy_range", description: "复制范围" }],
    ["excel_find", { name: "excel_find", description: "查找内容" }],
    ["excel_replace", { name: "excel_replace", description: "替换内容" }],
    ["excel_merge_cells", { name: "excel_merge_cells", description: "合并单元格" }],
    ["excel_set_number_format", { name: "excel_set_number_format", description: "设置数字格式" }],
    ["excel_calculate", { name: "excel_calculate", description: "执行计算(SUM/AVG/MAX/MIN等)" }],
    ["excel_get_used_range", { name: "excel_get_used_range", description: "获取已用范围" }],
    ["respond_to_user", { name: "respond_to_user", description: "回复用户" }],
    ["clarify_request", { name: "clarify_request", description: "向用户澄清请求" }],
  ]),
  getAll() { return Array.from(this.tools.values()); }
};

// ========== AI 后端调用 ==========
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
      path: '/agent/chat',
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
    req.setTimeout(90000, () => {  // 90秒超时，压力测试需要更长时间
      req.destroy();
      reject(new Error('Request timeout (90s)'));
    });
    req.write(postData);
    req.end();
  });
}

// ========== System Prompt ==========
function buildSystemPrompt() {
  const toolList = mockToolRegistry.getAll()
    .map(t => `- ${t.name}: ${t.description}`)
    .join('\n');

  return `你是Excel Office Add-in助手。根据用户请求生成执行计划。

## 可用工具
${toolList}

## 感知工具（重要！）
- get_table_schema: 获取表格结构（列名、数据类型、行数、样本值）
- sample_rows: 获取前N行样本数据
- excel_read_range: 读取指定区域数据

## 核心规则（必须严格遵守）
1. **先感知再操作**：写任何数据前，必须先调用感知工具确认目标区域结构
2. **操作后验证**：写入公式后，系统会自动验证结果是否正确
3. **★★★ 必须回复用户 ★★★**：每个计划的最后一步**必须是** respond_to_user 工具，向用户反馈操作结果。没有 respond_to_user 的计划是无效的！
4. 跨表操作前必须分别获取各表结构
5. 复杂任务要拆分成多个步骤，每个步骤必须是原子操作

## 公式生成规则（重要！）
当用户明确要求写公式时：
1. **直接生成公式，不要澄清**：即使信息不完整，也基于合理假设生成
2. 默认使用当前活动工作表的A、B、C列
3. 公式必须以 = 开头
4. 使用 excel_set_formula 设置公式
5. 对于嵌套IF、VLOOKUP等复杂公式，先感知数据结构，然后直接生成

例如：用户说"如果A列是X就显示B列，否则显示C列"
→ 直接生成：=IF(A1="X",B1,C1) 并填充到D列

## 澄清规则
1. 如果用户请求模糊不清，使用 clarify_request 工具询问
2. 如果任务超出能力范围，诚实说明并提供替代建议
3. 但如果用户明确要求写公式，不要过度澄清，先假设再生成

## 输出JSON格式
{
  "intent": "operation|query|clarify",
  "steps": [
    {"order":1, "action":"感知工具", "parameters":{}, "description":"先了解数据"},
    {"order":2, "action":"操作工具", "parameters":{}, "description":"执行操作"},
    {"order":N, "action":"respond_to_user", "parameters":{"message":"操作总结"}, "description":"最后一步必须回复用户"}
  ],
  "clarifyQuestion": "如果需要澄清，这里写问题"
}

## 错误示例（禁止）
- 计划只有1个步骤且不是respond_to_user
- 计划最后一步是excel_xxx操作而不是respond_to_user`;
}

// ========== 构建用户消息 ==========
function buildUserMessage(request) {
  return `## 当前请求
用户: ${request}

## 工作簿信息
${JSON.stringify(mockEnvironmentState.workbook, null, 2)}`;
}

// ========== 分析测试结果 ==========
function analyzeResult(testCase, response) {
  const issues = [];
  const strengths = [];
  
  try {
    const plan = typeof response.message === 'string' 
      ? JSON.parse(response.message) 
      : response.message;
    
    // 检查是否生成了计划
    if (!plan || !plan.steps || plan.steps.length === 0) {
      if (plan?.intent === 'clarify' && plan?.clarifyQuestion) {
        strengths.push('正确识别需要澄清');
        return { status: 'clarify', issues, strengths, plan };
      }
      issues.push('未生成有效计划');
      return { status: 'fail', issues, strengths, plan };
    }
    
    // 检查感知优先
    const firstAction = plan.steps[0]?.action;
    const perceptionTools = ['get_table_schema', 'sample_rows', 'excel_read_range', 'excel_get_used_range'];
    const hasEarlyPerception = plan.steps.slice(0, 2).some(s => perceptionTools.includes(s.action));
    
    if (!hasEarlyPerception && plan.steps.some(s => s.isWriteOperation)) {
      issues.push('写操作前未进行感知');
    } else if (hasEarlyPerception) {
      strengths.push('遵循感知优先');
    }
    
    // 检查步骤数量合理性
    if (plan.steps.length > 15) {
      issues.push(`步骤过多(${plan.steps.length}步)，可能过度拆分`);
    }
    if (plan.steps.length === 1 && testCase.category === 'complex') {
      issues.push('复杂任务只有1步，可能遗漏');
    }
    
    // 检查公式任务
    if (testCase.category === 'formula') {
      const hasFormula = plan.steps.some(s => 
        s.action === 'excel_set_formula' || 
        s.action === 'excel_fill_formula' ||
        (s.parameters?.formula)
      );
      if (hasFormula) {
        strengths.push('生成了公式');
        // 检查公式语法
        const formulaStep = plan.steps.find(s => s.parameters?.formula);
        if (formulaStep?.parameters?.formula) {
          const formula = formulaStep.parameters.formula;
          if (!formula.startsWith('=')) {
            issues.push('公式未以=开头');
          }
        }
      } else {
        issues.push('公式任务未生成公式');
      }
    }
    
    // 检查模糊指令处理
    if (testCase.category === 'fuzzy') {
      if (plan.intent === 'clarify') {
        strengths.push('对模糊指令请求澄清');
      } else if (plan.steps.length > 0) {
        strengths.push('对模糊指令做出合理假设');
      }
    }
    
    // 检查错误场景处理
    if (testCase.category === 'error') {
      if (plan.intent === 'clarify' || plan.steps.some(s => s.action === 'respond_to_user')) {
        strengths.push('识别了潜在问题');
      }
    }
    
    // 检查跨表操作
    if (testCase.category === 'cross-table') {
      const mentionsMultipleTables = plan.steps.filter(s => 
        s.parameters?.tableName || s.parameters?.sheetName
      ).length >= 2;
      if (mentionsMultipleTables) {
        strengths.push('正确处理跨表引用');
      } else {
        issues.push('跨表任务可能遗漏某些表');
      }
    }
    
    // 检查是否有 respond_to_user
    const hasResponse = plan.steps.some(s => s.action === 'respond_to_user');
    if (!hasResponse && testCase.category !== 'fuzzy') {
      issues.push('缺少用户反馈步骤');
    }
    
    const status = issues.length === 0 ? 'pass' : (issues.length <= 1 ? 'warn' : 'fail');
    return { status, issues, strengths, plan };
    
  } catch (e) {
    issues.push(`解析失败: ${e.message}`);
    return { status: 'error', issues, strengths: [], plan: null };
  }
}

// ========== 运行单个测试 ==========
async function runTest(testCase, index, total) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`[${index}/${total}] 📋 ${testCase.name}`);
  console.log(`📁 类别: ${testCase.category}`);
  console.log(`📝 请求: ${testCase.request.substring(0, 80)}${testCase.request.length > 80 ? '...' : ''}`);
  if (testCase.note) console.log(`📌 备注: ${testCase.note}`);
  console.log('='.repeat(70));

  const startTime = Date.now();
  
  try {
    const systemPrompt = buildSystemPrompt();
    const userMessage = buildUserMessage(testCase.request);
    
    console.log('\n⏳ 发送请求...');
    const response = await callAIBackend(userMessage, systemPrompt);
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`✅ 响应耗时: ${duration}s`);
    
    // 分析结果
    const analysis = analyzeResult(testCase, response);
    
    // 显示计划
    if (analysis.plan?.steps) {
      console.log(`\n📊 生成计划 (${analysis.plan.steps.length}步):`);
      analysis.plan.steps.forEach((step, i) => {
        console.log(`  ${i + 1}. ${step.action} - ${step.description?.substring(0, 50) || '无描述'}`);
      });
    }
    
    if (analysis.plan?.clarifyQuestion) {
      console.log(`\n❓ 澄清问题: ${analysis.plan.clarifyQuestion}`);
    }
    
    // 显示分析
    console.log('\n📈 分析结果:');
    if (analysis.strengths.length > 0) {
      analysis.strengths.forEach(s => console.log(`  ✅ ${s}`));
    }
    if (analysis.issues.length > 0) {
      analysis.issues.forEach(i => console.log(`  ⚠️  ${i}`));
    }
    
    // 状态标记
    const statusIcon = {
      pass: '🟢 通过',
      warn: '🟡 警告',
      fail: '🔴 失败',
      clarify: '🔵 澄清',
      error: '⛔ 错误'
    };
    console.log(`\n[结果] ${statusIcon[analysis.status]}`);
    
    return {
      name: testCase.name,
      category: testCase.category,
      status: analysis.status,
      duration: parseFloat(duration),
      issues: analysis.issues,
      strengths: analysis.strengths,
      stepCount: analysis.plan?.steps?.length || 0
    };
    
  } catch (error) {
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    console.log(`\n⛔ 测试异常: ${error.message}`);
    return {
      name: testCase.name,
      category: testCase.category,
      status: 'error',
      duration: parseFloat(duration),
      issues: [error.message],
      strengths: [],
      stepCount: 0
    };
  }
}

// ========== 生成报告 ==========
function generateReport(results) {
  console.log('\n' + '='.repeat(70));
  console.log('📊 暴力测试报告');
  console.log('='.repeat(70));
  
  // 按状态统计
  const statusCounts = { pass: 0, warn: 0, fail: 0, clarify: 0, error: 0 };
  results.forEach(r => statusCounts[r.status]++);
  
  console.log('\n📈 总体统计:');
  console.log(`  总测试数: ${results.length}`);
  console.log(`  🟢 通过: ${statusCounts.pass}`);
  console.log(`  🟡 警告: ${statusCounts.warn}`);
  console.log(`  🔵 澄清: ${statusCounts.clarify}`);
  console.log(`  🔴 失败: ${statusCounts.fail}`);
  console.log(`  ⛔ 错误: ${statusCounts.error}`);
  
  const successRate = ((statusCounts.pass + statusCounts.warn + statusCounts.clarify) / results.length * 100).toFixed(1);
  console.log(`  成功率: ${successRate}%`);
  
  // 按类别统计
  const categories = [...new Set(results.map(r => r.category))];
  console.log('\n📁 按类别统计:');
  categories.forEach(cat => {
    const catResults = results.filter(r => r.category === cat);
    const catPass = catResults.filter(r => ['pass', 'warn', 'clarify'].includes(r.status)).length;
    console.log(`  ${cat}: ${catPass}/${catResults.length} (${(catPass/catResults.length*100).toFixed(0)}%)`);
  });
  
  // 常见问题
  const allIssues = results.flatMap(r => r.issues);
  const issueCounts = {};
  allIssues.forEach(i => { issueCounts[i] = (issueCounts[i] || 0) + 1; });
  
  const topIssues = Object.entries(issueCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  if (topIssues.length > 0) {
    console.log('\n⚠️  常见问题 Top 5:');
    topIssues.forEach(([issue, count], i) => {
      console.log(`  ${i + 1}. ${issue} (${count}次)`);
    });
  }
  
  // 失败用例
  const failures = results.filter(r => r.status === 'fail' || r.status === 'error');
  if (failures.length > 0) {
    console.log('\n🔴 失败用例:');
    failures.forEach(f => {
      console.log(`  - ${f.name} [${f.category}]`);
      f.issues.forEach(i => console.log(`      └─ ${i}`));
    });
  }
  
  // 性能统计
  const avgDuration = (results.reduce((sum, r) => sum + r.duration, 0) / results.length).toFixed(1);
  const maxDuration = Math.max(...results.map(r => r.duration)).toFixed(1);
  const avgSteps = (results.reduce((sum, r) => sum + r.stepCount, 0) / results.length).toFixed(1);
  
  console.log('\n⏱️  性能统计:');
  console.log(`  平均响应时间: ${avgDuration}s`);
  console.log(`  最长响应时间: ${maxDuration}s`);
  console.log(`  平均步骤数: ${avgSteps}`);
  
  // 亮点
  const allStrengths = results.flatMap(r => r.strengths);
  const strengthCounts = {};
  allStrengths.forEach(s => { strengthCounts[s] = (strengthCounts[s] || 0) + 1; });
  
  const topStrengths = Object.entries(strengthCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  
  if (topStrengths.length > 0) {
    console.log('\n✅ 能力亮点:');
    topStrengths.forEach(([strength, count]) => {
      console.log(`  - ${strength} (${count}次)`);
    });
  }
  
  console.log('\n' + '='.repeat(70));
  
  return {
    total: results.length,
    successRate: parseFloat(successRate),
    statusCounts,
    topIssues,
    failures: failures.map(f => ({ name: f.name, issues: f.issues }))
  };
}

// ========== 入口 ==========
async function main() {
  console.log('🔥 Agent 暴力压力测试');
  console.log('='.repeat(70));
  console.log(`测试用例: ${stressTestCases.length} 个`);
  
  const categories = [...new Set(stressTestCases.map(t => t.category))];
  console.log(`测试类别: ${categories.join(', ')}`);
  console.log('='.repeat(70));

  // 支持命令行参数
  const args = process.argv.slice(2);
  let casesToRun = stressTestCases;
  
  if (args.includes('--category')) {
    const catIndex = args.indexOf('--category');
    const category = args[catIndex + 1];
    casesToRun = stressTestCases.filter(t => t.category === category);
    console.log(`\n🔍 筛选类别: ${category}, 共 ${casesToRun.length} 个`);
  } else if (args.includes('--quick')) {
    // 每个类别取1个
    casesToRun = [];
    categories.forEach(cat => {
      const first = stressTestCases.find(t => t.category === cat);
      if (first) casesToRun.push(first);
    });
    console.log(`\n⚡ 快速模式: 每类别1个, 共 ${casesToRun.length} 个`);
  } else if (args.length > 0 && !args[0].startsWith('--')) {
    const keyword = args[0];
    casesToRun = stressTestCases.filter(t => 
      t.name.includes(keyword) || t.request.includes(keyword)
    );
    console.log(`\n🔍 关键词: "${keyword}", 共 ${casesToRun.length} 个`);
  }

  if (casesToRun.length === 0) {
    console.log('❌ 没有匹配的测试用例');
    return;
  }

  const results = [];
  for (let i = 0; i < casesToRun.length; i++) {
    const result = await runTest(casesToRun[i], i + 1, casesToRun.length);
    results.push(result);
  }

  // 生成报告
  const report = generateReport(results);
  
  // 最终结论
  console.log('\n🎯 测试结论:');
  if (report.successRate >= 90) {
    console.log('  ✅ 助手表现优秀，大多数场景处理良好');
  } else if (report.successRate >= 70) {
    console.log('  🟡 助手表现一般，部分场景需要改进');
  } else {
    console.log('  🔴 助手表现较差，需要重点优化');
  }
  
  if (report.topIssues.length > 0) {
    console.log(`\n  🔧 优先修复: ${report.topIssues[0][0]}`);
  }
}

main().catch(console.error);
