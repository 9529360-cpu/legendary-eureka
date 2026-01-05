/**
 * 澄清机制专项测试
 * 测试对模糊+有副作用请求的澄清行为
 */

const http = require('http');

// 模拟工作簿环境
const mockEnvironmentState = {
  workbook: {
    sheets: [{ name: "Sheet1", isActive: true }],
    tables: [
      { name: "销售表", columns: ["日期", "客户", "UID", "金额"], rowCount: 100 }
    ]
  }
};

// 澄清测试用例
const clarifyTestCases = [
  {
    name: "删除没用的 - 应触发澄清",
    request: "这个表格的可读性很低 请帮我优化一下 删除没有用的",
    expectClarify: true
  },
  {
    name: "清理数据 - 应触发澄清",
    request: "帮我把这个表清理一下",
    expectClarify: true
  },
  {
    name: "优化表格 - 应触发澄清",
    request: "这个表太乱了，帮我优化一下",
    expectClarify: true
  },
  {
    name: "删除空行 - 明确，不需要澄清",
    request: "删除所有空行",
    expectClarify: false
  },
  {
    name: "按金额排序 - 明确，不需要澄清",
    request: "把销售表按金额从大到小排序",
    expectClarify: false
  },
  {
    name: "删除A列 - 明确，不需要澄清",
    request: "删除A列",
    expectClarify: false
  }
];

// 调用 AI 后端
async function callAIBackend(message, systemPrompt) {
  return new Promise((resolve, reject) => {
    const postData = JSON.stringify({ message, systemPrompt, responseFormat: "json" });
    const req = http.request({
      hostname: 'localhost',
      port: 3001,
      path: '/agent/chat',
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(postData) }
    }, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(JSON.parse(data)));
    });
    req.on('error', reject);
    req.setTimeout(60000, () => { req.destroy(); reject(new Error('Timeout')); });
    req.write(postData);
    req.end();
  });
}

// 构建 System Prompt
function buildSystemPrompt() {
  return `你是Excel Office Add-in助手。根据用户请求生成执行计划。

## 可用工具
- excel_read_range: 读取数据
- excel_sort_range: 排序
- excel_delete_rows: 删除行
- excel_delete_columns: 删除列
- excel_clear: 清除内容
- respond_to_user: 回复用户
- clarify_request: 向用户澄清模糊请求

## ★★★ 澄清优先规则（最重要！）★★★
以下情况**必须**先用 clarify_request 澄清，**禁止**直接操作：

1. **模糊+删除类请求**：
   - "删除没用的" → 什么是"没用的"？空行？空列？重复数据？
   - "清理一下" → 清理什么？格式？数据？
   - "优化表格" → 优化什么？格式？结构？删除数据？

2. **有副作用+不明确范围**：
   - "把这些数据整理一下" → 整理到哪里？
   - "帮我处理一下" → 处理什么？

3. **澄清示例**：
   用户说"删除没用的列"
   → 先 clarify_request: {question: "您想删除哪些列？", options: ["空白列", "指定的列"]}

## 明确请求不需要澄清：
- "删除A列" → 明确，直接执行
- "删除所有空行" → 明确，直接执行
- "按金额排序" → 明确，直接执行

## 输出JSON格式
{
  "intent": "query" | "operation" | "clarify",
  "clarifyReason": "如果intent是clarify，说明原因",
  "steps": [{"order":1, "action":"工具名", "parameters":{}}]
}`;
}

function buildUserPrompt(request) {
  return `用户请求: ${request}\n\n工作簿信息:\n${JSON.stringify(mockEnvironmentState.workbook, null, 2)}`;
}

function parsePlan(response) {
  const msg = response.message || '';
  const match = msg.match(/\{[\s\S]*\}/);
  if (match) {
    try { return JSON.parse(match[0]); } catch { return null; }
  }
  return null;
}

async function runTest(testCase) {
  console.log(`\n${'='.repeat(50)}`);
  console.log(`📋 ${testCase.name}`);
  console.log(`📝 请求: ${testCase.request}`);
  console.log(`期望: ${testCase.expectClarify ? '需要澄清' : '直接执行'}`);

  const response = await callAIBackend(buildUserPrompt(testCase.request), buildSystemPrompt());
  const plan = parsePlan(response);

  if (!plan) {
    console.log('❌ 计划解析失败');
    return { success: false };
  }

  const isClarify = plan.intent === 'clarify' || 
                    plan.steps?.some(s => s.action === 'clarify_request');
  
  console.log(`  Intent: ${plan.intent}`);
  console.log(`  Steps: ${plan.steps?.map(s => s.action).join(' -> ')}`);
  if (plan.clarifyReason) {
    console.log(`  澄清原因: ${plan.clarifyReason}`);
  }

  const passed = isClarify === testCase.expectClarify;
  console.log(`\n${passed ? '✅ 通过' : '❌ 失败'} - ${isClarify ? '触发了澄清' : '直接执行'}`);
  return { success: passed, plan };
}

async function main() {
  console.log('🧪 澄清机制专项测试');
  console.log('=' .repeat(50));

  const args = process.argv.slice(2);
  let cases = clarifyTestCases;
  
  if (args[0]) {
    cases = [{ name: '自定义测试', request: args[0], expectClarify: true }];
  }

  let passed = 0, failed = 0;
  for (const tc of cases) {
    const result = await runTest(tc);
    result.success ? passed++ : failed++;
  }

  console.log(`\n${'='.repeat(50)}`);
  console.log(`📊 汇总: ${passed}/${passed + failed} 通过 (${((passed/(passed+failed))*100).toFixed(0)}%)`);
  console.log('=' .repeat(50));
}

main().catch(console.error);
