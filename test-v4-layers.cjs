/**
 * v4.0 三层架构测试脚本
 * 
 * 测试内容:
 * 1. IntentParser System Prompt 验证
 * 2. API 调用验证
 * 3. SpecCompiler 编译验证
 * 
 * 运行方式: node test-v4-layers.cjs
 */

const http = require('http');

// ========== 测试用例 ==========
const TEST_CASES = [
  { input: "读取当前表格", expectedIntent: "query_data", expectClarify: false },
  { input: "看看这个表", expectedIntent: "query_data", expectClarify: false },
  { input: "帮我求和", expectedIntent: "create_formula", expectClarify: false },
  { input: "在A1写入测试", expectedIntent: "write_data", expectClarify: false },
  { input: "删除没用的", expectedIntent: "clarify", expectClarify: true },
  { input: "你好", expectedIntent: "respond_only", expectClarify: false },
];

// ========== 测试结果 ==========
const testResults = [];

// ========== API 调用函数 ==========
function callAiApi(message, systemPrompt) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify({
      message,
      systemPrompt,
      responseFormat: 'json',
    });

    const options = {
      hostname: 'localhost',
      port: 3001,
      path: '/agent/chat',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data),
      },
    };

    const req = http.request(options, (res) => {
      let body = '';
      res.on('data', (chunk) => body += chunk);
      res.on('end', () => {
        try {
          const result = JSON.parse(body);
          resolve(result);
        } catch (e) {
          reject(new Error(`JSON解析失败: ${body.substring(0, 200)}`));
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(30000, () => reject(new Error('请求超时')));
    req.write(data);
    req.end();
  });
}

// ========== 简化版 System Prompt (与 IntentParser.ts 一致) ==========
const SYSTEM_PROMPT = `你是 Excel 智能助手的意图理解模块。你的任务是理解用户想做什么，而不是怎么做。

## 你的职责
1. 理解用户的真实意图
2. 提取关键信息（范围、数据、条件等）
3. 判断信息是否完整
4. 如果信息不足，指出需要澄清什么

## 意图类型（选择最匹配的一个）
- create_table: 创建新表格
- write_data: 写入数据
- update_data: 更新数据
- delete_data: 删除数据
- format_range: 格式化（加粗、颜色、边框等）
- create_formula: 创建公式
- analyze_data: 分析数据
- create_chart: 创建图表
- sort_data: 排序
- filter_data: 筛选
- query_data: 查询/读取
- create_sheet: 创建工作表
- clarify: 需要澄清
- respond_only: 只需回复，不需操作

## ★★★ 澄清规则（非常重要）★★★

### 不需要澄清的情况（直接执行！）
- "读取当前表格" → query_data，读取选中区域或活动工作表
- "看看这个表" → query_data，读取并返回概览
- "帮我求和" → create_formula，对选中区域求和
- "在A1写入xxx" → write_data，明确的写入操作

### 必须澄清的情况
- "删除没用的" → 什么是"没用的"？
- "清理一下" → 清理什么？

### ⚠️ 重要：不要过度澄清！
查询类操作不需要澄清，直接返回数据。

## 输出 JSON 格式
{
  "intent": "query_data|write_data|...",
  "needsClarification": false,
  "spec": { ... },
  "summary": "简要描述"
}`;

// ========== 测试函数 ==========
async function runTests() {
  console.log('\n========================================');
  console.log('  v4.0 三层架构测试');
  console.log('========================================\n');

  // 首先检查服务是否可用
  console.log('📡 检查 AI 后端服务 (localhost:3001)...');
  try {
    await callAiApi('test', 'test');
    console.log('✅ AI 后端服务正常\n');
  } catch (e) {
    console.log('❌ AI 后端服务不可用:', e.message);
    console.log('\n请先启动服务: npm run dev:full\n');
    return;
  }

  // 运行测试用例
  for (let i = 0; i < TEST_CASES.length; i++) {
    const testCase = TEST_CASES[i];
    console.log(`\n--- 测试 ${i + 1}/${TEST_CASES.length}: "${testCase.input}" ---`);
    
    try {
      const response = await callAiApi(testCase.input, SYSTEM_PROMPT);
      const message = response.message || '';
      
      // 尝试解析 JSON
      const jsonMatch = message.match(/\{[\s\S]*\}/);
      if (!jsonMatch) {
        console.log('❌ 未返回有效 JSON');
        console.log('   原始回复:', message.substring(0, 200));
        testResults.push({
          input: testCase.input,
          passed: false,
          error: '未返回 JSON',
          rawResponse: message.substring(0, 200),
        });
        continue;
      }

      const parsed = JSON.parse(jsonMatch[0]);
      const actualIntent = parsed.intent;
      const actualClarify = parsed.needsClarification === true;

      // 验证意图
      const intentMatch = actualIntent === testCase.expectedIntent;
      const clarifyMatch = actualClarify === testCase.expectClarify;
      const passed = intentMatch && clarifyMatch;

      if (passed) {
        console.log('✅ 通过');
        console.log(`   意图: ${actualIntent} (预期: ${testCase.expectedIntent})`);
        console.log(`   澄清: ${actualClarify} (预期: ${testCase.expectClarify})`);
      } else {
        console.log('❌ 失败');
        console.log(`   意图: ${actualIntent} (预期: ${testCase.expectedIntent}) ${intentMatch ? '✓' : '✗'}`);
        console.log(`   澄清: ${actualClarify} (预期: ${testCase.expectClarify}) ${clarifyMatch ? '✓' : '✗'}`);
        if (parsed.clarificationQuestion) {
          console.log(`   澄清问题: ${parsed.clarificationQuestion}`);
        }
      }

      testResults.push({
        input: testCase.input,
        passed,
        expectedIntent: testCase.expectedIntent,
        actualIntent,
        expectedClarify: testCase.expectClarify,
        actualClarify,
        clarificationQuestion: parsed.clarificationQuestion,
      });

    } catch (e) {
      console.log('❌ 测试异常:', e.message);
      testResults.push({
        input: testCase.input,
        passed: false,
        error: e.message,
      });
    }
  }

  // 汇总
  console.log('\n========================================');
  console.log('  测试汇总');
  console.log('========================================');
  
  const passed = testResults.filter(r => r.passed).length;
  const failed = testResults.filter(r => !r.passed).length;
  
  console.log(`\n✅ 通过: ${passed}/${testResults.length}`);
  console.log(`❌ 失败: ${failed}/${testResults.length}`);
  
  if (failed > 0) {
    console.log('\n失败的测试用例:');
    testResults.filter(r => !r.passed).forEach(r => {
      console.log(`  - "${r.input}"`);
      if (r.error) {
        console.log(`    错误: ${r.error}`);
      } else {
        console.log(`    预期意图: ${r.expectedIntent}, 实际: ${r.actualIntent}`);
        console.log(`    预期澄清: ${r.expectedClarify}, 实际: ${r.actualClarify}`);
      }
    });
  }

  console.log('\n========================================\n');
  
  return { passed, failed, total: testResults.length, results: testResults };
}

// 运行测试
runTests()
  .then(summary => {
    if (summary && summary.failed > 0) {
      process.exit(1);
    }
  })
  .catch(e => {
    console.error('测试脚本异常:', e);
    process.exit(1);
  });
