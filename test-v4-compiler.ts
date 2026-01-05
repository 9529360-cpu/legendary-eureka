/**
 * v4.0 SpecCompiler + AgentExecutor 测试脚本
 * 
 * 不依赖 Excel，仅测试编译逻辑
 * 
 * 运行方式: npx ts-node test-v4-compiler.ts
 * 或者:    npx tsx test-v4-compiler.ts
 */

import { SpecCompiler } from './src/agent/SpecCompiler';
import { IntentSpec, IntentType } from './src/agent/types/intent';

// ========== 测试用例 ==========
interface TestCase {
  name: string;
  intentSpec: IntentSpec;
  expectedSteps: number;
  expectedActions: string[];
}

const TEST_CASES: TestCase[] = [
  {
    name: '查询数据',
    intentSpec: {
      intent: 'query_data' as IntentType,
      needsClarification: false,
      confidence: 1.0,
      spec: {
        target: 'selection',
      },
      summary: '读取当前选区',
    },
    expectedSteps: 2,
    expectedActions: ['excel_read_selection', 'respond_to_user'],
  },
  {
    name: '写入数据',
    intentSpec: {
      intent: 'write_data' as IntentType,
      needsClarification: false,
      confidence: 1.0,
      spec: {
        target: 'A1',
        data: [['测试数据']],
      },
      summary: '写入测试数据到 A1',
    },
    expectedSteps: 2,
    expectedActions: ['excel_write_range', 'respond_to_user'],
  },
  {
    name: '创建公式',
    intentSpec: {
      intent: 'create_formula' as IntentType,
      needsClarification: false,
      confidence: 1.0,
      spec: {
        type: 'sum',
        sourceRange: 'A1:A10',
        targetCell: 'A11',
      },
      summary: '对 A1:A10 求和',
    },
    expectedSteps: 2, // 至少有写公式 + 回复
    expectedActions: ['excel_set_formula', 'respond_to_user'],
  },
  {
    name: '格式化范围',
    intentSpec: {
      intent: 'format_range' as IntentType,
      needsClarification: false,
      confidence: 1.0,
      spec: {
        range: 'A1:D1',
        format: {
          bold: true,
          backgroundColor: '#FF0000',
        },
      },
      summary: '加粗并设置背景色',
    },
    expectedSteps: 2,
    expectedActions: ['excel_format_range', 'respond_to_user'],
  },
  {
    name: '需要澄清',
    intentSpec: {
      intent: 'clarify' as IntentType,
      needsClarification: true,
      clarificationQuestion: '请问您想删除什么内容？',
      confidence: 0.3,
      spec: null,
      summary: '需要澄清',
    },
    expectedSteps: 1,
    expectedActions: ['clarify_request'],  // 澄清用 clarify_request 工具
  },
];

// ========== 测试运行 ==========
function runTests() {
  console.log('\n========================================');
  console.log('  SpecCompiler 测试');
  console.log('========================================\n');

  const compiler = new SpecCompiler();
  let passed = 0;
  let failed = 0;
  const failures: string[] = [];

  for (const testCase of TEST_CASES) {
    console.log(`\n--- 测试: ${testCase.name} ---`);
    
    try {
      const result = compiler.compile(testCase.intentSpec, { currentSelection: 'A1' });
      
      if (!result.success && !testCase.intentSpec.needsClarification) {
        console.log(`❌ 编译失败: ${result.error}`);
        failed++;
        failures.push(`${testCase.name}: 编译失败 - ${result.error}`);
        continue;
      }

      if (testCase.intentSpec.needsClarification && result.needsClarification) {
        console.log('✅ 正确识别需要澄清');
        passed++;
        continue;
      }

      const plan = result.plan;
      if (!plan) {
        console.log('❌ 未生成执行计划');
        failed++;
        failures.push(`${testCase.name}: 未生成执行计划`);
        continue;
      }

      const actualSteps = plan.steps.length;
      const actualActions = plan.steps.map(s => s.action);
      
      // 检查步骤数
      const stepCountOk = actualSteps >= testCase.expectedSteps;
      
      // 检查必须的 action 是否存在
      const actionsOk = testCase.expectedActions.every(a => actualActions.includes(a));
      
      if (stepCountOk && actionsOk) {
        console.log(`✅ 通过`);
        console.log(`   步骤数: ${actualSteps} (预期 >= ${testCase.expectedSteps})`);
        console.log(`   动作: ${actualActions.join(' → ')}`);
        passed++;
      } else {
        console.log(`❌ 失败`);
        console.log(`   步骤数: ${actualSteps} (预期 >= ${testCase.expectedSteps})`);
        console.log(`   动作: ${actualActions.join(' → ')}`);
        console.log(`   预期动作: ${testCase.expectedActions.join(', ')}`);
        failed++;
        failures.push(`${testCase.name}: 步骤/动作不匹配`);
      }

    } catch (e) {
      console.log(`❌ 异常: ${e instanceof Error ? e.message : String(e)}`);
      failed++;
      failures.push(`${testCase.name}: ${e instanceof Error ? e.message : String(e)}`);
    }
  }

  // 汇总
  console.log('\n========================================');
  console.log('  测试汇总');
  console.log('========================================');
  console.log(`\n✅ 通过: ${passed}/${TEST_CASES.length}`);
  console.log(`❌ 失败: ${failed}/${TEST_CASES.length}`);
  
  if (failures.length > 0) {
    console.log('\n失败详情:');
    failures.forEach(f => console.log(`  - ${f}`));
  }
  
  console.log('\n========================================\n');
  
  return failed === 0;
}

// 运行
const success = runTests();
process.exit(success ? 0 : 1);
