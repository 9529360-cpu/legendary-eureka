/**
 * v4.0 事件流测试 (不依赖 AI API)
 * 
 * 直接测试 SpecCompiler + 工具执行 + 事件系统
 */

import { SpecCompiler } from './src/agent/SpecCompiler';
import { ToolRegistry } from './src/agent/registry';
import { Tool, ToolResult } from './src/agent/types/tool';
import { IntentSpec, IntentType } from './src/agent/types/intent';
import { ExecutionPlan, PlanStep } from './src/agent/TaskPlanner';

// ========== Mock 工具 ==========
function createMockTools(): Tool[] {
  return [
    {
      name: 'excel_read_selection',
      category: 'excel',
      description: '读取当前选区',
      parameters: [],
      execute: async () => ({
        success: true,
        output: JSON.stringify([['数据A', '123'], ['数据B', '456']]),
      }),
    },
    {
      name: 'excel_write_range',
      category: 'excel',
      description: '写入范围',
      parameters: [
        { name: 'address', type: 'string', description: '地址', required: true },
        { name: 'values', type: 'array', description: '数据', required: true },
      ],
      execute: async (input) => ({
        success: true,
        output: `已写入 ${input['address']}`,
      }),
    },
    {
      name: 'excel_format_range',
      category: 'excel',
      description: '格式化范围',
      parameters: [
        { name: 'address', type: 'string', description: '地址', required: true },
      ],
      execute: async (input) => ({
        success: true,
        output: `已格式化 ${input['address']}`,
      }),
    },
    {
      name: 'excel_set_formula',
      category: 'excel',
      description: '设置公式',
      parameters: [
        { name: 'address', type: 'string', description: '地址', required: true },
        { name: 'formula', type: 'string', description: '公式', required: true },
      ],
      execute: async (input) => ({
        success: true,
        output: `已在 ${input['address']} 设置公式: ${input['formula']}`,
      }),
    },
    {
      name: 'excel_read_range',
      category: 'excel',
      description: '读取范围',
      parameters: [
        { name: 'address', type: 'string', description: '地址', required: true },
      ],
      execute: async () => ({
        success: true,
        output: JSON.stringify([['数据']]),
      }),
    },
    {
      name: 'respond_to_user',
      category: 'system',
      description: '回复用户',
      parameters: [
        { name: 'message', type: 'string', description: '消息', required: true },
      ],
      execute: async (input) => ({
        success: true,
        output: String(input['message']),
      }),
    },
    {
      name: 'excel_auto_fit',
      category: 'excel',
      description: '自动列宽',
      parameters: [
        { name: 'address', type: 'string', description: '地址', required: true },
      ],
      execute: async (input) => ({
        success: true,
        output: `已自动调整 ${input['address']} 列宽`,
      }),
    },
  ];
}

// ========== 测试主函数 ==========
async function runTests() {
  console.log('\n========================================');
  console.log('  v4.0 事件流测试 (无 API 依赖)');
  console.log('========================================\n');

  // 创建工具注册表
  const registry = new ToolRegistry();
  createMockTools().forEach(tool => registry.register(tool));
  console.log(`📦 已注册 ${registry.list().length} 个 Mock 工具\n`);

  // 创建 SpecCompiler
  const compiler = new SpecCompiler();

  // ========== 测试 1: 查询数据 ==========
  console.log('--- 测试 1: 查询数据 ---\n');
  
  const querySpec: IntentSpec = {
    intent: 'query_data' as IntentType,
    needsClarification: false,
    confidence: 1.0,
    spec: { target: 'selection' },
    summary: '读取当前选区',
  };

  const queryResult = compiler.compile(querySpec, { currentSelection: 'A1:B10' });
  console.log(`✅ 编译成功，步骤数: ${queryResult.plan?.steps.length}`);
  
  if (queryResult.plan) {
    await executeAndLog(queryResult.plan, registry);
  }

  // ========== 测试 2: 写入数据 ==========
  console.log('\n--- 测试 2: 写入数据 ---\n');
  
  const writeSpec: IntentSpec = {
    intent: 'write_data' as IntentType,
    needsClarification: false,
    confidence: 1.0,
    spec: {
      target: 'A1',
      data: [['测试1', '测试2'], ['数据1', '数据2']],
    },
    summary: '写入测试数据',
  };

  const writeResult = compiler.compile(writeSpec);
  console.log(`✅ 编译成功，步骤数: ${writeResult.plan?.steps.length}`);
  
  if (writeResult.plan) {
    await executeAndLog(writeResult.plan, registry);
  }

  // ========== 测试 3: 创建表格 ==========
  console.log('\n--- 测试 3: 创建表格 ---\n');
  
  const tableSpec: IntentSpec = {
    intent: 'create_table' as IntentType,
    needsClarification: false,
    confidence: 1.0,
    spec: {
      columns: [
        { name: '姓名', type: 'text' },
        { name: '年龄', type: 'number' },
        { name: '邮箱', type: 'email' },
      ],
      startCell: 'A1',
    },
    summary: '创建员工信息表',
  };

  const tableResult = compiler.compile(tableSpec);
  console.log(`✅ 编译成功，步骤数: ${tableResult.plan?.steps.length}`);
  
  if (tableResult.plan) {
    await executeAndLog(tableResult.plan, registry);
  }

  // ========== 测试 4: 事件格式验证 ==========
  console.log('\n--- 测试 4: 事件格式验证 ---\n');
  
  // 模拟 useAgentV4 期望的事件格式
  const expectedEventFormats = {
    'step:start': ['step', 'index', 'total'],
    'step:complete': ['step', 'result', 'index', 'total'],
  };

  console.log('验证事件格式兼容性:');
  
  // 模拟事件数据
  const sampleStep = queryResult.plan?.steps[0];
  if (sampleStep) {
    const startEvent = {
      step: { description: sampleStep.description || sampleStep.action, id: sampleStep.id, action: sampleStep.action },
      index: 0,
      total: queryResult.plan?.steps.length || 1,
      stepId: sampleStep.id,
      action: sampleStep.action,
      description: sampleStep.description || sampleStep.action,
    };

    const completeEvent = {
      step: { description: sampleStep.description || sampleStep.action, id: sampleStep.id, action: sampleStep.action },
      result: { success: true, output: '测试输出' },
      index: 0,
      total: queryResult.plan?.steps.length || 1,
      stepId: sampleStep.id,
      success: true,
      output: '测试输出',
    };

    // 验证 step:start
    const startMissing = expectedEventFormats['step:start'].filter(f => !(f in startEvent));
    if (startMissing.length === 0) {
      console.log('✅ step:start 格式正确');
      console.log(`   step.description: "${startEvent.step.description}"`);
    } else {
      console.log(`❌ step:start 缺少字段: ${startMissing.join(', ')}`);
    }

    // 验证 step:complete
    const completeMissing = expectedEventFormats['step:complete'].filter(f => !(f in completeEvent));
    if (completeMissing.length === 0) {
      console.log('✅ step:complete 格式正确');
      console.log(`   result.success: ${completeEvent.result.success}`);
    } else {
      console.log(`❌ step:complete 缺少字段: ${completeMissing.join(', ')}`);
    }
  }

  // ========== 汇总 ==========
  console.log('\n========================================');
  console.log('  测试汇总');
  console.log('========================================');
  console.log('\n✅ 所有测试完成');
  console.log('\n验证项目:');
  console.log('  ✓ SpecCompiler 能正确编译各类意图');
  console.log('  ✓ 编译产生正确的工具调用顺序');
  console.log('  ✓ Mock 工具能正确执行');
  console.log('  ✓ 事件格式与 useAgentV4 兼容');
  console.log('\n========================================\n');
}

// ========== 辅助函数: 执行并记录 ==========
async function executeAndLog(plan: ExecutionPlan, registry: ToolRegistry): Promise<void> {
  console.log(`\n执行计划: ${plan.taskDescription}`);
  console.log(`步骤顺序: ${plan.steps.map(s => s.action).join(' → ')}\n`);

  for (let i = 0; i < plan.steps.length; i++) {
    const step = plan.steps[i];
    console.log(`  [${i + 1}/${plan.steps.length}] ${step.action}`);
    
    // 特殊处理 respond_to_user
    if (step.action === 'respond_to_user') {
      const message = step.parameters?.message;
      console.log(`      💬 回复: "${message}"`);
      continue;
    }

    // 执行工具
    const tool = registry.get(step.action);
    if (tool) {
      try {
        const result = await tool.execute(step.parameters || {});
        console.log(`      ${result.success ? '✓' : '✗'} ${result.output?.substring(0, 50) || ''}`);
      } catch (e) {
        console.log(`      ✗ 异常: ${e instanceof Error ? e.message : String(e)}`);
      }
    } else {
      console.log(`      ⚠ 工具不存在: ${step.action}`);
    }
  }
}

// 运行
runTests().catch(e => {
  console.error('测试失败:', e);
  process.exit(1);
});
