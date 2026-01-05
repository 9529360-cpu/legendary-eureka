/**
 * v4.0 AgentExecutor 事件系统测试
 * 
 * 验证事件格式与 useAgentV4 兼容
 * 
 * 注意: 此测试绕过 IntentParser（因为 Node.js 下 ApiService 需要绝对 URL）
 * 直接测试 SpecCompiler + AgentExecutor 的事件流
 */

import { AgentExecutor, ExecutorEvent } from './src/agent/AgentExecutor';
import { SpecCompiler } from './src/agent/SpecCompiler';
import { ToolRegistry } from './src/agent/registry';
import { Tool } from './src/agent/types/tool';
import { IntentSpec, IntentType } from './src/agent/types/intent';

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
        output: JSON.stringify([['测试数据', '123'], ['数据2', '456']]),
      }),
    },
    {
      name: 'excel_write_range',
      category: 'excel',
      description: '写入范围',
      parameters: [
        { name: 'address', description: '地址', type: 'string', required: true },
        { name: 'values', description: '数据', type: 'array', required: true },
      ],
      execute: async (input: Record<string, unknown>) => ({
        success: true,
        output: `已写入 ${input['address']}`,
      }),
    },
    {
      name: 'excel_format_range',
      category: 'excel',
      description: '格式化范围',
      parameters: [
        { name: 'address', description: '地址', type: 'string', required: true },
      ],
      execute: async (input: Record<string, unknown>) => ({
        success: true,
        output: `已格式化 ${input['address']}`,
      }),
    },
    {
      name: 'excel_set_formula',
      category: 'excel',
      description: '设置公式',
      parameters: [
        { name: 'address', description: '地址', type: 'string', required: true },
        { name: 'formula', description: '公式', type: 'string', required: true },
      ],
      execute: async (input: Record<string, unknown>) => ({
        success: true,
        output: `已在 ${input['address']} 设置公式`,
      }),
    },
    {
      name: 'excel_read_range',
      category: 'excel',
      description: '读取范围',
      parameters: [
        { name: 'address', description: '地址', type: 'string', required: true },
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
        { name: 'message', description: '消息', type: 'string', required: true },
      ],
      execute: async (input: Record<string, unknown>) => ({
        success: true,
        output: String(input['message']),
      }),
    },
  ];
}

// ========== 事件验证 ==========
interface EventValidation {
  type: string;
  requiredFields: string[];
}

const EVENT_VALIDATIONS: EventValidation[] = [
  {
    type: 'intent:parsed',
    requiredFields: ['intent', 'confidence'],
  },
  {
    type: 'plan:compiled',
    requiredFields: ['stepCount', 'plan'],
  },
  {
    type: 'step:start',
    requiredFields: ['step', 'index', 'total'],
  },
  {
    type: 'step:complete',
    requiredFields: ['step', 'result', 'index', 'total'],
  },
];

// ========== 测试主函数 ==========
async function runEventTests() {
  console.log('\n========================================');
  console.log('  AgentExecutor 事件系统测试');
  console.log('========================================\n');

  // 创建工具注册表
  const registry = new ToolRegistry();
  createMockTools().forEach(tool => registry.register(tool));
  
  // 创建执行器
  const executor = new AgentExecutor(registry);
  
  // 收集事件
  const collectedEvents: ExecutorEvent[] = [];
  const eventTypes: string[] = ['intent:parsed', 'plan:compiled', 'step:start', 'step:complete', 'execution:complete'];
  
  eventTypes.forEach(type => {
    executor.on(type as any, (event: ExecutorEvent) => {
      collectedEvents.push(event);
    });
  });

  console.log('📡 检查 AI 后端服务...');
  
  // 执行一个简单的查询
  try {
    console.log('🚀 执行请求: "读取当前表格"\n');
    
    const result = await executor.execute({
      userMessage: '读取当前表格',
      selection: { address: 'A1:B3', rowCount: 3, columnCount: 2 },
      activeSheet: 'Sheet1',
    });

    console.log(`\n✅ 执行完成: ${result.success ? '成功' : '失败'}`);
    console.log(`   消息: ${result.message.substring(0, 100)}...`);
    console.log(`   步骤数: ${result.executedSteps.length}`);

    // 验证事件
    console.log('\n--- 事件验证 ---\n');
    
    let allValid = true;
    
    for (const validation of EVENT_VALIDATIONS) {
      const events = collectedEvents.filter(e => e.type === validation.type);
      
      if (events.length === 0) {
        // step:error 可能没有，这是正常的
        if (validation.type === 'step:error') continue;
        
        console.log(`❌ 缺少事件: ${validation.type}`);
        allValid = false;
        continue;
      }

      const event = events[0];
      const data = event.data as Record<string, unknown>;
      const missingFields = validation.requiredFields.filter(f => !(f in data));
      
      if (missingFields.length > 0) {
        console.log(`❌ ${validation.type} 缺少字段: ${missingFields.join(', ')}`);
        console.log(`   实际字段: ${Object.keys(data).join(', ')}`);
        allValid = false;
      } else {
        console.log(`✅ ${validation.type} - 字段完整`);
        
        // 详细验证 step 相关事件
        if (validation.type === 'step:start' || validation.type === 'step:complete') {
          const step = data['step'] as Record<string, unknown>;
          if (!step || typeof step !== 'object') {
            console.log(`   ❌ step 不是对象`);
            allValid = false;
          } else if (!step['description']) {
            console.log(`   ❌ step.description 缺失`);
            allValid = false;
          } else {
            console.log(`   ✓ step.description: "${step['description']}"`);
          }
        }
      }
    }

    // 汇总
    console.log('\n========================================');
    console.log('  测试汇总');
    console.log('========================================');
    
    console.log(`\n收集到的事件: ${collectedEvents.length}`);
    collectedEvents.forEach(e => {
      console.log(`  - ${e.type}`);
    });
    
    console.log(`\n${allValid ? '✅ 所有事件格式正确' : '❌ 存在格式问题'}`);
    console.log('\n========================================\n');
    
    return allValid;

  } catch (error) {
    console.error('❌ 测试失败:', error);
    return false;
  }
}

// 运行测试
runEventTests()
  .then(success => process.exit(success ? 0 : 1))
  .catch(e => {
    console.error('测试脚本异常:', e);
    process.exit(1);
  });
