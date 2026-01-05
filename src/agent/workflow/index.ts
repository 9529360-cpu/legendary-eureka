/**
 * 工作流模块统一导出
 *
 * @packageDocumentation
 */

// 导出工作流引擎
export {
  createWorkflowEvent,
  createInitialWorkflowState,
  WorkflowEvents,
  WorkflowContext,
  WorkflowEventRegistry,
  WorkflowEventStream,
  createSimpleWorkflow,
} from "./WorkflowEngine";

// 重导出类型（从 types 模块）
export type {
  WorkflowEvent,
  WorkflowState,
  WorkflowEventHandler,
  SimpleWorkflow,
  WorkflowContextInterface,
  WorkflowEventStreamInterface,
  WorkflowEventRegistryInterface,
  WorkflowEventFactory,
  AgentStreamData,
  AgentOutputData,
  AgentStreamStructuredOutputData,
  PlanStep,
  ExecutionPlan,
  PlanExecutionResult,
  TaskProgress,
  ProgressStep,
} from "../types/workflow";
