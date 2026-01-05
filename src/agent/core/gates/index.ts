/**
 * Gates 模块统一导出
 *
 * 反假完成闭环系统
 */

// 类型定义
export * from "./types";

// 状态机
export { StateMachine, stateMachine, TransitionResult } from "./StateMachine";

// 完成门槛
export { CompletionGate, completionGate, GateCheckResult } from "./CompletionGate";

// 公式验证器
export { FormulaValidator, formulaValidator } from "./FormulaValidator";

// 验证引擎
export { ValidationEngine, validationEngine, ValidationReport } from "./ValidationEngine";

// 提交包解析器
export { SubmissionParser, submissionParser, ParseResult } from "./SubmissionParser";

// 拦截器
export {
  CompletionInterceptor,
  SelfReferenceInterceptor,
  MaxIterationsInterceptor,
  completionInterceptor,
  selfReferenceInterceptor,
  maxIterationsInterceptor,
  InterceptResult,
} from "./Interceptors";

// 反假完成控制器
export {
  AntiHallucinationController,
  antiHallucinationController,
  TurnResult,
} from "./AntiHallucinationController";
