/**
 * 反假完成系统 - 核心类型定义
 *
 * 基于用户协议设计：
 * - AgentRun: 运行状态容器
 * - Artifact: 可执行产物（公式/步骤/模板）
 * - Checklist: 完成放行清单
 * - Validation: 系统验证结果
 */

// ========== 状态机枚举 ==========

/**
 * Agent 运行状态（6 阶段状态机）
 * 不允许 INIT → DEPLOYED 跳跃
 */
export enum AgentState {
  INIT = "INIT",
  ANALYZED = "ANALYZED",
  DESIGNED = "DESIGNED",
  EXECUTED = "EXECUTED",
  VERIFIED = "VERIFIED",
  DEPLOYED = "DEPLOYED",
}

/**
 * 目标平台
 */
export enum Platform {
  EXCEL = "excel",
  GOOGLE_SHEETS = "google_sheets",
}

/**
 * 产物类型
 */
export enum ArtifactType {
  FORMULA = "FORMULA",
  STEPS = "STEPS",
  TEMPLATE = "TEMPLATE",
  SCHEMA_PLAN = "SCHEMA_PLAN",
}

/**
 * 验证状态
 */
export enum ValidationStatus {
  PASS = "PASS",
  FAIL = "FAIL",
  WARN = "WARN",
}

// ========== 产物定义 ==========

/**
 * 产物目标位置
 */
export interface ArtifactTarget {
  sheet?: string;
  range?: string;
  column?: string;
  cell?: string;
}

/**
 * 可执行产物（公式/步骤/模板）
 */
export interface Artifact {
  id: string;
  type: ArtifactType;
  platform: Platform;
  target: ArtifactTarget;
  content: string;
  version: string;
  createdAt: number;
}

// ========== 完成清单 ==========

/**
 * 完成放行清单（硬性门槛）
 * 所有项目必须为 true 才允许结束任务
 */
export interface Checklist {
  /** 是否给出可复制执行的公式或步骤 */
  hasExecutableArtifact: boolean;
  /** 是否明确公式放置位置（列/行） */
  hasPlacementInfo: boolean;
  /** 是否验证新增行自动生效 */
  supportsAutoExpand: boolean;
  /** 是否避免结果列自引用 */
  avoidsSelfReference: boolean;
  /** 是否给出至少 3 条验收测试 */
  has3AcceptanceTests: boolean;
  /** 是否给出失败回退方案 */
  hasFallbackPlan: boolean;
  /** 是否给出部署与防错清单 */
  hasDeployNotes: boolean;
}

/**
 * 创建空的 Checklist
 */
export function createEmptyChecklist(): Checklist {
  return {
    hasExecutableArtifact: false,
    hasPlacementInfo: false,
    supportsAutoExpand: false,
    avoidsSelfReference: false,
    has3AcceptanceTests: false,
    hasFallbackPlan: false,
    hasDeployNotes: false,
  };
}

/**
 * 检查 Checklist 是否全部通过
 */
export function isChecklistComplete(checklist: Checklist): boolean {
  return (
    checklist.hasExecutableArtifact &&
    checklist.hasPlacementInfo &&
    checklist.supportsAutoExpand &&
    checklist.avoidsSelfReference &&
    checklist.has3AcceptanceTests &&
    checklist.hasFallbackPlan &&
    checklist.hasDeployNotes
  );
}

// ========== 验证结果 ==========

/**
 * 单条验证结果
 */
export interface Validation {
  name: string;
  ruleId: string;
  status: ValidationStatus;
  reason?: string;
  details?: Record<string, unknown>;
}

/**
 * 验收测试
 */
export interface AcceptanceTest {
  id: string;
  description: string;
  expectedResult: string;
  passed?: boolean;
}

/**
 * 回退方案
 */
export interface FallbackPlan {
  condition: string;
  action: string;
}

/**
 * 部署说明
 */
export interface DeployNotes {
  protectedRanges?: string[];
  namingConventions?: string[];
  permissions?: string[];
  changeImpact?: string[];
}

// ========== 提交包 ==========

/**
 * 模型提交包（解析自模型输出）
 */
export interface Submission {
  proposedState: AgentState;
  artifacts: Artifact[];
  acceptanceTests: AcceptanceTest[];
  fallback: FallbackPlan[];
  deployNotes?: DeployNotes;
  nextAction?: {
    systemWillValidate: string;
    userNeedsToProvide?: string;
    ifFailAgentWill?: string;
  };
  rawOutput: string;
}

// ========== Agent 运行容器 ==========

/**
 * 消息角色
 */
export type MessageRole = "user" | "assistant" | "system" | "tool";

/**
 * 对话消息
 */
export interface Message {
  role: MessageRole;
  content: string;
  timestamp: number;
  metadata?: Record<string, unknown>;
}

/**
 * Agent 运行状态容器
 */
export interface AgentRun {
  runId: string;
  userId: string;
  taskId: string;
  state: AgentState;
  iteration: number;
  maxIterations: number;
  artifacts: Artifact[];
  checklist: Checklist;
  validations: Validation[];
  lastModelOutput: string;
  history: Message[];
  createdAt: number;
  updatedAt: number;
}

/**
 * 创建新的 AgentRun
 */
export function createAgentRun(userId: string, taskId: string, maxIterations = 8): AgentRun {
  const now = Date.now();
  return {
    runId: `run_${now}_${Math.random().toString(36).slice(2, 8)}`,
    userId,
    taskId,
    state: AgentState.INIT,
    iteration: 0,
    maxIterations,
    artifacts: [],
    checklist: createEmptyChecklist(),
    validations: [],
    lastModelOutput: "",
    history: [],
    createdAt: now,
    updatedAt: now,
  };
}

// ========== 语义抽取结果 ==========

/**
 * 语义抽取结果（协议要求显式输出）
 */
export interface SemanticExtract {
  intent: string;
  entities: Record<string, string | string[]>;
  constraints: string[];
}

/**
 * 快速诊断结果
 */
export interface QuickDiagnosis {
  cause: string;
  probability: number;
  verificationSteps: string[];
}
