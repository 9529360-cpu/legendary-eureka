/**
 * Excel 其他工具
 *
 * 包含工具：
 * - createRespondToUserTool: 响应用户
 * - createClarifyRequestTool: 澄清请求
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";

/**
 * 响应用户工具
 */
export function createRespondToUserTool(): Tool {
  return {
    name: "respond_to_user",
    description:
      "向用户发送最终回复。当你完成了用户的任务，或者需要向用户提供信息/解答时，使用此工具。调用此工具后，任务将被标记为完成。",
    category: "general",
    parameters: [
      {
        name: "message",
        type: "string",
        description: "发送给用户的消息，应该清晰地总结完成了什么操作或回答用户的问题",
        required: true,
      },
      {
        name: "success",
        type: "boolean",
        description: "任务是否成功完成（默认为 true）",
        required: false,
      },
    ],
    execute: async (params: Record<string, unknown>) => {
      const message = params.message as string;
      const success = params.success !== false;

      return {
        success: success,
        output: message,
        data: {
          isResponse: true,
          shouldComplete: true,
        },
      };
    },
  };
}

/**
 * 澄清请求工具
 */
export function createClarifyRequestTool(): Tool {
  return {
    name: "clarify_request",
    description:
      "向用户澄清模糊的请求。当用户请求不够明确且可能有副作用（如删除、修改数据）时使用。调用此工具后，等待用户回复后再继续。",
    category: "general",
    parameters: [
      {
        name: "question",
        type: "string",
        description: "向用户提问的内容，应该清晰地说明需要澄清什么",
        required: true,
      },
      {
        name: "options",
        type: "array",
        description: "提供给用户的选项列表（可选）",
        required: false,
      },
      {
        name: "context",
        type: "string",
        description: "为什么需要澄清的上下文说明（可选）",
        required: false,
      },
    ],
    execute: async (params: Record<string, unknown>) => {
      const question = params.question as string;
      const options = params.options as string[] | undefined;
      const context = params.context as string | undefined;

      let message = question;
      if (options && options.length > 0) {
        message += "\n\n请选择：\n" + options.map((opt, i) => `${i + 1}. ${opt}`).join("\n");
      }
      if (context) {
        message = `${context}\n\n${message}`;
      }

      return {
        success: true,
        output: message,
        data: {
          isClarification: true,
          shouldComplete: true,
          question,
          options,
        },
      };
    },
  };
}

/**
 * 创建所有其他工具
 */
export function createMiscTools(): Tool[] {
  return [
    createRespondToUserTool(),
    createClarifyRequestTool(),
  ];
}
