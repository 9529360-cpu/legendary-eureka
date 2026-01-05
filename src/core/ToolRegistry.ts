/**
 * ToolRegistry - LLM工具白名单注册表
 *
 * 设计原则：
 * 1. 所有LLM可调用的工具必须在此注册
 * 2. 每个工具必须有明确的输入参数Schema
 * 3. 工具描述必须清晰说明功能和限制
 * 4. 工具按功能域分组，便于维护和扩展
 */

import { ToolDefinition, ToolCategory, ParameterType, ParameterValidationRule } from "../types";

/**
 * Excel操作工具定义
 */
export const EXCEL_TOOLS: ToolDefinition[] = [
  {
    id: "excel.select_range",
    name: "select_range",
    category: ToolCategory.EXCEL_OPERATION,
    description: "选择Excel中的单元格范围",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格地址，如A1:B10",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
          errorMessage: "请输入有效的单元格地址，如A1或A1:B10",
        },
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "A1:C5" },
      description: "选择A1到C5的单元格范围",
    },
  },
  {
    id: "excel.set_cell_value",
    name: "set_cell_value",
    category: ToolCategory.EXCEL_OPERATION,
    description: "设置单个单元格的值",
    parameters: [
      {
        name: "cellAddress",
        type: ParameterType.STRING,
        description: "单元格地址，如A1",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+$",
          errorMessage: "请输入有效的单元格地址，如A1",
        },
      },
      {
        name: "value",
        type: ParameterType.ANY,
        description: "要设置的值（字符串、数字、布尔值）",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { cellAddress: "B2", value: "总计" },
      description: '在B2单元格设置"总计"文本',
    },
  },
  {
    id: "excel.set_range_values",
    name: "set_range_values",
    category: ToolCategory.EXCEL_OPERATION,
    description: "设置单元格范围的值",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围地址，如A1:B2",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
          errorMessage: "请输入有效的单元格范围地址",
        },
      },
      {
        name: "values",
        type: ParameterType.ARRAY,
        description: "二维数组的值",
        required: true,
        validation: {
          minItems: 1,
          errorMessage: "至少需要一个值",
        },
      },
    ],
    returns: "void",
    example: {
      input: {
        rangeAddress: "A1:B2",
        values: [
          ["姓名", "年龄"],
          ["张三", 25],
        ],
      },
      description: "在A1:B2范围设置表格数据",
    },
  },
  {
    id: "excel.get_range_values",
    name: "get_range_values",
    category: ToolCategory.EXCEL_OPERATION,
    description: "获取单元格范围的值",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围地址",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
          errorMessage: "请输入有效的单元格范围地址",
        },
      },
    ],
    returns: "any[][]",
    example: {
      input: { rangeAddress: "A1:C3" },
      description: "获取A1:C3范围的值",
    },
  },
  {
    id: "excel.clear_range",
    name: "clear_range",
    category: ToolCategory.EXCEL_OPERATION,
    description: "清除单元格范围的内容",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围地址",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
          errorMessage: "请输入有效的单元格范围地址",
        },
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "D5:F10" },
      description: "清除D5:F10范围的内容",
    },
  },
  {
    id: "excel.format_range",
    name: "format_range",
    category: ToolCategory.EXCEL_OPERATION,
    description: "格式化单元格范围",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围地址",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
          errorMessage: "请输入有效的单元格范围地址",
        },
      },
      {
        name: "format",
        type: ParameterType.OBJECT,
        description: "格式化选项",
        required: true,
        properties: {
          font: {
            type: ParameterType.OBJECT,
            description: "字体设置",
            properties: {
              bold: {
                name: "bold",
                type: ParameterType.BOOLEAN,
                description: "是否加粗",
                required: false,
              },
              color: {
                name: "color",
                type: ParameterType.STRING,
                description: "字体颜色，如#FF0000",
                required: false,
              },
              size: {
                name: "size",
                type: ParameterType.NUMBER,
                description: "字体大小",
                required: false,
              },
            },
          },
          fill: {
            type: ParameterType.OBJECT,
            description: "填充颜色",
            properties: {
              color: {
                name: "color",
                type: ParameterType.STRING,
                description: "填充颜色",
                required: false,
              },
            },
          },
          alignment: {
            type: ParameterType.OBJECT,
            description: "对齐方式",
            properties: {
              horizontal: {
                name: "horizontal",
                type: ParameterType.STRING,
                description: "水平对齐：left/center/right",
                required: false,
                enum: ["left", "center", "right"],
              },
              vertical: {
                name: "vertical",
                type: ParameterType.STRING,
                description: "垂直对齐：top/center/bottom",
                required: false,
                enum: ["top", "center", "bottom"],
              },
            },
          },
        },
      },
    ],
    returns: "void",
    example: {
      input: {
        rangeAddress: "A1:A10",
        format: {
          font: { bold: true, color: "#FF0000", size: 12 },
          fill: { color: "#FFFF00" },
          alignment: { horizontal: "center", vertical: "center" },
        },
      },
      description: "将A1:A10范围格式化为红色加粗、黄色背景、居中对齐",
    },
  },
  {
    id: "excel.create_chart",
    name: "create_chart",
    category: ToolCategory.EXCEL_OPERATION,
    description: "创建图表",
    parameters: [
      {
        name: "chartType",
        type: ParameterType.STRING,
        description: "图表类型",
        required: true,
        enum: ["ColumnClustered", "Line", "Pie", "Bar", "Area"],
      },
      {
        name: "dataRange",
        type: ParameterType.STRING,
        description: "数据范围地址",
        required: true,
      },
      {
        name: "title",
        type: ParameterType.STRING,
        description: "图表标题",
        required: false,
      },
      {
        name: "position",
        type: ParameterType.STRING,
        description: "图表位置单元格",
        required: false,
      },
    ],
    returns: "void",
    example: {
      input: {
        chartType: "ColumnClustered",
        dataRange: "A1:B10",
        title: "销售数据",
        position: "D1",
      },
      description: '基于A1:B10数据创建柱状图，标题为"销售数据"，放置在D1位置',
    },
  },
];

/**
 * 工作表操作工具
 */
export const WORKSHEET_TOOLS: ToolDefinition[] = [
  {
    id: "worksheet.add",
    name: "add_worksheet",
    category: ToolCategory.WORKSHEET_OPERATION,
    description: "添加新工作表",
    parameters: [
      {
        name: "name",
        type: ParameterType.STRING,
        description: "工作表名称",
        required: true,
        validation: {
          minLength: 1,
          maxLength: 31,
          errorMessage: "工作表名称长度必须在1-31个字符之间",
        },
      },
    ],
    returns: "void",
    example: {
      input: { name: "2025年数据" },
      description: '添加名为"2025年数据"的工作表',
    },
  },
  {
    id: "worksheet.delete",
    name: "delete_worksheet",
    category: ToolCategory.WORKSHEET_OPERATION,
    description: "删除工作表",
    parameters: [
      {
        name: "name",
        type: ParameterType.STRING,
        description: "工作表名称",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { name: "临时数据" },
      description: '删除名为"临时数据"的工作表',
    },
  },
  {
    id: "worksheet.rename",
    name: "rename_worksheet",
    category: ToolCategory.WORKSHEET_OPERATION,
    description: "重命名工作表",
    parameters: [
      {
        name: "oldName",
        type: ParameterType.STRING,
        description: "原工作表名称",
        required: true,
      },
      {
        name: "newName",
        type: ParameterType.STRING,
        description: "新工作表名称",
        required: true,
        validation: {
          minLength: 1,
          maxLength: 31,
          errorMessage: "工作表名称长度必须在1-31个字符之间",
        },
      },
    ],
    returns: "void",
    example: {
      input: { oldName: "Sheet1", newName: "主数据" },
      description: '将Sheet1重命名为"主数据"',
    },
  },
];

/**
 * 单元格操作扩展工具
 */
export const CELL_OPERATION_TOOLS: ToolDefinition[] = [
  {
    id: "excel.merge_cells",
    name: "merge_cells",
    category: ToolCategory.EXCEL_OPERATION,
    description: "合并单元格",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "要合并的单元格范围，如A1:C3",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "A1:C1" },
      description: "合并A1到C1的单元格",
    },
  },
  {
    id: "excel.unmerge_cells",
    name: "unmerge_cells",
    category: ToolCategory.EXCEL_OPERATION,
    description: "取消合并单元格",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "已合并的单元格范围",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "A1:C1" },
      description: "取消A1到C1的合并",
    },
  },
  {
    id: "excel.insert_rows",
    name: "insert_rows",
    category: ToolCategory.EXCEL_OPERATION,
    description: "插入行",
    parameters: [
      {
        name: "startRow",
        type: ParameterType.NUMBER,
        description: "起始行号",
        required: true,
      },
      {
        name: "count",
        type: ParameterType.NUMBER,
        description: "要插入的行数，默认为1",
        required: false,
      },
    ],
    returns: "void",
    example: {
      input: { startRow: 5, count: 3 },
      description: "从第5行开始插入3行",
    },
  },
  {
    id: "excel.insert_columns",
    name: "insert_columns",
    category: ToolCategory.EXCEL_OPERATION,
    description: "插入列",
    parameters: [
      {
        name: "startColumn",
        type: ParameterType.STRING,
        description: "起始列字母，如A",
        required: true,
      },
      {
        name: "count",
        type: ParameterType.NUMBER,
        description: "要插入的列数，默认为1",
        required: false,
      },
    ],
    returns: "void",
    example: {
      input: { startColumn: "C", count: 2 },
      description: "从C列开始插入2列",
    },
  },
  {
    id: "excel.delete_rows",
    name: "delete_rows",
    category: ToolCategory.EXCEL_OPERATION,
    description: "删除行",
    parameters: [
      {
        name: "startRow",
        type: ParameterType.NUMBER,
        description: "起始行号",
        required: true,
      },
      {
        name: "count",
        type: ParameterType.NUMBER,
        description: "要删除的行数，默认为1",
        required: false,
      },
    ],
    returns: "void",
    example: {
      input: { startRow: 5, count: 2 },
      description: "从第5行开始删除2行",
    },
  },
  {
    id: "excel.delete_columns",
    name: "delete_columns",
    category: ToolCategory.EXCEL_OPERATION,
    description: "删除列",
    parameters: [
      {
        name: "startColumn",
        type: ParameterType.STRING,
        description: "起始列字母",
        required: true,
      },
      {
        name: "count",
        type: ParameterType.NUMBER,
        description: "要删除的列数，默认为1",
        required: false,
      },
    ],
    returns: "void",
    example: {
      input: { startColumn: "D", count: 1 },
      description: "删除D列",
    },
  },
  {
    id: "excel.set_formula",
    name: "set_formula",
    category: ToolCategory.EXCEL_OPERATION,
    description: "设置单元格公式",
    parameters: [
      {
        name: "cellAddress",
        type: ParameterType.STRING,
        description: "单元格地址",
        required: true,
      },
      {
        name: "formula",
        type: ParameterType.STRING,
        description: "Excel公式，如=SUM(A1:A10)",
        required: true,
      },
    ],
    returns: "any",
    example: {
      input: { cellAddress: "B11", formula: "=SUM(B1:B10)" },
      description: "在B11设置求和公式",
    },
  },
  {
    id: "excel.copy_range",
    name: "copy_range",
    category: ToolCategory.EXCEL_OPERATION,
    description: "复制区域到新位置",
    parameters: [
      {
        name: "sourceRange",
        type: ParameterType.STRING,
        description: "源区域地址",
        required: true,
      },
      {
        name: "destinationRange",
        type: ParameterType.STRING,
        description: "目标区域地址",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { sourceRange: "A1:B10", destinationRange: "D1:E10" },
      description: "将A1:B10复制到D1:E10",
    },
  },
  {
    id: "excel.autofit_columns",
    name: "autofit_columns",
    category: ToolCategory.EXCEL_OPERATION,
    description: "自动调整列宽",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "A:D" },
      description: "自动调整A到D列的宽度",
    },
  },
  {
    id: "excel.autofit_rows",
    name: "autofit_rows",
    category: ToolCategory.EXCEL_OPERATION,
    description: "自动调整行高",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "1:10" },
      description: "自动调整第1到10行的高度",
    },
  },
];

/**
 * 条件格式工具
 */
export const CONDITIONAL_FORMAT_TOOLS: ToolDefinition[] = [
  {
    id: "excel.add_data_bars",
    name: "add_data_bars",
    category: ToolCategory.EXCEL_OPERATION,
    description: "添加数据条条件格式",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "应用数据条的单元格范围",
        required: true,
      },
      {
        name: "color",
        type: ParameterType.STRING,
        description: "数据条颜色，默认#0066CC",
        required: false,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "B2:B20", color: "#00AA00" },
      description: "在B2:B20添加绿色数据条",
    },
  },
  {
    id: "excel.add_color_scale",
    name: "add_color_scale",
    category: ToolCategory.EXCEL_OPERATION,
    description: "添加色阶条件格式",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "应用色阶的单元格范围",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "C2:C50" },
      description: "在C2:C50添加色阶格式",
    },
  },
  {
    id: "excel.add_icon_set",
    name: "add_icon_set",
    category: ToolCategory.EXCEL_OPERATION,
    description: "添加图标集条件格式",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "应用图标集的单元格范围",
        required: true,
      },
      {
        name: "iconStyle",
        type: ParameterType.STRING,
        description: "图标样式：ThreeArrows/ThreeFlags/ThreeTrafficLights1",
        required: false,
        enum: ["ThreeArrows", "ThreeFlags", "ThreeTrafficLights1"],
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "D2:D30", iconStyle: "ThreeArrows" },
      description: "在D2:D30添加箭头图标集",
    },
  },
  {
    id: "excel.clear_conditional_formats",
    name: "clear_conditional_formats",
    category: ToolCategory.EXCEL_OPERATION,
    description: "清除条件格式",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "要清除条件格式的范围",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "A1:Z100" },
      description: "清除A1:Z100范围的所有条件格式",
    },
  },
];

/**
 * 数据验证工具
 */
export const DATA_VALIDATION_TOOLS: ToolDefinition[] = [
  {
    id: "excel.add_dropdown_validation",
    name: "add_dropdown_validation",
    category: ToolCategory.EXCEL_OPERATION,
    description: "添加下拉列表数据验证",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "应用验证的单元格范围",
        required: true,
      },
      {
        name: "options",
        type: ParameterType.ARRAY,
        description: "下拉选项列表",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "E2:E50", options: ["已完成", "进行中", "未开始"] },
      description: "在E2:E50添加状态下拉列表",
    },
  },
  {
    id: "excel.add_number_validation",
    name: "add_number_validation",
    category: ToolCategory.EXCEL_OPERATION,
    description: "添加数值范围数据验证",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "应用验证的单元格范围",
        required: true,
      },
      {
        name: "min",
        type: ParameterType.NUMBER,
        description: "允许的最小值",
        required: true,
      },
      {
        name: "max",
        type: ParameterType.NUMBER,
        description: "允许的最大值",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "F2:F100", min: 0, max: 100 },
      description: "在F2:F100限制输入0到100的数值",
    },
  },
  {
    id: "excel.clear_data_validation",
    name: "clear_data_validation",
    category: ToolCategory.EXCEL_OPERATION,
    description: "清除数据验证",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "要清除验证的范围",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { rangeAddress: "A1:Z100" },
      description: "清除A1:Z100范围的数据验证",
    },
  },
];

/**
 * 命名区域工具
 */
export const NAMED_RANGE_TOOLS: ToolDefinition[] = [
  {
    id: "excel.create_named_range",
    name: "create_named_range",
    category: ToolCategory.EXCEL_OPERATION,
    description: "创建命名区域",
    parameters: [
      {
        name: "name",
        type: ParameterType.STRING,
        description: "区域名称",
        required: true,
      },
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围地址",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { name: "SalesData", rangeAddress: "A1:D100" },
      description: '将A1:D100命名为"SalesData"',
    },
  },
  {
    id: "excel.delete_named_range",
    name: "delete_named_range",
    category: ToolCategory.EXCEL_OPERATION,
    description: "删除命名区域",
    parameters: [
      {
        name: "name",
        type: ParameterType.STRING,
        description: "要删除的区域名称",
        required: true,
      },
    ],
    returns: "void",
    example: {
      input: { name: "OldData" },
      description: '删除名为"OldData"的命名区域',
    },
  },
];

/**
 * 数据分析工具
 */
export const ANALYSIS_TOOLS: ToolDefinition[] = [
  {
    id: "analysis.sum_range",
    name: "sum_range",
    category: ToolCategory.DATA_ANALYSIS,
    description: "计算单元格范围的总和",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围地址",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
          errorMessage: "请输入有效的单元格地址，如A1或A1:B10",
        },
      },
      {
        name: "resultCell",
        type: ParameterType.STRING,
        description: "结果存放的单元格",
        required: false,
        validation: {
          pattern: "^[A-Z]+[0-9]+$",
          errorMessage: "请输入有效的单元格地址，如A1",
        },
      },
    ],
    returns: "number",
    example: {
      input: { rangeAddress: "B2:B10", resultCell: "B11" },
      description: "计算B2:B10的总和，结果显示在B11",
    },
  },
  {
    id: "analysis.average_range",
    name: "average_range",
    category: ToolCategory.DATA_ANALYSIS,
    description: "计算单元格范围的平均值",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "单元格范围地址",
        required: true,
      },
      {
        name: "resultCell",
        type: ParameterType.STRING,
        description: "结果存放的单元格",
        required: false,
      },
    ],
    returns: "number",
    example: {
      input: { rangeAddress: "C2:C20", resultCell: "C21" },
      description: "计算C2:C20的平均值，结果显示在C21",
    },
  },
  {
    id: "analysis.sort_range",
    name: "sort_range",
    category: ToolCategory.DATA_ANALYSIS,
    description: "对单元格范围进行排序",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "要排序的范围",
        required: true,
      },
      {
        name: "keyColumn",
        type: ParameterType.NUMBER,
        description: "作为排序依据的列索引（从0开始）",
        required: true,
      },
      {
        name: "order",
        type: ParameterType.STRING,
        description: "排序顺序：asc/desc",
        required: false,
        enum: ["asc", "desc"],
      },
    ],
    returns: "void",
    example: {
      input: {
        rangeAddress: "A2:C50",
        keyColumn: 1,
        order: "desc",
      },
      description: "按第2列（B列）降序排序A2:C50范围",
    },
  },
  {
    id: "analysis.filter_range",
    name: "filter_range",
    category: ToolCategory.DATA_ANALYSIS,
    description: "对单元格范围应用筛选",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "要筛选的范围",
        required: true,
      },
      {
        name: "criteria",
        type: ParameterType.OBJECT,
        description: "筛选条件",
        required: true,
        properties: {
          column: {
            name: "column",
            type: ParameterType.NUMBER,
            description: "要筛选的列索引（从0开始）",
            required: true,
          },
          operator: {
            name: "operator",
            type: ParameterType.STRING,
            description: "比较运算符：equals/notEquals/greaterThan/lessThan/contains",
            required: true,
            enum: ["equals", "notEquals", "greaterThan", "lessThan", "contains"],
          },
          value: {
            name: "value",
            type: ParameterType.ANY,
            description: "比较值",
            required: true,
          },
        },
      },
    ],
    returns: "void",
    example: {
      input: {
        rangeAddress: "A2:D100",
        criteria: {
          column: 2,
          operator: "greaterThan",
          value: 1000,
        },
      },
      description: "筛选A2:D100范围，显示第3列大于1000的行",
    },
  },
];

/**
 * 上下文感知工具定义 - Excel 感知层
 */
export const CONTEXT_TOOLS: ToolDefinition[] = [
  {
    id: "context.get_workbook_summary",
    name: "get_workbook_summary",
    category: ToolCategory.DATA_ANALYSIS,
    description: "获取工作簿的完整结构摘要，包括所有工作表、表格、图表等信息",
    parameters: [],
    returns: "WorkbookSummary",
    example: {
      input: {},
      description: "获取当前工作簿的结构概览",
    },
  },
  {
    id: "context.get_selection_context",
    name: "get_selection_context",
    category: ToolCategory.DATA_ANALYSIS,
    description: "获取当前选区的详细上下文信息，包括值、公式、数据类型分布等",
    parameters: [],
    returns: "SelectionContext",
    example: {
      input: {},
      description: "分析当前选中的单元格范围",
    },
  },
  {
    id: "context.detect_headers",
    name: "detect_headers",
    category: ToolCategory.DATA_ANALYSIS,
    description: "自动检测数据区域的表头，返回表头信息和置信度",
    parameters: [
      {
        name: "rangeAddress",
        type: ParameterType.STRING,
        description: "要检测的单元格范围地址",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
          errorMessage: "请输入有效的单元格地址，如A1或A1:B10",
        },
      },
    ],
    returns: "HeaderDetectionResult",
    example: {
      input: { rangeAddress: "A1:E20" },
      description: "检测A1:E20范围是否有表头",
    },
  },
  {
    id: "context.get_formula_dependencies",
    name: "get_formula_dependencies",
    category: ToolCategory.DATA_ANALYSIS,
    description: "获取指定单元格的公式依赖关系，包括引用的单元格和是否为易失函数",
    parameters: [
      {
        name: "cellAddress",
        type: ParameterType.STRING,
        description: "要分析的单元格地址",
        required: true,
        validation: {
          pattern: "^[A-Z]+[0-9]+$",
          errorMessage: "请输入有效的单元格地址，如A1",
        },
      },
    ],
    returns: "FormulaDependencyInfo",
    example: {
      input: { cellAddress: "C10" },
      description: "分析C10单元格的公式依赖",
    },
  },
  {
    id: "context.get_all_tables",
    name: "get_all_tables",
    category: ToolCategory.DATA_ANALYSIS,
    description: "获取工作簿中所有Excel表格对象的信息",
    parameters: [],
    returns: "TableInfo[]",
    example: {
      input: {},
      description: "列出所有Excel表格",
    },
  },
  {
    id: "context.get_named_ranges",
    name: "get_named_ranges",
    category: ToolCategory.DATA_ANALYSIS,
    description: "获取工作簿中所有命名范围的信息",
    parameters: [],
    returns: "NamedRangeInfo[]",
    example: {
      input: {},
      description: "列出所有命名范围",
    },
  },
];

/**
 * 获取所有已注册的工具
 */
export function getAllTools(): ToolDefinition[] {
  return [
    ...EXCEL_TOOLS,
    ...WORKSHEET_TOOLS,
    ...CELL_OPERATION_TOOLS,
    ...CONDITIONAL_FORMAT_TOOLS,
    ...DATA_VALIDATION_TOOLS,
    ...NAMED_RANGE_TOOLS,
    ...ANALYSIS_TOOLS,
    ...CONTEXT_TOOLS,
  ];
}

/**
 * 按ID查找工具
 */
export function getToolById(id: string): ToolDefinition | undefined {
  return getAllTools().find((tool) => tool.id === id);
}

/**
 * 按类别获取工具
 */
export function getToolsByCategory(category: ToolCategory): ToolDefinition[] {
  return getAllTools().filter((tool) => tool.category === category);
}

/**
 * 验证工具参数
 */
export function validateToolParameters(
  toolId: string,
  parameters: Record<string, any>
): { isValid: boolean; errors: string[] } {
  const tool = getToolById(toolId);
  if (!tool) {
    return {
      isValid: false,
      errors: [`工具 ${toolId} 未注册`],
    };
  }

  const errors: string[] = [];

  // 检查必需参数
  for (const param of tool.parameters) {
    if (param.required && !(param.name in parameters)) {
      errors.push(`缺少必需参数: ${param.name}`);
    }
  }

  // 验证参数类型和格式
  for (const [key, value] of Object.entries(parameters)) {
    const paramDef = tool.parameters.find((p) => p.name === key);
    if (!paramDef) {
      errors.push(`未知参数: ${key}`);
      continue;
    }

    // 类型检查
    if (!validateParameterType(value, paramDef.type)) {
      errors.push(`参数 ${key} 类型错误，期望 ${paramDef.type}`);
    }

    // 验证规则检查
    if (paramDef.validation) {
      const validationError = validateParameter(value, paramDef.validation);
      if (validationError) {
        errors.push(`参数 ${key} 验证失败: ${validationError}`);
      }
    }

    // 枚举值检查
    if (paramDef.enum && !paramDef.enum.includes(value)) {
      errors.push(`参数 ${key} 必须是以下值之一: ${paramDef.enum.join(", ")}`);
    }
  }

  return {
    isValid: errors.length === 0,
    errors,
  };
}

/**
 * 参数类型验证
 */
function validateParameterType(value: any, expectedType: ParameterType): boolean {
  switch (expectedType) {
    case ParameterType.STRING:
      return typeof value === "string";
    case ParameterType.NUMBER:
      return typeof value === "number";
    case ParameterType.BOOLEAN:
      return typeof value === "boolean";
    case ParameterType.ARRAY:
      return Array.isArray(value);
    case ParameterType.OBJECT:
      return typeof value === "object" && value !== null && !Array.isArray(value);
    case ParameterType.ANY:
      return true;
    default:
      return false;
  }
}

/**
 * 参数验证规则检查
 */
function validateParameter(value: any, validation: ParameterValidationRule): string | null {
  if (validation.pattern && typeof value === "string") {
    const regex = new RegExp(validation.pattern);
    if (!regex.test(value)) {
      return validation.errorMessage || "格式不符合要求";
    }
  }

  if (validation.minLength && typeof value === "string" && value.length < validation.minLength) {
    return validation.errorMessage || `长度不能少于 ${validation.minLength} 个字符`;
  }

  if (validation.maxLength && typeof value === "string" && value.length > validation.maxLength) {
    return validation.errorMessage || `长度不能超过 ${validation.maxLength} 个字符`;
  }

  if (validation.minItems && Array.isArray(value) && value.length < validation.minItems) {
    return validation.errorMessage || `至少需要 ${validation.minItems} 个元素`;
  }

  if (validation.maxItems && Array.isArray(value) && value.length > validation.maxItems) {
    return validation.errorMessage || `最多允许 ${validation.maxItems} 个元素`;
  }

  if (validation.minValue && typeof value === "number" && value < validation.minValue) {
    return validation.errorMessage || `值不能小于 ${validation.minValue}`;
  }

  if (validation.maxValue && typeof value === "number" && value > validation.maxValue) {
    return validation.errorMessage || `值不能大于 ${validation.maxValue}`;
  }

  return null;
}
