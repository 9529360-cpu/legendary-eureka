# sv-excel-agent 学习笔记

> 来源：https://github.com/Sylvian/sv-excel-agent  
> 项目类型：Python MCP Server + LLM Agent  
> 与我们项目相关度：**★★★★★ 极高**

## 项目概述

这是一个开源的 Excel AI Agent，由两部分组成：
1. **excel_mcp** - MCP Server，提供 ~30 个 Excel 操作工具
2. **excel_agent** - Agent Runner，使用 pydantic-ai 框架

### 技术栈对比

| 特性 | sv-excel-agent | 我们的项目 |
|------|---------------|-----------|
| 后端语言 | Python | TypeScript/Node.js |
| Excel 操作 | openpyxl (文件级) | Office.js API (运行时) |
| Agent 框架 | pydantic-ai | 自研 ReAct |
| 工具协议 | MCP (FastMCP) | 自定义 Tool 协议 |
| 运行环境 | 独立 Python 进程 | Office Add-in (浏览器沙箱) |

---

## 可借鉴的设计模式

### 1. ToolError/ToolSuccess 统一响应格式

**文件**: `excel_mcp/errors.py`

```python
class ToolError(Exception):
    """自动转换为 JSON 错误响应"""
    
    def __init__(self, message: str, code: str | None = None, **kwargs):
        self.message = message
        self.code = code
        self.extra = kwargs
        
    def to_json(self) -> str:
        result = {"status": "error", "error": self.message}
        if self.code:
            result["code"] = self.code
        result.update(self.extra)
        return json.dumps(result)


class ToolSuccess:
    """统一成功响应"""
    
    def __init__(self, data: dict | None = None, **kwargs):
        self.data = {"status": "success"}
        if data:
            self.data.update(data)
        self.data.update(kwargs)
        
    def to_json(self) -> str:
        return json.dumps(self.data)
```

**借鉴价值**：我们的 `ToolResult` 可以参考这种简洁的设计。

---

### 2. 会话管理模式

**文件**: `excel_mcp/sessions.py`

```python
class SessionManager:
    """管理用户会话，映射 user_id 到 UserState"""
    
    def __init__(self):
        self._sessions: dict[str, UserState] = {}
    
    def get_session(self, user_id: str) -> UserState:
        if user_id not in self._sessions:
            self._sessions[user_id] = UserState(user_id=user_id)
        return self._sessions[user_id]
    
    def remove_session(self, user_id: str) -> bool:
        if user_id in self._sessions:
            state = self._sessions[user_id]
            if state.workbook is not None:
                state.workbook.close()
            del self._sessions[user_id]
            return True
        return False

# 装饰器模式注入会话
@with_workbook
def get_range_data(mwb: ManagedWorkbook, operation: GetRangeDataOperation) -> str:
    ...

@with_workbook_mutation  # 标记为会修改数据
def set_range_data(mwb: ManagedWorkbook, operation: SetRangeDataOperation) -> str:
    ...
```

**借鉴价值**：装饰器区分 **只读** vs **修改** 操作，自动处理脏标记。

---

### 3. Pydantic Schema 驱动的工具参数

**文件**: `excel_mcp/schemas.py`

```python
class SetRangeDataOperation(BaseModel):
    range_a1: Annotated[
        str,
        Field(description="Range in A1 notation (e.g., 'A1', 'B2:C3', 'Sheet1!A1:B2')")
    ]
    value: Annotated[
        CellValueOrList,
        Field(description="Value(s) to set. Single value, list for row/column, or nested list for grid")
    ]

class StyleOptions(BaseModel):
    bold: Annotated[bool | None, Field(description="Set bold text")] = None
    italic: Annotated[bool | None, Field(description="Set italic text")] = None
    font_size: Annotated[int | None, Field(description="Font size in points")] = None
    font_color: Annotated[str | None, Field(description="Font color as hex")] = None
    bg_color: Annotated[str | None, Field(description="Background color as hex")] = None
    h_align: Annotated[Literal["left", "center", "right"] | None, ...] = None
```

**借鉴价值**：使用 Pydantic + Annotated 生成工具的 JSON Schema，LLM 可以直接理解参数格式。

---

### 4. 公式翻译函数

**文件**: `excel_mcp/helpers.py`

```python
def translate_formula(formula: str, row_offset: int, col_offset: int) -> str:
    """
    翻译公式中的单元格引用，处理：
    - 相对引用 (A1) - 按偏移调整
    - 绝对行 ($A1) - 列调整，行固定
    - 绝对列 (A$1) - 行调整，列固定
    - 完全绝对 ($A$1) - 不调整
    """
    cell_ref_pattern = r"(\$?)([A-Z]+)(\$?)(\d+)"
    
    def replace_ref(match):
        col_absolute = match.group(1) == "$"
        col_letters = match.group(2)
        row_absolute = match.group(3) == "$"
        row_num = int(match.group(4))
        
        if not col_absolute:
            col_idx = column_index_from_string(col_letters)
            col_idx += col_offset
            col_letters = get_column_letter(col_idx)
        
        if not row_absolute:
            row_num += row_offset
        
        return f"{'$' if col_absolute else ''}{col_letters}{'$' if row_absolute else ''}{row_num}"
    
    return re.sub(cell_ref_pattern, replace_ref, formula, flags=re.IGNORECASE)
```

**借鉴价值**：实现 auto_fill 功能时需要的公式引用偏移逻辑。

---

### 5. 数组公式检测

**文件**: `excel_mcp/helpers.py`

```python
def is_array_formula(formula: str) -> bool:
    """
    检测公式是否需要作为数组公式 (CSE) 输入。
    
    常见模式：
    - MATCH(1, (range=value)*(range=value), 0) - 数组 AND 逻辑
    - INDEX(..., MATCH(1, ...*..., 0)) - 带数组匹配的查找
    
    注意：SUMPRODUCT 原生处理数组，不需要 CSE。
    """
    if not formula.startswith("="):
        return False
    
    formula_upper = formula.upper()
    
    # SUMPRODUCT 原生处理数组
    if formula_upper.startswith("=SUMPRODUCT"):
        return False
    
    # 模式：MATCH(1, ...) 内部有乘法 - 数组 AND 逻辑
    if "MATCH(1," in formula_upper:
        if re.search(r"\([^)]*\$?[A-Z]+\$?\d*:\$?[A-Z]+\$?\d*[^)]*\)\s*\*\s*\(", formula):
            return True
    
    return False
```

**借鉴价值**：智能识别需要特殊处理的公式类型。

---

### 6. 批量操作优化

**文件**: `excel_mcp/excel_server.py`

```python
@mcp.tool()
async def batch_set_range_data(operations: list[SetRangeOperation]) -> str:
    """单次调用设置多个区域，比多次调用 set_range_data 更高效"""
    cells_set = []
    for op in operations:
        cells_set.extend(set_range_values(mwb.wb, op.range_a1, op.value))
    
    cell_results = await get_cell_results(mwb, cells_set)
    return ToolSuccess(cells=cell_results).to_json()


@mcp.tool()
def batch_delete_rows(operations: list[DeleteRowsOperation]) -> str:
    """批量删除行 - 按位置倒序排列避免索引偏移问题"""
    sorted_ops = sorted(operations, key=lambda x: x.position, reverse=True)
    for op in sorted_ops:
        sheet.delete_rows(op.position, amount=op.count)
    return ToolSuccess().to_json()
```

**借鉴价值**：
1. 提供 batch 版本的工具减少 LLM 调用次数
2. 删除操作倒序执行避免索引偏移

---

### 7. 表格格式化输出

**文件**: `excel_mcp/formatting.py`

```python
def format_sheet_table(sheet, cached_sheet=None, max_rows=50, width=300) -> str:
    """将工作表格式化为 Rich 表格输出"""
    
    table = Table(
        show_header=True,
        header_style="bold cyan",
        border_style="dim",
    )
    
    # 行号列
    table.add_column("#", style="dim", justify="right", width=4)
    
    # 数据列 (A, B, C, ...)
    for i in range(num_cols):
        table.add_column(get_column_letter(i + 1), justify="left")
    
    # 公式单元格显示：=A1+B1 [5]
    for cell in row:
        if is_formula and cached_value:
            display = f"{formula} [{cached_value}]"
```

**借鉴价值**：向 LLM 展示数据时，公式同时显示计算结果，便于理解。

---

### 8. System Prompt 设计

**文件**: `excel_agent/config.py`

```python
SYSTEM_PROMPT = """You are Sylvian, an expert Excel spreadsheet assistant...

RULES:
    1. ONLY update/format the cells required for the task; avoid re-formatting unless specified.
    2. Do not adjust cells for the purposes of testing.
    3. Do not exit early or ask user of input, just perform the task.
    4. Always check the output of tool calls to ensure the cells are updated correctly.
    5. After performing the task, ALWAYS EXAMINE YOUR OUTPUT AND ENSURE IT IS CORRECT.

<solution_persistence>
- Treat yourself as an autonomous senior pair-programmer
- Persist until the task is fully handled end-to-end
- Be extremely biased for action. If the user asks "should we do x?" and answer is "yes", also perform the action.
</solution_persistence>
"""
```

**借鉴价值**：强调"最小修改原则"和"行动偏向"。

---

### 9. RetryPrompt 转换

**文件**: `excel_agent/agent_runner.py`

```python
def retry_prompt_to_user_message(messages: list) -> list:
    """
    将 RetryPrompt 部分转换为普通用户消息，
    让模型能看到验证错误详情并自我纠正。
    """
    for message in messages:
        if isinstance(message, ModelRequest):
            for part in message.parts:
                if isinstance(part, RetryPromptPart):
                    error_text = part.model_response()
                    user_message = (
                        f"Tool '{part.tool_name}' failed validation.\n\n"
                        f"{error_text}\n\n"
                        "Please fix the tool arguments and issue another tool call."
                    )
                    new_parts.append(UserPromptPart(content=user_message))
```

**借鉴价值**：工具调用失败时，将错误信息作为用户消息反馈给 LLM，让其自我纠正。

---

## 工具清单（30+ 个）

### 数据读写
- `get_range_data` - 获取区域数据
- `set_range_data` - 设置单个区域
- `batch_set_range_data` - 批量设置多个区域
- `auto_fill` - 自动填充（带公式翻译）
- `display_sheet` - 显示工作表内容

### 工作表管理
- `get_workbook_info` - 获取工作簿信息
- `get_sheets` - 获取所有工作表名
- `create_sheet` - 创建工作表
- `delete_sheet` - 删除工作表
- `rename_sheet` / `batch_rename_sheets` - 重命名

### 行列操作
- `insert_rows` / `insert_columns` - 插入行列
- `delete_rows` / `delete_columns` - 删除行列
- `batch_insert_rows` / `batch_delete_rows` - 批量操作

### 样式格式
- `set_range_style` - 设置区域样式
- `batch_set_styles` - 批量设置样式
- `set_merge` / `unmerge` - 合并/取消合并单元格

### 条件格式
- `add_conditional_formatting_rule` - 添加条件格式
- `delete_conditional_formatting_rule` - 删除条件格式
- `get_conditional_formatting_rules` - 获取条件格式列表

### 数据验证
- `add_data_validation_rule` - 添加数据验证（下拉框、日期选择等）
- `delete_data_validation_rule` - 删除数据验证
- `get_data_validation_rules` - 获取验证规则

### 搜索
- `search_cells` - 按值或公式搜索单元格

---

## 应用到我们项目的建议

### 高优先级

1. **统一 ToolResult 格式** - 参考 ToolError/ToolSuccess 模式
2. **batch 版本工具** - 减少 LLM 调用，提高效率
3. **公式翻译函数** - 实现 auto_fill 功能
4. **@with_workbook 装饰器** - 区分只读/修改操作

### 中优先级

5. **Pydantic Schema** - 使用 Zod 生成工具参数 Schema
6. **表格格式化输出** - 公式显示计算值 `=A1+B1 [5]`
7. **System Prompt 优化** - 加入"最小修改原则"

### 低优先级

8. **条件格式工具** - 当前我们暂未实现
9. **数据验证工具** - 下拉框、复选框等

---

## 关键差异点

| 方面 | sv-excel-agent | 我们的项目 | 备注 |
|------|---------------|-----------|------|
| Excel 操作 | openpyxl (文件) | Office.js (运行时) | 我们可以直接操作打开的 Excel |
| 公式计算 | 需调用 Excel 进程 | 自动计算 | 我们的优势 |
| 并发 | 单会话 | 单会话 | 类似 |
| 工具数量 | ~30 | ~90 | 我们更多 |

**结论**：这个项目的设计模式非常值得学习，但核心 Excel 操作代码无法直接复用（技术栈不同）。
