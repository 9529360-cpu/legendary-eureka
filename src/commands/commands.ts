/**
 * Excel 智能助手 commands (Excel-safe, lint/prettier-safe)
 * - No mailbox APIs (Outlook-only)
 * - Console notify only (never overwrite user cells)
 * - Same-origin backend call: /chat (needs webpack devServer proxy)
 */

export {}; // Make this file a module to avoid global redeclare issues.

Office.onReady(() => {
  console.log("Excel 智能助手 commands loaded.");
});

// ----------------------
// Notify helpers (Excel-safe)
// ----------------------
function notify(message: string) {
  console.log(`[ExcelCopilot] ${message}`);
}

function notifyError(prefix: string, error: unknown) {
  const msg = error instanceof Error ? error.message : String(error);
  console.error(`[ExcelCopilot][ERROR] ${prefix}: ${msg}`);
}

// ----------------------
// Backend helper (same-origin)
// ----------------------
async function postJson<T = unknown>(url: string, data: unknown): Promise<T> {
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });

  const text = await resp.text();
  let parsed: Record<string, unknown> = {};
  try {
    parsed = text ? JSON.parse(text) : {};
  } catch {
    parsed = { raw: text };
  }

  if (!resp.ok) {
    throw new Error((parsed?.error as string) || `POST ${url} failed: ${resp.status}`);
  }
  return parsed as T;
}

// ----------------------
// Utilities
// ----------------------
async function ensureUniqueSheetName(
  context: Excel.RequestContext,
  baseName: string
): Promise<string> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  const names = new Set(sheets.items.map((s) => s.name));
  if (!names.has(baseName)) return baseName;

  let i = 2;
  while (names.has(`${baseName} (${i})`)) i++;
  return `${baseName} (${i})`;
}

// ----------------------
// Commands
// ----------------------
function action(event: Office.AddinCommands.Event) {
  try {
    notify("Excel 智能助手 is ready");
  } catch (e) {
    notifyError("Command failed", e);
  } finally {
    event.completed();
  }
}

function cleanAndFormatData(event: Office.AddinCommands.Event) {
  Excel.run(async (context) => {
    try {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "columnCount", "values"]);
      await context.sync();

      notify(`Cleaning & formatting ${range.address}...`);

      range.clear(Excel.ClearApplyTo.formats);
      range.format.autofitColumns();
      range.format.autofitRows();

      const edges = ["EdgeBottom", "EdgeTop", "EdgeLeft", "EdgeRight"] as const;
      edges.forEach((edge) => {
        range.format.borders.getItem(edge).style = Excel.BorderLineStyle.continuous;
      });

      const values = range.values || [];
      let numeric = 0;
      let total = 0;

      for (const row of values) {
        for (const v of row) {
          if (v === null || v === "" || v === undefined) continue;
          total++;
          if (typeof v === "number") numeric++;
        }
      }

      if (total > 0 && numeric / total >= 0.6) {
        range.numberFormat = [["#,##0.00"]];
      }

      await context.sync();
      notify(`Done (${range.address})`);
    } catch (e) {
      notifyError("Clean & Format failed", e);
    } finally {
      event.completed();
    }
  }).catch((e) => {
    notifyError("Excel.run failed", e);
    event.completed();
  });
}

function generateInsights(event: Office.AddinCommands.Event) {
  Excel.run(async (context) => {
    try {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values", "rowCount", "columnCount"]);
      await context.sync();

      notify(`Generating insights for ${range.address}...`);

      const values = range.values || [];
      const rowCount = Math.max(1, range.rowCount);
      const colCount = Math.max(1, range.columnCount);

      const sample = values.slice(0, Math.min(20, values.length));
      const aiPrompt =
        `请对我在Excel中选中的区域 ${range.address} 做数据洞察。\n` +
        `数据行数=${rowCount}，列数=${colCount}。\n` +
        `下面是最多20行的样例数据(JSON二维数组)：\n` +
        JSON.stringify(sample) +
        `\n请输出简短结论/建议（中文），不要输出代码。`;

      let insightText = "";

      // 1) Try backend
      try {
        interface AIResponse {
          message?: string;
          explanation?: string;
        }
        const ai = await postJson<AIResponse>("/chat", {
          message: aiPrompt,
          context: {
            source: "commands.generateInsights",
            range: range.address,
          },
        });

        insightText =
          (ai && (ai.message || ai.explanation)) ||
          "AI已返回结果，但格式不符合预期（已降级显示）。";
      } catch (aiErr) {
        // 2) Local fallback
        let numericCount = 0;
        let textCount = 0;
        let emptyCount = 0;
        const nums: number[] = [];

        for (let i = 0; i < rowCount; i++) {
          for (let j = 0; j < colCount; j++) {
            const v = values?.[i]?.[j];
            if (v === null || v === undefined || v === "") emptyCount++;
            else if (typeof v === "number") {
              numericCount++;
              nums.push(v);
            } else {
              textCount++;
            }
          }
        }

        insightText =
          `(本地分析，AI后端不可用)\n` +
          `区域：${range.address}\n` +
          `维度：${rowCount}行 × ${colCount}列\n` +
          `数字：${numericCount}，文本：${textCount}，空值：${emptyCount}`;

        if (nums.length > 0) {
          const sum = nums.reduce((a, b) => a + b, 0);
          const avg = sum / nums.length;
          const max = Math.max(...nums);
          const min = Math.min(...nums);
          insightText +=
            `\n数值范围：${min.toFixed(2)} ~ ${max.toFixed(2)}` +
            `，均值：${avg.toFixed(2)}，总和：${sum.toFixed(2)}`;
        }

        console.warn("AI backend failed, fallback used:", aiErr);
      }

      // Write into new column to the right of selection
      const headerCell = range.getCell(0, colCount);
      headerCell.values = [["Insights"]];
      headerCell.format.fill.color = "#FFEB9C";
      headerCell.format.font.bold = true;

      // Insight on next row (do not overwrite header)
      const insightCell = headerCell.getOffsetRange(1, 0);
      insightCell.values = [[insightText]];
      insightCell.format.wrapText = true;

      // Fill "-" starting from row 3 (offset 2) for remaining selection rows
      if (rowCount > 2) {
        const dashRange = headerCell.getOffsetRange(2, 0).getResizedRange(rowCount - 3, 0);
        dashRange.values = Array.from({ length: rowCount - 2 }, () => ["-"]);
      }

      // Autofit at least header + insight (2 rows)
      const fullCol = headerCell.getResizedRange(Math.max(2, rowCount) - 1, 0);
      fullCol.format.autofitColumns();

      await context.sync();
      notify(`Insights added (${range.address} -> new column)`);
    } catch (e) {
      notifyError("Generate Insights failed", e);
    } finally {
      event.completed();
    }
  }).catch((e) => {
    notifyError("Excel.run failed", e);
    event.completed();
  });
}

function convertToTable(event: Office.AddinCommands.Event) {
  Excel.run(async (context) => {
    try {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();

      notify(`Converting ${range.address} to table...`);

      const table = context.workbook.tables.add(range, true);
      table.name = `Table_${Date.now()}`;
      table.load("name");

      table.style = "TableStyleMedium2";
      const header = table.getHeaderRowRange();
      header.format.fill.color = "#4472C4";
      header.format.font.color = "white";
      header.format.font.bold = true;

      await context.sync();
      notify(`Table created (${table.name})`);
    } catch (e) {
      notifyError("Convert to Table failed", e);
    } finally {
      event.completed();
    }
  }).catch((e) => {
    notifyError("Excel.run failed", e);
    event.completed();
  });
}

function createSummarySheet(event: Office.AddinCommands.Event) {
  Excel.run(async (context) => {
    try {
      notify("Creating summary sheet...");

      const name = await ensureUniqueSheetName(context, "Summary");
      const sheet = context.workbook.worksheets.add(name);
      sheet.position = 0;

      const title = sheet.getRange("A1");
      title.values = [["Excel 智能助手 Summary"]];
      title.format.font.size = 18;
      title.format.font.bold = true;
      title.format.fill.color = "#4472C4";
      title.format.font.color = "white";

      const ts = sheet.getRange("A2");
      ts.values = [[`Generated: ${new Date().toLocaleString()}`]];
      ts.format.font.italic = true;

      const hdr = sheet.getRange("A4");
      hdr.values = [["Workbook Statistics:"]];
      hdr.format.font.bold = true;

      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const usedRanges: Array<{ sheet: Excel.Worksheet; used: Excel.Range }> = [];
      for (const s of sheets.items) {
        if (s.name === name) continue;
        const used = s.getUsedRangeOrNullObject();
        used.load(["address", "rowCount", "columnCount"]);
        usedRanges.push({ sheet: s, used });
      }
      await context.sync();

      let row = 5;
      for (const item of usedRanges) {
        if (item.used.isNullObject) continue;
        sheet.getRange(`A${row}`).values = [
          [
            `${item.sheet.name}: ${item.used.rowCount} rows x ${item.used.columnCount} (${item.used.address})`,
          ],
        ];
        row++;
      }

      sheet.getUsedRange().format.autofitColumns();
      await context.sync();

      notify(`Summary created (${name})`);
    } catch (e) {
      notifyError("Create Summary failed", e);
    } finally {
      event.completed();
    }
  }).catch((e) => {
    notifyError("Excel.run failed", e);
    event.completed();
  });
}

// Associate commands with manifest function names
Office.actions.associate("action", action);
Office.actions.associate("cleanAndFormatData", cleanAndFormatData);
Office.actions.associate("generateInsights", generateInsights);
Office.actions.associate("convertToTable", convertToTable);
Office.actions.associate("createSummarySheet", createSummarySheet);
