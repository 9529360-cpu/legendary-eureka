/**
 * llmOutputParser.ts
 * 辅助函数：从大模型返回的混乱文本中提取并解析出 JSON
 * 能处理：带有代码块的 JSON、前后杂文、单引号/双引号、转义字符、多重 JSON 嵌套、甚至常见小错误（尾逗号）
 */

function tryJsonParse(input: string): any | null {
  try {
    return JSON.parse(input);
  } catch (e) {
    return null;
  }
}

function stripCodeFences(text: string): string {
  return text.replace(/```[\s\S]*?```/g, (m) => {
    // remove surrounding ``` and optional language tag, keep inner content
    // match opening fence with optional language: ```json\n
    const inner = m.replace(/^```\w*\n?/, '').replace(/\n?```$/, '');
    return inner;
  });
}

function firstJsonLikeSegment(text: string): string | null {
  // Try to find first {...} or [...] group with balanced braces
  const startChars = ['{', '['];
  for (const start of startChars) {
    const stack: string[] = [];
    let started = false;
    let buf = "";
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (!started) {
        if (ch === start) {
          started = true;
          stack.push(ch);
          buf += ch;
        }
      } else {
        buf += ch;
        if (ch === '{' || ch === '[') stack.push(ch);
        else if (ch === '}' || ch === ']') {
          const last = stack.pop();
          if (!last) {
            // unbalanced
            started = false;
            buf = "";
            break;
          }
          if (stack.length === 0) {
            return buf;
          }
        }
      }
    }
  }
  return null;
}

function tidyJsonLike(text: string): string {
  // Replace single quotes for object keys/strings when safe
  let s = text.trim();
  // 如果文本并非以 JSON 起始符号开头，则可能包含前导标签（例如 "Result: {..}"）
  if (!s.startsWith('{') && !s.startsWith('[')) {
    s = s.replace(/^.*?:\s*/, '');
  }

  // Remove code fences
  s = stripCodeFences(s);

  // Remove trailing commas before closing braces/brackets
  s = s.replace(/,\s*(}|\])/g, '$1');

  return s;
}

export function parseLlmOutput(text: string): { ok: boolean; data?: any; error?: string } {
  if (!text || typeof text !== 'string') return { ok: false, error: 'empty input' };

  // 1. Try direct parse
  let direct = tryJsonParse(text);
  if (direct !== null) return { ok: true, data: direct };

  // 2. Remove code fences and try again
  const noFences = stripCodeFences(text);
  direct = tryJsonParse(noFences);
  if (direct !== null) return { ok: true, data: direct };

  // 3. Find first json-like segment
  let segment = firstJsonLikeSegment(noFences);
  if (!segment) {
    const objMatch = noFences.match(/\{[\s\S]*\}/);
    const arrMatch = noFences.match(/\[[\s\S]*\]/);
    segment = (objMatch && objMatch[0]) || (arrMatch && arrMatch[0]) || null;
  }
  if (segment) {
    // tidy common issues
    const tidied = tidyJsonLike(segment);
    const p = tryJsonParse(tidied) || tryJsonParse(tidied.replace(/'/g, '"'));
    if (p !== null) return { ok: true, data: p };
  }

  // 4. Try to recover by replacing single quotes and removing trailing commas globally
  const replaced = tidyJsonLike(noFences).replace(/'/g, '"');
  const cleaned = replaced.replace(/,\s*([}\]])/g, '$1');
  const finalTry = tryJsonParse(cleaned);
  if (finalTry !== null) return { ok: true, data: finalTry };

  return { ok: false, error: 'unable to parse JSON' };
}

export default parseLlmOutput;
