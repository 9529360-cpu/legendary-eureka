import { parseLlmOutput } from "./llmOutputParser";

describe("llmOutputParser", () => {
  it("parses clean JSON", () => {
    const input = '{"a":1,"b":2}';
    const res = parseLlmOutput(input);
    expect(res.ok).toBe(true);
    expect(res.data).toEqual({ a: 1, b: 2 });
  });

  it("parses JSON inside text", () => {
    const input = 'Here is the result:\n```json\n{"x":true, "y": [1,2,3]}\n```\nThanks';
    const res = parseLlmOutput(input);
    expect(res.ok).toBe(true);
    expect(res.data.x).toBe(true);
    expect(res.data.y).toEqual([1, 2, 3]);
  });

  it("recovers single quotes and trailing commas", () => {
    const input = `{
  'name': 'test',
  'items': [1,2,],
}`;
    const res = parseLlmOutput(input);
    expect(res.ok).toBe(true);
    expect(res.data.name).toBe("test");
    expect(res.data.items).toEqual([1, 2]);
  });

  it("fails gracefully for non-json", () => {
    const input = "I think the answer is forty-two.";
    const res = parseLlmOutput(input);
    expect(res.ok).toBe(false);
    expect(res.error).toBeDefined();
  });
});
