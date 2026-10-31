import { describe, expect, it, vi } from "vitest";
import type { TemplateRequirement } from "../types";
import { SampleJson } from "./sampleJson";

const requirement = (name: string, type: TemplateRequirement["type"], fields?: TemplateRequirement[]): TemplateRequirement => ({ name, type, ...(fields ? { fields } : {}) });
const scalar = [requirement("Name", "value"), requirement("Active", "boolean")];
const empty = { itemsPerCollection: 10, values: "empty" as const };

describe("sample JSON conversion", () => {
  it("preserves order, recursive shapes and Empty values with the count at every nesting level", () => {
    const random = vi.fn();
    const requirements = [requirement("Customer", "object", scalar), requirement("Orders", "array", [
      requirement("Number", "value"), requirement("Paid", "boolean"), requirement("Items", "array", scalar)
    ])];
    expect(new SampleJson(random).generate(requirements, { ...empty, itemsPerCollection: 2 })).toBe(JSON.stringify({
      Customer: { Name: "", Active: false },
      Orders: Array.from({ length: 2 }, () => ({ Number: "", Paid: false, Items: [{ Name: "", Active: false }, { Name: "", Active: false }] }))
    }, null, 2));
    expect(random).not.toHaveBeenCalled();
  });

  it("uses injected randomness for generic strings and booleans without inventing scalar types", () => {
    const random = vi.fn().mockReturnValueOnce(0.25).mockReturnValueOnce(0.75).mockReturnValueOnce(0.5).mockReturnValueOnce(0);
    const generated = JSON.parse(new SampleJson(random).generate([requirement("Rows", "array", scalar)], { itemsPerCollection: 2, values: "random" }));
    expect(generated).toEqual({ Rows: [{ Name: "Sample 250000", Active: true }, { Name: "Sample 500000", Active: false }] });
    expect(random).toHaveBeenCalledTimes(4);
  });

  it("keeps unknown-shape arrays empty in both modes, including very large valid counts", () => {
    for (const values of ["empty", "random"] as const) {
      expect(new SampleJson(Math.random).generate([requirement("Tags", "array")], { itemsPerCollection: 1e100, values })).toBe(JSON.stringify({ Tags: [] }, null, 2));
    }
  });

  it("returns an empty object for no dependencies", () => {
    expect(new SampleJson(Math.random).generate([], empty)).toBe("{}");
  });

  it.each([0, -1, 1.5, NaN, Infinity])("rejects invalid count %s", itemsPerCollection => {
    expect(() => new SampleJson(Math.random).generate(scalar, { ...empty, itemsPerCollection })).toThrow("positive integer");
  });

  it("allows the exact aggregate boundary and collection counts above the earlier proposed maximum", () => {
    const generated = JSON.parse(new SampleJson(Math.random).generate([requirement("Rows", "array", scalar)], { ...empty, itemsPerCollection: 100_000 }));
    expect(generated.Rows).toHaveLength(100_000);
  });

  it("rejects oversized sibling, object-contained, nested and huge collections before producing any values", () => {
    const random = vi.fn();
    const generator = new SampleJson(random);
    const rows = requirement("Rows", "array", scalar);
    for (const [requirements, count] of [
      [[rows], 100_001],
      [[rows, requirement("Other", "object", [rows])], 50_001],
      [[requirement("Outer", "array", [rows])], 316],
      [[rows], 1e100]
    ] as [TemplateRequirement[], number][]) {
      expect(() => generator.generate(requirements, { itemsPerCollection: count, values: "random" })).toThrow("100,000 array elements in total");
    }
    expect(random).not.toHaveBeenCalled();
  });

  it("counts all nesting levels rather than only leaves", () => {
    const rows = requirement("Rows", "array", [requirement("Items", "array", [requirement("Value", "value")])]);
    const generated = JSON.parse(new SampleJson(Math.random).generate([rows], { ...empty, itemsPerCollection: 315 }));
    expect(generated.Rows).toHaveLength(315);
    expect(generated.Rows[314].Items).toHaveLength(315);
    // 316² is below the limit, but 316 + 316² is over it.
    expect(() => new SampleJson(Math.random).generate([rows], { ...empty, itemsPerCollection: 316 })).toThrow("100,000");
  });
});
