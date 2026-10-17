import { describe, expect, it } from "vitest";
import type { TemplateRequirement } from "../types";
import { requirementsToSampleJson } from "./sampleJson";

const requirement = (name: string, type: TemplateRequirement["type"], fields?: TemplateRequirement[]): TemplateRequirement => ({ name, type, ...(fields ? { fields } : {}) });

describe("sample JSON conversion", () => {
  it("preserves dependency order and recursive object and array shapes", () => {
    const requirements = [
      requirement("Customer", "object", [requirement("Name", "value")]),
      requirement("Orders", "array", [requirement("Number", "value"), requirement("Paid", "boolean"), requirement("Items", "array", [requirement("Product", "value")])])
    ];

    expect(requirementsToSampleJson(requirements)).toBe(JSON.stringify({
      Customer: { Name: "" },
      Orders: [{ Number: "", Paid: false, Items: [{ Product: "" }] }]
    }, null, 2));
  });

  it("does not invent an element shape for a childless array dependency", () => {
    expect(requirementsToSampleJson([requirement("Tags", "array")])).toBe(JSON.stringify({ Tags: [] }, null, 2));
  });

  it("returns an empty object for no dependencies", () => {
    expect(requirementsToSampleJson([])).toBe("{}");
  });
});
