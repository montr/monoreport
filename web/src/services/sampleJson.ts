import type { TemplateRequirement } from "../types";

export interface SampleJsonSettings {
  itemsPerCollection: number;
  values: "empty" | "random";
}

const elementLimit = 100_000;

export class SampleJson {
  constructor(private readonly random: () => number) {}

  generate(requirements: readonly TemplateRequirement[], settings: SampleJsonSettings): string {
    const { itemsPerCollection: count, values } = settings;
    if (!Number.isInteger(count) || count <= 0) throw new Error("Items per collection must be a positive integer.");

    // Count before allocating: every shaped array repeats all its nested elements.
    function elements(fields: readonly TemplateRequirement[]): number {
      let total = 0;
      for (const field of fields) {
        if (field.type === "array" && field.fields?.length) total += count * (1 + elements(field.fields));
        else if (field.type === "object") total += elements(field.fields ?? []);
        if (total > elementLimit) throw new Error("Generated JSON would exceed 100,000 array elements in total. Reduce Items per collection.");
      }
      return total;
    }
    elements(requirements);

    const objectValue = (fields: readonly TemplateRequirement[] = []): Record<string, unknown> =>
      Object.fromEntries(fields.map(field => [field.name, value(field)]));
    const value = (field: TemplateRequirement): unknown => {
      switch (field.type) {
        case "boolean": return values === "random" ? this.random() >= 0.5 : false;
        case "object": return objectValue(field.fields);
        case "array": return field.fields?.length ? Array.from({ length: count }, () => objectValue(field.fields)) : [];
        default: return values === "random" ? `Sample ${Math.floor(this.random() * 1_000_000)}` : "";
      }
    };
    return JSON.stringify(objectValue(requirements), null, 2);
  }
}
