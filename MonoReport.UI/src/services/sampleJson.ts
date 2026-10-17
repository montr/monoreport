import type { TemplateRequirement } from "../types";

function objectValue(fields: readonly TemplateRequirement[] | undefined): Record<string, unknown> {
  return Object.fromEntries((fields ?? []).map(field => [field.name, value(field)]));
}

function value(requirement: TemplateRequirement): unknown {
  switch (requirement.type) {
    case "boolean": return false;
    case "object": return objectValue(requirement.fields);
    case "array": return requirement.fields?.length ? [objectValue(requirement.fields)] : [];
    default: return "";
  }
}

export function requirementsToSampleJson(requirements: readonly TemplateRequirement[]): string {
  return JSON.stringify(Object.fromEntries(requirements.map(requirement => [requirement.name, value(requirement)])), null, 2);
}
