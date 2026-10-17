export interface Diagnostic {
  title: string;
  detail: string;
  code?: string;
  field?: string;
  part?: string;
  location?: string;
  formatter?: string;
  dataPath?: string;
}

export interface GeneratedDocument {
  blob: Blob;
  filename: string;
}

export interface TemplateRequirement {
  name: string;
  type: "object" | "array" | "value" | "boolean";
  fields?: TemplateRequirement[];
}

export interface TemplateOccurrence {
  kind: "field" | "for" | "if";
  expression: string;
  path: string;
  part: string;
  location: string;
}

export interface TemplateAnalysis {
  valid: boolean;
  errors: Diagnostic[];
  warnings: Diagnostic[];
  requirements: TemplateRequirement[];
  occurrences: TemplateOccurrence[];
}
