import type { DocumentFormat } from "./services/documentFormat";
import type { LocalFileHandle } from "./services/localFiles";
import type { TemplateReference } from "./services/workspacePersistence";

export interface Diagnostic {
  title: string;
  detail: string;
  code?: string;
  field?: string;
  part?: string;
  location?: string;
  formatter?: string;
  dataPath?: string;
  fieldCode?: string;
  removedCharacterCount?: number;
  traceId?: string;
  limitReached?: boolean;
}

export interface GeneratedDocument {
  blob: Blob;
  filename: string;
  format: DocumentFormat;
  warnings?: Diagnostic[];
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


export type Tab = "template" | "result";
export type DiagnosticsTab = "problems" | "analysis";
export interface Problem extends Diagnostic { action?: "Generate" | "Analyze" | "Generate Dataset"; severity?: "Error" | "Warning"; }
export type ExampleSelection = string | "custom";

export interface Dataset {
  id: string;
  name: string;
  kind: "example" | "generated" | "upload" | "file";
  text: string;
  baseline?: string;
  handle?: LocalFileHandle;
  lastModified?: number;
  exampleId?: string;
  restoreError?: string;
  needsPermission?: boolean;
}
export interface TemplateEntry extends TemplateReference { file?: File; restoreError?: string; needsPermission?: boolean; }
