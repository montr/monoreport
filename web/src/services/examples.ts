import { documentMimes } from "./documentFormat";
import type { DocumentFormat } from "./documentFormat";

export interface ExampleDefinition {
  id: string;
  title: string;
  format: DocumentFormat;
  templatePath: string;
  dataPath: string;
}

export interface ExampleManifest {
  defaultExample: string;
  examples: ExampleDefinition[];
}

export function validateManifest(value: unknown): ExampleManifest {
  if (!value || typeof value !== "object") throw new Error("The examples registry is invalid.");
  const candidate = value as Partial<ExampleManifest>;
  if (typeof candidate.defaultExample !== "string" || !Array.isArray(candidate.examples) || candidate.examples.length === 0) throw new Error("The examples registry is invalid.");
  const ids = new Set<string>();
  for (const example of candidate.examples) {
    if (!example || typeof example.id !== "string" || !example.id || ids.has(example.id) || typeof example.title !== "string" || !example.title || (example.format !== "docx" && example.format !== "xlsx") || typeof example.templatePath !== "string" || !example.templatePath || typeof example.dataPath !== "string" || !example.dataPath) throw new Error("The examples registry is invalid.");
    ids.add(example.id);
  }
  if (!ids.has(candidate.defaultExample)) throw new Error("The default example is missing from the registry.");
  return candidate as ExampleManifest;
}

export function filenameFromPath(path: string) {
  const segment = path.split("/").filter(Boolean).at(-1) ?? "template.docx";
  try { return decodeURIComponent(segment); } catch { return segment; }
}


export class ExampleCatalog {
  constructor(private readonly fetcher: typeof fetch) {}

  async loadManifest(): Promise<ExampleManifest> {
    const response = await this.fetcher("/examples.json", { cache: "no-cache" });
    if (!response.ok) throw new Error("The examples registry could not be loaded.");
    return validateManifest(await response.json());
  }

  async loadAssets(example: ExampleDefinition): Promise<{ file: File; data: string }> {
    const [templateResponse, dataResponse] = await Promise.all([
      this.fetcher(example.templatePath, { cache: "no-cache" }),
      this.fetcher(example.dataPath, { cache: "no-cache" })
    ]);
    if (!templateResponse.ok || !dataResponse.ok) throw new Error("The selected example assets could not be loaded.");
    const [blob, data] = await Promise.all([templateResponse.blob(), dataResponse.text()]);
    JSON.parse(data);
    return { file: new File([blob], filenameFromPath(example.templatePath), { type: documentMimes[example.format] }), data };
  }
}
