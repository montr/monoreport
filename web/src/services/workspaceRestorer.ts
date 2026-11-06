import type { LocalFiles } from "./localFiles";
import type { WorkspaceContext, DataReference, TemplateReference } from "./workspacePersistence";
import { documentMimes, templateFormat } from "./documentFormat";
import { filenameFromPath, type ExampleDefinition } from "./examples";
import type { Dataset, TemplateEntry } from "../types";

export class WorkspaceRestorer {
  constructor(private readonly files: LocalFiles) {}

  async restore(context: WorkspaceContext, catalog: ExampleDefinition[]): Promise<{ dataset?: Dataset; template?: TemplateEntry }> {
    const restoreData = async (entry: DataReference): Promise<Dataset> => {
      try {
        let file: File;
        if (entry.kind === "example" && entry.exampleId) {
          const example = catalog.find(item => item.id === entry.exampleId);
          if (!example) throw new Error("This example is no longer available. Reopen a JSON file to continue.");
          const response = await fetch(example.dataPath, { cache: "no-cache" });
          if (!response.ok) throw new Error("The example dataset could not be reopened.");
          file = new File([await response.text()], filenameFromPath(example.dataPath));
        } else {
          if (!entry.handle) throw new Error("Only file names and reopen references are stored. Session-only content is not recoverable; Reopen this dataset to continue.");
          file = await this.files.reopen(entry.handle);
        }
        const text = await this.files.readJson(file);
        return { ...entry, name: file.name, text, baseline: text, lastModified: entry.handle ? file.lastModified : undefined };
      } catch (error) {
        return { ...entry, text: "", baseline: "", restoreError: error instanceof Error || error instanceof DOMException ? error.message : "This dataset requires Reopen.", needsPermission: error instanceof DOMException && error.name === "NotAllowedError" };
      }
    };
    const restoreTemplate = async (entry: TemplateReference): Promise<TemplateEntry> => {
      try {
        let file: File;
        if (entry.exampleId) {
          const example = catalog.find(item => item.id === entry.exampleId);
          if (!example) throw new Error("This example is no longer available. Reopen the template to continue.");
          const response = await fetch(example.templatePath, { cache: "no-cache" });
          if (!response.ok) throw new Error("The example template could not be reopened.");
          file = new File([await response.blob()], filenameFromPath(example.templatePath), { type: documentMimes[example.format] });
        } else {
          if (!entry.handle) throw new Error("Template bytes are not stored. Reopen this template to continue.");
          file = await this.files.reopen(entry.handle);
        }
        if (!templateFormat(file.name)) throw new Error("Select a .docx or .xlsx file.");
        return { ...entry, name: file.name, file };
      } catch (error) {
        return { ...entry, restoreError: error instanceof Error || error instanceof DOMException ? error.message : "This template requires Reopen.", needsPermission: error instanceof DOMException && error.name === "NotAllowedError" };
      }
    };
    const dataset = context.datasets.find(entry => entry.id === context.datasetId);
    const template = context.templates.find(entry => entry.id === context.templateId);
    const [restoredDataset, restoredTemplate] = await Promise.all([
      dataset ? restoreData(dataset) : undefined,
      template ? restoreTemplate(template) : undefined
    ]);
    return { dataset: restoredDataset, template: restoredTemplate };
  }
}
