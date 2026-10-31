export type DocumentFormat = "docx" | "xlsx";

export const documentMimes: Record<DocumentFormat, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

export function documentFormatFromMime(value: string | null | undefined): DocumentFormat | undefined {
  const mime = value?.split(";", 1)[0]?.trim().toLowerCase();
  return Object.entries(documentMimes).find(([, supported]) => supported === mime)?.[0] as DocumentFormat | undefined;
}

export function documentFallbackFilename(format: DocumentFormat): string {
  return `document.${format}`;
}

export const templateAccept = Object.entries(documentMimes).map(([format, mime]) => `.${format},${mime}`).join(",");

// The extension selects a preview; the server validates the document content.
export function templateFormat(filename: string): DocumentFormat | undefined {
  const extension = filename.split(".").at(-1)?.toLowerCase();
  return filename.includes(".") && (extension === "docx" || extension === "xlsx") ? extension : undefined;
}
