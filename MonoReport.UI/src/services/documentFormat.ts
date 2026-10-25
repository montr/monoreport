export type DocumentFormat = "docx" | "xlsx";

const templateMimes: Record<DocumentFormat, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

export const templateAccept = Object.entries(templateMimes).map(([format, mime]) => `.${format},${mime}`).join(",");

// The extension selects a preview; the server validates the document content.
export function templateFormat(filename: string): DocumentFormat | undefined {
  const extension = filename.split(".").at(-1)?.toLowerCase();
  return filename.includes(".") && (extension === "docx" || extension === "xlsx") ? extension : undefined;
}
