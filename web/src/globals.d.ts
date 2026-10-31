declare const __APP_VERSION__: string;
declare module "*.css";

declare module "docx-preview" {
  export function renderAsync(
    document: Blob | ArrayBuffer | Uint8Array,
    bodyContainer: HTMLElement,
    styleContainer?: HTMLElement,
    options?: Record<string, unknown>
  ): Promise<void>;
}
