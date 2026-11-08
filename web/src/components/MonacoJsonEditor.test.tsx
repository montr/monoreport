import { act, render } from "@testing-library/react";
import { createRef } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";

const monaco = vi.hoisted(() => {
  const subscription = { dispose: vi.fn() };
  const model = { updateOptions: vi.fn(), detectIndentation: vi.fn(), dispose: vi.fn(), getValueInRange: vi.fn(() => "selected") };
  const action = { run: vi.fn(async () => undefined) };
  const editor = {
    getValue: vi.fn(() => '{\n  "value": 1\n}'),
    setValue: vi.fn(),
    updateOptions: vi.fn(),
    getModel: vi.fn(() => model),
    getSelection: vi.fn(() => ({ isEmpty: (): boolean => true })),
    getAction: vi.fn(() => action),
    addAction: vi.fn((_action: { id: string; keybindingContext?: string; keybindings?: number[]; run(): void }) => ({ dispose: vi.fn() })),
    onDidChangeModelContent: vi.fn((_listener: () => void) => subscription),
    dispose: vi.fn()
  };
  return { subscription, model, editor, action, createModel: vi.fn(() => model), create: vi.fn(() => editor) };
});

vi.mock("monaco-editor", () => ({
  KeyMod: { CtrlCmd: 2048, Shift: 1024 }, KeyCode: { KeyO: 45, KeyS: 49 },
  editor: { createModel: monaco.createModel, create: monaco.create }
}));

import { MonacoJsonEditor, type JsonEditorCommands } from "./MonacoJsonEditor";

const text = '{\n  "value": 1\n}';

describe("Monaco JSON editor", () => {
  beforeEach(() => { vi.clearAllMocks(); });

  it("creates an owned JSON model with detected two-space indentation and preserved editor options", () => {
    render(<MonacoJsonEditor value={'{\n    "value": 1\n}'} onChange={() => undefined} />);
    expect(monaco.createModel).toHaveBeenCalledWith('{\n    "value": 1\n}', "json");
    expect(monaco.model.updateOptions).toHaveBeenCalledWith({ tabSize: 2, insertSpaces: true });
    expect(monaco.model.detectIndentation).toHaveBeenCalledWith(true, 2);
    expect(monaco.create).toHaveBeenCalledWith(expect.any(HTMLDivElement), expect.objectContaining({ model: monaco.model, pasteAs: { enabled: false } }));
  });

  it("configures a sample preview as a read-only navigable JSON editor", () => {
    render(<MonacoJsonEditor value={text} onChange={() => undefined} readOnly ariaLabel="Sample JSON preview" minimap={false} />);
    expect(monaco.create).toHaveBeenCalledWith(expect.any(HTMLDivElement), expect.objectContaining({
      ariaLabel: "Sample JSON preview",
      folding: true,
      lineNumbers: "on",
      minimap: { enabled: false },
      readOnly: true,
      scrollbar: { vertical: "auto", horizontal: "auto" }
    }));
  });

  it("uses the local editor for value updates and disposes only its owned resources", () => {
    const onChange = vi.fn();
    const view = render(<MonacoJsonEditor value={text} onChange={onChange} />);
    monaco.editor.onDidChangeModelContent.mock.calls[0][0]();
    expect(onChange).toHaveBeenCalledWith(text);
    view.rerender(<MonacoJsonEditor value="{}" onChange={onChange} />);
    expect(monaco.editor.setValue).toHaveBeenCalledWith("{}");
    const actionDisposals = monaco.editor.addAction.mock.results.map(result => result.value.dispose);
    view.unmount();
    expect(monaco.subscription.dispose).toHaveBeenCalledOnce();
    expect(monaco.editor.dispose).toHaveBeenCalledOnce();
    expect(monaco.model.dispose).toHaveBeenCalledOnce();
    actionDisposals.forEach(dispose => expect(dispose).toHaveBeenCalledOnce());
  });

  it("flushes history on every dataset identity change even for identical text and suppresses change echoes", () => {
    const onChange = vi.fn();
    const view = render(<MonacoJsonEditor datasetId="one" value={text} onChange={onChange} />);
    expect(monaco.editor.setValue).not.toHaveBeenCalled();
    monaco.editor.setValue.mockImplementation(() => monaco.editor.onDidChangeModelContent.mock.calls[0][0]());
    view.rerender(<MonacoJsonEditor datasetId="two" value={text} onChange={onChange} />);
    view.rerender(<MonacoJsonEditor datasetId="one" value={text} onChange={onChange} />);
    expect(monaco.editor.setValue).toHaveBeenCalledTimes(2);
    expect(monaco.createModel).toHaveBeenCalledOnce();
    expect(onChange).not.toHaveBeenCalled();
  });

  it("formats natively, copies selection or whole data, and applies word wrap through public options", async () => {
    const ref = createRef<JsonEditorCommands>();
    const wrap = vi.fn();
    const clipboard = vi.spyOn(navigator.clipboard, "writeText").mockResolvedValue(undefined);
    const view = render(<MonacoJsonEditor ref={ref} value={text} onChange={() => undefined} onWordWrapChange={wrap} />);
    await ref.current!.format();
    expect(monaco.editor.getAction).toHaveBeenCalledWith("editor.action.formatDocument");
    expect(monaco.action.run).toHaveBeenCalledOnce();
    await ref.current!.copy();
    expect(clipboard).toHaveBeenLastCalledWith(text);
    monaco.editor.getSelection.mockReturnValueOnce({ isEmpty: () => false });
    await ref.current!.copy();
    expect(clipboard).toHaveBeenLastCalledWith("selected");
    act(() => ref.current!.toggleWrap());
    expect(wrap).toHaveBeenCalledWith(true);
    view.rerender(<MonacoJsonEditor ref={ref} value={text} wordWrap onChange={() => undefined} />);
    expect(monaco.editor.updateOptions).toHaveBeenLastCalledWith({ wordWrap: "on" });
  });

  it("registers file shortcuts only in editor focus and calls the latest handlers", () => {
    const oldOpen = vi.fn(), open = vi.fn(), save = vi.fn(), saveAs = vi.fn();
    const view = render(<MonacoJsonEditor value={text} onChange={() => undefined} fileCommands={{ open: oldOpen, save, saveAs }} />);
    view.rerender(<MonacoJsonEditor value={text} onChange={() => undefined} fileCommands={{ open, save, saveAs }} />);
    const actions = monaco.editor.addAction.mock.calls.map(([action]) => action);
    expect(actions.filter(action => action.keybindings).map(action => action.id)).toEqual(["workspace.open", "workspace.save", "workspace.saveAs"]);
    actions.forEach(action => expect(action.keybindingContext).toBe("editorTextFocus"));
    actions.find(action => action.id === "workspace.open")!.run();
    actions.find(action => action.id === "workspace.save")!.run();
    actions.find(action => action.id === "workspace.saveAs")!.run();
    expect(open).toHaveBeenCalledOnce();
    expect(oldOpen).not.toHaveBeenCalled();
    expect(save).toHaveBeenCalledOnce();
    expect(saveAs).toHaveBeenCalledOnce();
  });

});
