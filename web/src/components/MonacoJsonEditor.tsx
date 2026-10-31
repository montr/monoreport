import * as monaco from "monaco-editor";
import { Ref, useEffect, useImperativeHandle, useRef } from "react";

export interface JsonEditorCommands {
  format(): Promise<void>;
  copy(): Promise<void>;
  toggleWrap(): void;
}

interface Props {
  value: string;
  datasetId?: string;
  onChange(value: string): void;
  ref?: Ref<JsonEditorCommands>;
  wordWrap?: boolean;
  onWordWrapChange?(enabled: boolean): void;
  fileCommands?: { open(): void; save?(): void; saveAs?(): void };
}

export function MonacoJsonEditor({ value, datasetId, onChange, ref, wordWrap = false, onWordWrapChange, fileCommands }: Props) {
  const host = useRef<HTMLDivElement>(null);
  const editorRef = useRef<monaco.editor.IStandaloneCodeEditor | null>(null);
  const props = useRef({ onChange, onWordWrapChange, fileCommands, wordWrap });
  props.current = { onChange, onWordWrapChange, fileCommands, wordWrap };
  const applying = useRef(false);
  const previousDataset = useRef(datasetId);

  function toggleWrap() { props.current.onWordWrapChange?.(!props.current.wordWrap); }

  useImperativeHandle(ref, () => ({
    async format() { await editorRef.current?.getAction("editor.action.formatDocument")?.run(); },
    async copy() {
      const editor = editorRef.current;
      if (!editor) return;
      const selection = editor.getSelection();
      // Monaco's empty-selection copy is a line; the workspace command copies the dataset.
      const text = selection && !selection.isEmpty() ? editor.getModel()!.getValueInRange(selection) : editor.getValue();
      await navigator.clipboard.writeText(text);
    },
    toggleWrap
  }), []);

  useEffect(() => {
    if (!host.current) return;
    const model = monaco.editor.createModel(value, "json");
    model.updateOptions({ tabSize: 2, insertSpaces: true });
    model.detectIndentation(true, 2);
    const editor = monaco.editor.create(host.current, {
      model,
      theme: "vs",
      automaticLayout: true,
      minimap: { enabled: true },
      glyphMargin: false,
      scrollBeyondLastLine: false,
      fontFamily: "ui-monospace, SFMono-Regular, Consolas, monospace",
      fontSize: 12,
      ariaLabel: "Dataset",
      pasteAs: { enabled: false },
      wordWrap: props.current.wordWrap ? "on" : "off"
    });
    editorRef.current = editor;
    const actions: monaco.IDisposable[] = [];
    function fileAction(id: "open" | "save" | "saveAs", label: string, key: number) {
      if (!props.current.fileCommands?.[id]) return;
      actions.push(editor.addAction({ id: `workspace.${id}`, label, keybindings: [key], keybindingContext: "editorTextFocus", run: () => props.current.fileCommands?.[id]?.() }));
    }
    fileAction("open", "Open dataset", monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyO);
    fileAction("save", "Save dataset", monaco.KeyMod.CtrlCmd | monaco.KeyCode.KeyS);
    fileAction("saveAs", "Save dataset as", monaco.KeyMod.CtrlCmd | monaco.KeyMod.Shift | monaco.KeyCode.KeyS);
    actions.push(editor.addAction({ id: "workspace.wordWrap", label: "Word wrap", keybindingContext: "editorTextFocus", run: toggleWrap }));
    const subscription = editor.onDidChangeModelContent(() => { if (!applying.current) props.current.onChange(editor.getValue()); });
    return () => {
      editorRef.current = null;
      subscription.dispose();
      actions.forEach(action => action.dispose());
      editor.dispose();
      model.dispose();
    };
  }, []);

  useEffect(() => {
    const editor = editorRef.current;
    if (editor && (previousDataset.current !== datasetId || editor.getValue() !== value)) {
      applying.current = true;
      // setValue flushes the single model's Undo/Redo stacks, even for identical text.
      try { editor.setValue(value); } finally { applying.current = false; }
    }
    previousDataset.current = datasetId;
  }, [value, datasetId]);

  useEffect(() => { editorRef.current?.updateOptions({ wordWrap: wordWrap ? "on" : "off" }); }, [wordWrap]);

  return <div className="editor" ref={host} data-testid="json-editor" />;
}
