import { useEffect, type Dispatch, type RefObject, type SetStateAction } from "react";

interface Props {
  generateAction: RefObject<() => void>;
  jsonSettingsOpen: boolean;
  shortcutsOpen: boolean;
  problemsOpen: boolean;
  saveData(saveAs?: boolean): Promise<void>;
  changeProblemsOpen(open: boolean): void;
  setShortcutsOpen: Dispatch<SetStateAction<boolean>>;
}

export function useWorkspaceShortcuts({ generateAction, jsonSettingsOpen, shortcutsOpen, problemsOpen, saveData, changeProblemsOpen, setShortcutsOpen }: Props) {
  useEffect(() => {
    function generateShortcut(event: KeyboardEvent) {
      if (event.defaultPrevented || event.isComposing || event.repeat || event.altKey || event.shiftKey || jsonSettingsOpen || shortcutsOpen) return;
      if (event.key !== "Enter" || event.metaKey === event.ctrlKey) return;
      const target = event.target instanceof Element ? event.target : undefined;
      if (target?.closest('[role="dialog"], [role="alertdialog"]')) return;
      // Monaco binds Ctrl/Cmd+Enter to insert a line below; generation takes precedence for this workspace shortcut.
      event.preventDefault();
      event.stopPropagation();
      generateAction.current();
    }
    function keydown(event: KeyboardEvent) {
      if (event.defaultPrevented || event.isComposing || event.repeat || event.altKey || jsonSettingsOpen || shortcutsOpen) return;
      const target = event.target instanceof Element ? event.target : undefined;
      if (target?.closest('[role="dialog"], [role="alertdialog"]')) return;
      const workspaceModifier = event.metaKey !== event.ctrlKey;
      const saveShortcut = event.key.toLowerCase() === "s" && workspaceModifier;
      if (saveShortcut) {
        const editable = target?.closest('input, textarea, select, [contenteditable]:not([contenteditable="false"]), [role="textbox"], [role="combobox"]');
        if (editable && !target?.closest(".monaco-editor")) return;
        event.preventDefault();
        if (event.shiftKey) void saveData(true);
        else void saveData();
        return;
      }
      if (event.ctrlKey || event.metaKey) return;
      // Native controls and Monaco handle their bindings before this bubble listener.
      if (event.key === "Escape" && event.shiftKey && problemsOpen) {
        event.preventDefault();
        changeProblemsOpen(false);
      } else if (event.key === "?" && event.shiftKey && !target?.closest('input, textarea, select, [contenteditable]:not([contenteditable="false"]), [role="textbox"], [role="combobox"], .monaco-editor')) {
        event.preventDefault();
        setShortcutsOpen(true);
      }
    }
    document.addEventListener("keydown", generateShortcut, true);
    document.addEventListener("keydown", keydown);
    return () => {
      document.removeEventListener("keydown", generateShortcut, true);
      document.removeEventListener("keydown", keydown);
    };
  }, [problemsOpen, jsonSettingsOpen, shortcutsOpen]);

}
