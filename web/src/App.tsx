import { Alert, ConfigProvider } from "antd";
import { WorkspaceLayout } from "./components/WorkspaceLayout";
import { useWorkspace } from "./hooks/useWorkspace";
import { KeyboardShortcutsDialog } from "./components/KeyboardShortcutsDialog";
import { GenerateDatasetDialog } from "./components/GenerateDatasetDialog";
import { WorkspaceHeader } from "./components/WorkspaceHeader";
import { DatasetPanel } from "./components/DatasetPanel";
import { DocumentPanel } from "./components/DocumentPanel";
import { DiagnosticsPanel } from "./components/DiagnosticsPanel";
import { WorkspaceFooter } from "./components/WorkspaceFooter";
import "./styles.css";

export default function App() {
  const workspace = useWorkspace();
  const { dialogs, layout } = workspace;
  return (
    <ConfigProvider>
      {dialogs.contexts.modalContext}
      {dialogs.contexts.feedbackContext}
      <KeyboardShortcutsDialog model={dialogs.shortcuts} />
      <GenerateDatasetDialog model={dialogs.generateDataset} />
      <div className="app">
        <WorkspaceHeader model={workspace.header} />
        {workspace.storageWarning && (
          <Alert type="warning" title={workspace.storageWarning} showIcon />
        )}
        <main className="layout">
          <WorkspaceLayout
            key={layout.workspaceKey}
            problemsOpen={layout.problemsOpen}
            onProblemsOpenChange={layout.onProblemsOpenChange}
            data={<DatasetPanel model={workspace.datasetPanel} />}
            document={<DocumentPanel model={workspace.documentPanel} />}
            problems={<DiagnosticsPanel model={workspace.diagnosticsPanel} />}
          />
        </main>
        <WorkspaceFooter model={workspace.footer} />
      </div>
    </ConfigProvider>
  );
}
