import { act, render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { ConfigProvider, Form, type FormInstance } from "antd";
import { useRef, useState } from "react";
import { describe, expect, it, vi } from "vitest";
import type { Workspace } from "../hooks/useWorkspace";
import type { SampleJsonSettings } from "../services/sampleJsonRequest";
import { GenerateDatasetDialog } from "./GenerateDatasetDialog";

describe("Generate Dataset dialog", () => {
  it("does not submit validation completed after cancellation and permits a new session", async () => {
    const user = userEvent.setup();
    const submit = vi.fn(async (_settings: SampleJsonSettings) => undefined);
    const resolvers: Array<(settings: SampleJsonSettings) => void> = [];
    let form: FormInstance<SampleJsonSettings>;

    function Harness() {
      const [open, setOpenState] = useState(true);
      const session = useRef(0);
      const [currentForm] = Form.useForm<SampleJsonSettings>();
      form = currentForm;
      const setOpen = (value: boolean) => {
        session.current += 1;
        setOpenState(value);
      };
      const model = {
        jsonSettingsOpen: open,
        pending: undefined,
        submitJsonSettings: submit,
        invalidateOperation: vi.fn(),
        setJsonSettingsOpen: setOpen,
        captureSampleValidation: () => {
          const expectedSession = session.current;
          return () => expectedSession === session.current;
        },
        token: {marginSM: 8, marginXS: 4},
        jsonSettings: currentForm,
        sampleCandidate: undefined,
        acceptSample: vi.fn(),
        copySampleJson: vi.fn(),
        modified: false,
        settingsChanged: vi.fn()
      } as unknown as Workspace["dialogs"]["generateDataset"];
      return (
        <ConfigProvider theme={{token: {motion: false}}}>
          <button onClick={() => setOpen(true)}>Reopen</button>
          <GenerateDatasetDialog model={model} />
        </ConfigProvider>
      );
    }

    render(<Harness />);
    vi.spyOn(form!, "validateFields").mockImplementation(() => new Promise(resolve => resolvers.push(resolve)));
    await user.click(screen.getByRole("button", {name: "Generate sample"}));
    await user.click(screen.getByRole("button", {name: "Cancel"}));
    await waitFor(() => expect(screen.queryByRole("dialog")).not.toBeInTheDocument());
    await act(async () => resolvers[0]({itemsPerCollection: 10, values: "filled"}));
    expect(submit).not.toHaveBeenCalled();

    await user.click(screen.getByRole("button", {name: "Reopen"}));
    await user.click(await screen.findByRole("button", {name: "Generate sample"}));
    await act(async () => resolvers[1]({itemsPerCollection: 2, values: "filled"}));
    await waitFor(() => expect(submit).toHaveBeenCalledOnce());
    expect(submit).toHaveBeenCalledWith({itemsPerCollection: 2, values: "filled"});
  });
});
