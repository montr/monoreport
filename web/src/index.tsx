import React from "react";
import { createRoot } from "react-dom/client";
import "antd/dist/reset.css";
import App from "./App";

const root = document.getElementById("root");
if (!root) throw new Error("Application root was not found.");
createRoot(root).render(<React.StrictMode><App /></React.StrictMode>);
