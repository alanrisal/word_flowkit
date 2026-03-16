import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./styles/taskpane.css";

Office.onReady(() => {
  const container = document.getElementById("root");
  if (!container) {
    throw new Error("Could not find #root element");
  }
  const root = ReactDOM.createRoot(container);
  root.render(<App />);
});
