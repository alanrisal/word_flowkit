import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./styles/taskpane.css";

function mountApp() {
  const container = document.getElementById("root");
  if (!container) return;
  ReactDOM.createRoot(container).render(<App />);
}

if (typeof Office !== "undefined") {
  Office.onReady(mountApp);
} else {
  mountApp();
}
