import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
import { initGA } from "./analytics.js";
import "./index.css";

initGA(import.meta.env.VITE_GA_ID);

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
