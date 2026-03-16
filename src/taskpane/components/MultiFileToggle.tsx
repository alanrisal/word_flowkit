import React from "react";

interface Props {
  multiFile: boolean;
  onToggle: (value: boolean) => void;
}

export default function MultiFileToggle({ multiFile, onToggle }: Props) {
  return (
    <div className="multi-file-toggle">
      <label className="toggle-label" title={multiFile ? "Search all files" : "Search first file only"}>
        <input
          type="checkbox"
          checked={multiFile}
          onChange={e => onToggle(e.target.checked)}
          aria-label="Enable multi-file search"
        />
        <span className="toggle-text">Multi-file</span>
      </label>
    </div>
  );
}
