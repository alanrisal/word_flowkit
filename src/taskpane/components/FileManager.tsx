import React, { useRef } from "react";
import { FileStore } from "../lib/fileStore";

interface Props {
  files: string[];
  enabledFiles: string[];
  multiFile: boolean;
  onFileLoaded: (file: File) => void;
  onFileRemoved: (name: string) => void;
  onToggleFile: (name: string, enabled: boolean) => void;
  onExport: () => void;
  onImportJson: (file: File) => void;
  isLoading: boolean;
  fileStore: FileStore;
}

export default function FileManager({
  files,
  enabledFiles,
  multiFile,
  onFileLoaded,
  onFileRemoved,
  onToggleFile,
  onExport,
  onImportJson,
  isLoading,
  fileStore,
}: Props) {
  const docxInputRef = useRef<HTMLInputElement>(null);
  const jsonInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) onFileLoaded(file);
    if (docxInputRef.current) docxInputRef.current.value = "";
  };

  const handleJsonChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) onImportJson(file);
    if (jsonInputRef.current) jsonInputRef.current.value = "";
  };

  return (
    <div className="file-manager">
      <div className="file-manager-header">
        <button
          className="load-btn"
          onClick={() => docxInputRef.current?.click()}
          disabled={isLoading}
          title="Load a .docx debate file"
        >
          {isLoading ? "Loading…" : "+ Load File"}
        </button>
        <button
          className="icon-btn"
          onClick={() => jsonInputRef.current?.click()}
          disabled={isLoading}
          title="Import blocks from a FlowKit JSON backup"
        >
          ↑ JSON
        </button>
        <button
          className="icon-btn"
          onClick={onExport}
          disabled={files.length === 0 || isLoading}
          title="Export all loaded files to JSON for backup"
        >
          ↓ Export
        </button>
        <input
          ref={docxInputRef}
          type="file"
          accept=".docx"
          style={{ display: "none" }}
          onChange={handleFileChange}
          aria-hidden="true"
        />
        <input
          ref={jsonInputRef}
          type="file"
          accept=".json"
          style={{ display: "none" }}
          onChange={handleJsonChange}
          aria-hidden="true"
        />
      </div>

      {files.length > 0 && (
        <ul className="file-list" aria-label="Loaded files">
          {files.map(name => {
            const enabled = enabledFiles.includes(name);
            const count = fileStore.getBlockCount(name);
            return (
              <li
                key={name}
                className={`file-item${enabled ? "" : " disabled"}`}
              >
                {multiFile && (
                  <input
                    type="checkbox"
                    checked={enabled}
                    onChange={e => onToggleFile(name, e.target.checked)}
                    title={enabled ? "Disable this file" : "Enable this file"}
                    aria-label={`Toggle ${name}`}
                  />
                )}
                <span className="file-name" title={name}>
                  {name}
                </span>
                <span className="file-count" title={`${count} blocks parsed`}>
                  {count} blocks
                </span>
                <button
                  className="remove-btn"
                  onClick={() => onFileRemoved(name)}
                  title={`Remove ${name}`}
                  aria-label={`Remove ${name}`}
                >
                  ×
                </button>
              </li>
            );
          })}
        </ul>
      )}

      {files.length === 0 && (
        <p className="file-list-empty">No files loaded. Click "+ Load File" to add a .docx.</p>
      )}
    </div>
  );
}
