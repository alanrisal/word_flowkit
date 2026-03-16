import React, { useState, useEffect, useCallback } from "react";
import SearchBar from "./components/SearchBar";
import FileManager from "./components/FileManager";
import BlockList from "./components/BlockList";
import MultiFileToggle from "./components/MultiFileToggle";
import { FileStore } from "./lib/fileStore";
import { searchBlocks, RankedResult } from "./lib/searcher";

const fileStore = new FileStore();

export default function App() {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<RankedResult[]>([]);
  const [multiFile, setMultiFile] = useState(false);
  const [enabledFiles, setEnabledFiles] = useState<string[]>([]);
  const [allFiles, setAllFiles] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [status, setStatus] = useState<string | null>(null);

  // Load persisted files on startup
  useEffect(() => {
    fileStore.init().then(() => {
      const files = fileStore.getFileNames();
      setAllFiles(files);
      setEnabledFiles(files);
    });
  }, []);

  // Re-search when query, enabledFiles, or multiFile changes
  useEffect(() => {
    if (!query.trim()) {
      setResults([]);
      return;
    }
    const activeFiles = multiFile ? enabledFiles : enabledFiles.slice(0, 1);
    const blocks = fileStore.getAllBlocks(activeFiles);
    const ranked = searchBlocks(query, blocks);
    setResults(ranked);
  }, [query, enabledFiles, multiFile]);

  const handleFileLoaded = useCallback(async (file: File) => {
    setIsLoading(true);
    setStatus(`Parsing ${file.name}...`);
    try {
      await fileStore.addFile(file);
      const files = fileStore.getFileNames();
      setAllFiles(files);
      setEnabledFiles(files);
      setStatus(`Loaded ${file.name} — ${fileStore.getBlockCount(file.name)} blocks`);
      setTimeout(() => setStatus(null), 3000);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      setStatus(`Error: ${msg}`);
    } finally {
      setIsLoading(false);
    }
  }, []);

  const handleFileRemoved = useCallback((name: string) => {
    fileStore.removeFile(name);
    const files = fileStore.getFileNames();
    setAllFiles(files);
    setEnabledFiles(prev => prev.filter(f => f !== name));
  }, []);

  const handleToggleFile = useCallback((name: string, enabled: boolean) => {
    setEnabledFiles(prev =>
      enabled ? [...prev, name] : prev.filter(f => f !== name)
    );
  }, []);

  return (
    <div className="app">
      <header className="app-header">
        <h1>Debate Block Search</h1>
        <MultiFileToggle multiFile={multiFile} onToggle={setMultiFile} />
      </header>
      <FileManager
        files={allFiles}
        enabledFiles={enabledFiles}
        multiFile={multiFile}
        onFileLoaded={handleFileLoaded}
        onFileRemoved={handleFileRemoved}
        onToggleFile={handleToggleFile}
        isLoading={isLoading}
        fileStore={fileStore}
      />
      {status && <div className="status-bar">{status}</div>}
      <SearchBar query={query} onChange={setQuery} />
      <BlockList results={results} multiFile={multiFile} />
    </div>
  );
}
