import React, { useState, useEffect, useCallback } from "react";
import SearchBar from "./components/SearchBar";
import FileManager from "./components/FileManager";
import BlockList from "./components/BlockList";
import MultiFileToggle from "./components/MultiFileToggle";
import { FileStore } from "./lib/fileStore";
import { loadReferenceFile } from "./lib/referenceDoc";
import { importStylesIntoActiveDocument } from "./lib/styleImporter";
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
      // Step 1: Parse the file — extracts block OOXML and style definitions upfront.
      // createDocument() is called exactly once per file load here.
      const { base64, blocks, styleNames, stylesJson } = await loadReferenceFile(file);

      // Step 2: Import any missing styles into the active document once.
      // Fast path when stylesJson is available (no createDocument needed).
      // Falls back to base64 re-open + addStyle() on older Word versions.
      await importStylesIntoActiveDocument(
        stylesJson ? { stylesJson } : { base64 },
        styleNames
      );

      // Step 3: Store blocks and style JSON (cachedOoxml and stylesJson
      // both persist to IndexedDB so they survive session restarts).
      await fileStore.addFile(file, blocks, stylesJson);

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
      {isLoading && <div className="loading-bar" aria-label="Loading…" role="progressbar" />}
      {status && <div className="status-bar">{status}</div>}
      <SearchBar query={query} onChange={setQuery} />
      <BlockList results={results} multiFile={multiFile} />
    </div>
  );
}
