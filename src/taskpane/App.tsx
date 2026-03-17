import React, { useState, useEffect, useCallback } from "react";
import SearchBar from "./components/SearchBar";
import FileManager from "./components/FileManager";
import BlockList from "./components/BlockList";
import MultiFileToggle from "./components/MultiFileToggle";
import { FileStore } from "./lib/fileStore";
import { loadReferenceFile } from "./lib/referenceDoc";
import { pasteBlock } from "./lib/paster";
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

  const handlePaste = useCallback(async (result: RankedResult) => {
    const stylesXml = fileStore.getStylesXml(result.block.sourceFile);
    await pasteBlock(result.block, stylesXml || undefined);
  }, []);

  const handleFileLoaded = useCallback(async (file: File) => {
    setIsLoading(true);
    setStatus(`Parsing ${file.name}...`);
    try {
      // Parse the file — extracts block OOXML and word/styles.xml upfront.
      // createDocument() is called exactly once per file load here.
      const { blocks, stylesXml } = await loadReferenceFile(file);

      // Store blocks and styles XML (both persist to IndexedDB across sessions).
      await fileStore.addFile(file, blocks, stylesXml);

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

  const handleExport = useCallback(() => {
    const json = fileStore.exportToJson();
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "flowkit-blocks.json";
    a.click();
    URL.revokeObjectURL(url);
  }, []);

  const handleImportJson = useCallback(async (file: File) => {
    setIsLoading(true);
    setStatus(`Importing ${file.name}...`);
    try {
      const text = await file.text();
      const imported = await fileStore.importFromJson(text);
      const files = fileStore.getFileNames();
      setAllFiles(files);
      setEnabledFiles(files);
      setStatus(`Imported ${imported.length} file(s) from JSON`);
      setTimeout(() => setStatus(null), 3000);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      setStatus(`Import error: ${msg}`);
    } finally {
      setIsLoading(false);
    }
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
        onExport={handleExport}
        onImportJson={handleImportJson}
        isLoading={isLoading}
        fileStore={fileStore}
      />
      {isLoading && <div className="loading-bar" aria-label="Loading…" role="progressbar" />}
      {status && <div className="status-bar">{status}</div>}
      <SearchBar query={query} onChange={setQuery} />
      <BlockList results={results} multiFile={multiFile} onPaste={handlePaste} />
    </div>
  );
}
