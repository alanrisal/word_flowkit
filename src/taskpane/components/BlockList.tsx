import React, { useState, useCallback, useEffect, useRef } from "react";
import { RankedResult } from "../lib/searcher";
import { pasteBlockAtCursor } from "../lib/paster";
import BlockPreview from "./BlockPreview";

interface Props {
  results: RankedResult[];
  multiFile: boolean;
}

export default function BlockList({ results, multiFile }: Props) {
  const [selectedIdx, setSelectedIdx] = useState(0);
  const [expandedId, setExpandedId] = useState<string | null>(null);
  const [pasteStatus, setPasteStatus] = useState<string | null>(null);
  const pasteTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Reset selection when results change
  useEffect(() => {
    setSelectedIdx(0);
  }, [results]);

  const handlePaste = useCallback(async (result: RankedResult) => {
    if (pasteTimerRef.current) clearTimeout(pasteTimerRef.current);
    setPasteStatus("Pasting…");
    try {
      await pasteBlockAtCursor(result.block);
      const title = result.block.title.length > 50
        ? result.block.title.slice(0, 47) + "…"
        : result.block.title;
      setPasteStatus(`Pasted: ${title}`);
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      console.error("[FlowKit] Paste failed for block:", result.block.title, e);
      setPasteStatus(`Error: ${msg}`);
    }
    pasteTimerRef.current = setTimeout(() => setPasteStatus(null), 2500);
  }, []);

  // Arrow key navigation + Enter to paste
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (results.length === 0) return;
      if (e.key === "ArrowDown") {
        e.preventDefault();
        setSelectedIdx(i => Math.min(i + 1, results.length - 1));
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        setSelectedIdx(i => Math.max(i - 1, 0));
      } else if (e.key === "Enter") {
        // Don't paste if focus is inside the search input
        const active = document.activeElement;
        if (active && active.tagName === "INPUT" && (active as HTMLInputElement).type === "text") {
          return;
        }
        e.preventDefault();
        handlePaste(results[selectedIdx]);
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [results, selectedIdx, handlePaste]);

  if (results.length === 0) {
    return <div className="block-list-empty">Type to search blocks</div>;
  }

  return (
    <div className="block-list" role="list" aria-label="Search results">
      {pasteStatus && <div className="paste-status">{pasteStatus}</div>}
      {results.map((result, idx) => {
        const { block, score } = result;
        const isSelected = idx === selectedIdx;
        const isExpanded = expandedId === block.id;
        const breadcrumb = block.parentHeadings.join(" › ");
        const scorePercent = Math.round(score * 100);

        return (
          <div
            key={block.id}
            className={`block-item${isSelected ? " selected" : ""}`}
            role="listitem"
            onClick={() => handlePaste(result)}
            onMouseEnter={() => setSelectedIdx(idx)}
            title="Click to paste into Word"
            aria-selected={isSelected}
          >
            <div className="block-item-header">
              <span className="block-title">{block.title}</span>
              <button
                className="expand-btn"
                onClick={e => {
                  e.stopPropagation();
                  setExpandedId(isExpanded ? null : block.id);
                }}
                title={isExpanded ? "Collapse preview" : "Show preview"}
                aria-expanded={isExpanded}
              >
                {isExpanded ? "▲" : "▼"}
              </button>
            </div>
            {breadcrumb && (
              <div className="block-breadcrumb" title={breadcrumb}>
                {breadcrumb}
              </div>
            )}
            {multiFile && (
              <div className="block-source" title={block.sourceFile}>
                {block.sourceFile}
              </div>
            )}
            <div className="score-bar" title={`Match score: ${scorePercent}%`}>
              <div
                className="score-fill"
                style={{ width: `${scorePercent}%` }}
              />
            </div>
            {isExpanded && <BlockPreview block={block} />}
          </div>
        );
      })}
    </div>
  );
}
