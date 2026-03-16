import React, { useRef, useEffect, useCallback } from "react";

interface Props {
  query: string;
  onChange: (value: string) => void;
}

export default function SearchBar({ query, onChange }: Props) {
  const inputRef = useRef<HTMLInputElement>(null);
  const debounceTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  // Keyboard shortcut: Ctrl/Cmd + Shift + F → focus search input
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === "F") {
        e.preventDefault();
        inputRef.current?.focus();
        inputRef.current?.select();
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, []);

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const val = e.target.value;
      if (debounceTimer.current) clearTimeout(debounceTimer.current);
      debounceTimer.current = setTimeout(() => onChange(val), 150);
    },
    [onChange]
  );

  // Keep input in sync when query is cleared externally
  useEffect(() => {
    if (query === "" && inputRef.current) {
      inputRef.current.value = "";
    }
  }, [query]);

  return (
    <div className="search-bar">
      <input
        ref={inputRef}
        type="text"
        defaultValue={query}
        onChange={handleChange}
        placeholder="Search blocks… (Ctrl+Shift+F)"
        className="search-input"
        autoFocus
        aria-label="Search debate blocks"
      />
    </div>
  );
}
