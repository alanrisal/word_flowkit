# Word Debate Block Search Add-In — Build Blueprint

## Project Overview

A Microsoft Word Task Pane Add-in for policy debate that enables semantic fuzzy search across one or multiple `.docx` debate files. Users type a header/argument into a sidebar, get ranked block matches, and paste them directly into their working document with one click — preserving all original Word formatting and styles.

---

## Architecture Decision

**Use: Office.js Task Pane Add-in (Web-based)**

This is the correct approach over VBA macros for the following reasons:

| Factor | Office.js Add-in | VBA Macro |
|---|---|---|
| Cross-platform (Mac + Windows) | ✅ Yes | ❌ Windows only |
| Shareable without IT admin | ✅ Sideload manifest | ⚠️ Must share .xlam |
| Modern UI (React/HTML) | ✅ Full control | ❌ Limited |
| Semantic search / NLP | ✅ JS ecosystem | ❌ Very hard |
| Multi-file search | ✅ Yes | ⚠️ Clunky |
| Paste with formatting preserved | ✅ Via OOXML | ✅ Native |

**Stack:**
- **Frontend:** React + TypeScript (in the task pane)
- **Office API:** `office-js` (Microsoft's official add-in SDK)
- **Fuzzy Search:** `fuse.js` for fast local fuzzy matching
- **Semantic Search (optional/upgrade):** Anthropic Claude API or `transformers.js` (runs in-browser)
- **File Parsing:** `mammoth.js` for reading `.docx` block content
- **Paste Engine:** Office Open XML (OOXML) inserted via `context.document.body.insertOoxml()`

---

## Repository Structure

```
debate-word-addin/
├── manifest.xml              # Office Add-in manifest (share this with your friend)
├── package.json
├── webpack.config.js
├── src/
│   ├── taskpane/
│   │   ├── index.html        # Task pane shell
│   │   ├── index.tsx         # React entry point
│   │   ├── App.tsx           # Root component
│   │   ├── components/
│   │   │   ├── SearchBar.tsx       # Input box with debounce
│   │   │   ├── FileManager.tsx     # Load/manage debate files
│   │   │   ├── BlockList.tsx       # Ranked results list
│   │   │   ├── BlockPreview.tsx    # Hover/expand preview
│   │   │   └── MultiFileToggle.tsx # Switch single ↔ multi-file mode
│   │   ├── lib/
│   │   │   ├── parser.ts       # Parse .docx → blocks by heading level
│   │   │   ├── searcher.ts     # Fuzzy + semantic search logic
│   │   │   ├── paster.ts       # OOXML paste engine
│   │   │   ├── fileStore.ts    # IndexedDB persistence for loaded files
│   │   │   └── ooxmlBuilder.ts # Rebuild OOXML from block data
│   │   └── styles/
│   │       └── taskpane.css
└── assets/
    └── icon-32.png
```

---

## Core Feature Breakdown

### 1. File Loading & Parsing

**How it works:**

The user clicks "Load File" and selects a `.docx` debate file via a standard `<input type="file">` picker. The file is parsed using `mammoth.js` to extract its raw OOXML XML, then a custom parser splits it into **blocks** — defined as everything under a heading element until the next heading of the same or higher level.

**Block definition for debate files:**

```
Heading 1 = Main argument / file title
Heading 2 = Sub-argument / block label  ← Primary search unit
Heading 3 = Card tag
Body text  = Card text (warrant, cite, etc.)
```

Each block is stored as:

```typescript
interface DebateBlock {
  id: string;
  sourceFile: string;
  headingLevel: 1 | 2 | 3;
  title: string;           // The heading text — this is what gets searched
  bodyText: string;        // Stripped plain text for search ranking
  rawOoxml: string;        // Full OOXML preserved for paste — never stripped
  parentHeadings: string[]; // Breadcrumb: ["Neg", "Politics DA", "Link Wall"]
}
```

**Parser logic (`parser.ts`):**

```typescript
import mammoth from "mammoth";

export async function parseDebateFile(file: File): Promise<DebateBlock[]> {
  // 1. Get raw OOXML from the .docx
  const arrayBuffer = await file.arrayBuffer();
  
  // 2. Use mammoth to extract paragraph-level data with style info
  const result = await mammoth.convertToHtml({ arrayBuffer }, {
    styleMap: [
      "p[style-name='Heading 1'] => h1",
      "p[style-name='Heading 2'] => h2",
      "p[style-name='Heading 3'] => h3",
    ]
  });

  // 3. Walk the paragraph tree, group by headings into blocks
  // Each time we hit a H1/H2 heading, start a new block
  // Accumulate all body paragraphs until the next heading
  
  // 4. For each block, also extract its raw OOXML slice 
  // (from the source .docx XML directly — see ooxmlBuilder.ts)
  
  return blocks;
}
```

**CRITICAL — Preserving OOXML for paste:**

`mammoth.js` is used ONLY for parsing structure. The actual OOXML for each block must be extracted directly from the `.docx` XML (which is a zip file). Use `JSZip` to unzip the `.docx`, read `word/document.xml`, and slice out the `<w:p>` nodes that belong to each block. This is what gets pasted.

```typescript
import JSZip from "jszip";

export async function extractRawOoxml(file: File): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const documentXml = await zip.file("word/document.xml")!.async("string");
  // Parse XML, walk <w:body><w:p> elements
  // Map heading text → slice of raw <w:p>...</w:p> nodes
}
```

---

### 2. Search Engine

**Two-tier search:**

#### Tier 1: Fuzzy Search (always on, instant)
Uses `fuse.js` with the following config:

```typescript
import Fuse from "fuse.js";

const fuse = new Fuse(blocks, {
  keys: [
    { name: "title", weight: 0.7 },       // heading text — highest weight
    { name: "bodyText", weight: 0.2 },     // card text
    { name: "parentHeadings", weight: 0.1 } // breadcrumb context
  ],
  threshold: 0.4,       // 0 = exact, 1 = match anything
  includeScore: true,
  useExtendedSearch: true,
  ignoreLocation: true,  // important: debate tags can be anywhere in the string
  minMatchCharLength: 2
});

export function searchBlocks(query: string, blocks: DebateBlock[]): RankedResult[] {
  return fuse.search(query).map(r => ({
    block: r.item,
    score: 1 - (r.score ?? 0),  // invert: higher = better
    matchedIn: inferMatchLocation(r)
  }));
}
```

#### Tier 2: Semantic Search (optional upgrade)
For users who want "find me blocks that answer topicality" (conceptual, not keyword):

Option A — **Claude API** (requires internet): Send the query + all block titles to Claude and ask it to rank by relevance. Best quality, requires API key.

Option B — **transformers.js** (fully offline): Embed blocks and query using `@xenova/transformers` with `all-MiniLM-L6-v2`. Run entirely in the browser. Slower on first load but private.

```typescript
// Semantic reranking using cosine similarity
import { pipeline } from "@xenova/transformers";

const embedder = await pipeline("feature-extraction", "Xenova/all-MiniLM-L6-v2");

export async function semanticRank(query: string, candidates: DebateBlock[]) {
  const queryEmbedding = await embedder(query, { pooling: "mean" });
  const blockEmbeddings = await Promise.all(
    candidates.map(b => embedder(b.title + " " + b.bodyText.slice(0, 200), { pooling: "mean" }))
  );
  // cosine similarity ranking
}
```

---

### 3. Search UI (Task Pane)

**`SearchBar.tsx`:**
- Text input that debounces 150ms before triggering search
- Keyboard shortcut `Cmd/Ctrl + Shift + F` to focus
- Shows "searching..." spinner during semantic rerank

**`BlockList.tsx`:**
- Renders top 10–20 ranked results
- Each result shows:
  - Block title (heading text) — bolded
  - Breadcrumb path: `Neg > Politics DA > Link Wall`
  - Source filename (for multi-file mode)
  - Score bar (subtle visual indicator)
- Click or press `Enter` on selected item → paste into Word
- Arrow keys to navigate the list
- Hover/click expand → shows first 3 lines of block body text

**`FileManager.tsx`:**
- "Load File" button (single file mode)
- "Manage Files" panel (multi-file mode):
  - Lists all loaded files with toggle on/off per file
  - Files stored in **IndexedDB** so they persist across sessions
  - "Remove" button per file
  - Shows block count per file

---

### 4. Paste Engine

This is the most critical part. Pasting must preserve ALL Word formatting.

**Method: Insert OOXML via Office.js**

```typescript
import * as Office from "@microsoft/office-js";

export async function pasteBlockAtCursor(block: DebateBlock): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    
    // Build valid OOXML wrapper around the block's raw paragraphs
    const ooxml = buildOoxmlDocument(block.rawOoxml, block.sourceFile);
    
    // Insert at cursor position, replacing current selection if any
    selection.insertOoxml(ooxml, Word.InsertLocation.replace);
    
    await context.sync();
  });
}

function buildOoxmlDocument(rawParagraphs: string, sourceFile: string): string {
  // OOXML must be wrapped in a full document fragment
  return `
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:wordprocessingML xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <w:body>
        ${rawParagraphs}
      </w:body>
    </w:wordprocessingML>
  `;
}
```

**Style preservation:** Since you're inserting raw OOXML from the source file, all `<w:rPr>` (run properties: bold, font, size, color) and `<w:pPr>` (paragraph properties: heading level, indent, spacing) are preserved verbatim. This is equivalent to Verbatim's tilde paste — but driven by search.

**Note on style conflicts:** If the destination document has different style definitions for e.g. "Heading 2" than the source, Word will apply the destination's Heading 2 style definition. To avoid this, consider normalizing to inline formatting in the OOXML builder.

---

### 5. Multi-File Search

When multi-file mode is active, the search index is a **merged Fuse index** across all loaded files:

```typescript
// fileStore.ts
export class FileStore {
  private db: IDBDatabase;
  private loadedFiles: Map<string, DebateBlock[]> = new Map();
  
  async addFile(file: File): Promise<void> {
    const blocks = await parseDebateFile(file);
    this.loadedFiles.set(file.name, blocks);
    await this.persistToIndexedDB(file.name, blocks);
  }
  
  getAllBlocks(enabledFiles?: string[]): DebateBlock[] {
    const files = enabledFiles ?? [...this.loadedFiles.keys()];
    return files.flatMap(name => this.loadedFiles.get(name) ?? []);
  }
  
  rebuildSearchIndex(): Fuse<DebateBlock> {
    return new Fuse(this.getAllBlocks(), fuseConfig);
  }
}
```

Search results in multi-file mode show the source filename so debaters know where each block is from.

---

## Sharing with Your Friend

### Option A: Sideloading (no App Store needed)

1. Build the project: `npm run build`
2. Host the built files anywhere (even locally via `npx serve dist`)
3. Share the `manifest.xml` file
4. Friend opens Word → Insert → Add-ins → Upload My Add-in → select `manifest.xml`
5. The add-in appears in the task pane

### Option B: GitHub Pages (easiest for sharing)

1. Push the `dist/` build to GitHub Pages
2. Update `manifest.xml` to point to `https://yourusername.github.io/debate-addin/`
3. Share only the `manifest.xml` — friend loads it via sideloading above
4. No local server needed

**`manifest.xml` key fields:**

```xml
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
           xsi:type="TaskPaneApp">
  <Id>YOUR-GUID-HERE</Id>
  <DisplayName DefaultValue="Debate Block Search"/>
  <DefaultLocale>en-US</DefaultLocale>
  <Hosts>
    <Host Name="Document"/>  <!-- Word -->
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://YOUR-HOSTING-URL/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

---

## Getting Started (Build Steps)

```bash
# 1. Scaffold with Yeoman (Microsoft's official generator)
npm install -g yo generator-office
yo office

# Choose: Task Pane Add-in, Word, TypeScript, React

# 2. Install additional dependencies
npm install fuse.js mammoth jszip @xenova/transformers idb

# 3. Dev server (auto-sideloads into Word on Mac/Windows)
npm run dev-server

# 4. Production build
npm run build
```

---

## Future Enhancements (Roadmap)

### Near-term
- **Tilde-style paste shortcut** — paste without leaving keyboard
- **Block preview on hover** — see first 3 lines before pasting
- **Recent pastes history** — quick re-paste last 5 blocks
- **Tag filtering** — filter results to only show cards tagged with certain authors/dates

### Medium-term
- **Cross-file answer finder** — type an argument ("Hegemony Good"), get blocks that *answer* it from all loaded files, ranked by counter-argument relevance (Claude API)
- **Block deduplication** — detect near-duplicate blocks across files
- **Export to new doc** — collect N searched blocks into a new document in one click

### Long-term
- **Cloud sync** — store your debate file library in the cloud, search from any machine
- **Tournament prep mode** — pre-load opponent's files, search "what do they say about [X]"
- **Auto-cite formatting** — detect card citation metadata and auto-format to your team's cite style

---

## Key Technical Gotchas

1. **OOXML must be a full document fragment** — `insertOoxml()` requires the wrapper `<w:wordprocessingML>` element, not bare `<w:p>` nodes.

2. **mammoth.js strips formatting** — use it only for structure parsing. Always go back to the raw ZIP XML for the OOXML you'll paste.

3. **IndexedDB size limits** — each `.docx` file can be 1–50MB. Modern browsers allow 50–500MB of IndexedDB storage. For very large file libraries, store only block metadata in IndexedDB and re-parse raw OOXML on demand.

4. **Office.js async** — all Word API calls must be inside `Word.run(async context => { ... await context.sync(); })`.

5. **Fuse.js threshold tuning** — debate block titles are often very terse ("UQ – Russia Heg High Now"). Set `ignoreLocation: true` and `threshold: 0.4` for best results. Expose threshold as a user setting.

6. **Cross-origin manifest** — if hosting on GitHub Pages, ensure the manifest `SourceLocation` uses `https://` not `http://`. Word requires HTTPS for production add-ins.