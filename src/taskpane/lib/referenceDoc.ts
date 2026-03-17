/* global Word */

import JSZip from "jszip";

export interface BlockIndex {
  id: string;           // unique key for React and IndexedDB
  title: string;
  headingLevel: 1 | 2 | 3;
  parentHeadings: string[];
  paragraphStart: number; // 0-indexed in body.paragraphs
  paragraphEnd: number;   // inclusive
  bodyText: string;       // plain text for fuzzy search
  sourceFile: string;
  cachedOoxml: string;    // Word-generated OOXML, extracted once at load time
}

function generateId(): string {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function getHeadingLevel(styleBuiltIn: string): 1 | 2 | 3 | null {
  if (styleBuiltIn === "Heading1") return 1;
  if (styleBuiltIn === "Heading2") return 2;
  if (styleBuiltIn === "Heading3") return 3;
  return null;
}

/**
 * Load a .docx file as base64, open it once via Word API, build the full
 * block index, and extract OOXML for every block in a single context.sync().
 *
 * After this call:
 *  - blocks[i].cachedOoxml is ready for instant paste (no re-opening needed)
 *  - styleNames contains every paragraph style used in the document
 *  - base64 is only needed for style import; it does not need to be persisted
 */
export async function loadReferenceFile(file: File): Promise<{
  blocks: BlockIndex[];
  stylesXml: string; // raw word/styles.xml from the .docx zip; "" if unreadable
}> {
  const arrayBuffer = await file.arrayBuffer();

  // Extract word/styles.xml directly from the .docx zip — no Office.js API needed.
  let stylesXml = "";
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    stylesXml = await zip.file("word/styles.xml")?.async("string") ?? "";
  } catch (e) {
    console.warn("[FlowKit] Could not extract word/styles.xml:", e);
  }

  // Convert to base64 for createDocument()
  const bytes = new Uint8Array(arrayBuffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  const base64 = btoa(binary);

  // These are populated inside Word.run and read outside
  const blocks: BlockIndex[] = [];

  await Word.run(async (context) => {
    // PASS 1: Open document and load all paragraph text + styles
    const refDoc = context.application.createDocument(base64);
    context.load(refDoc, "body");
    await context.sync();

    const paragraphs = refDoc.body.paragraphs;
    paragraphs.load("items/text,items/styleBuiltIn");
    await context.sync();

    const items = paragraphs.items;
    let h1: string | null = null;
    let h2: string | null = null;

    let current: {
      id: string;
      headingLevel: 1 | 2 | 3;
      title: string;
      parentHeadings: string[];
      bodyParts: string[];
      paragraphStart: number;
      paragraphEnd: number;
    } | null = null;

    const flushBlock = () => {
      if (!current || !current.title.trim()) return;
      blocks.push({
        id: current.id,
        title: current.title,
        headingLevel: current.headingLevel,
        parentHeadings: current.parentHeadings,
        paragraphStart: current.paragraphStart,
        paragraphEnd: current.paragraphEnd,
        bodyText: current.bodyParts.join(" ").trim(),
        sourceFile: file.name,
        cachedOoxml: "", // populated in PASS 2 below
      });
      current = null;
    };

    for (let i = 0; i < items.length; i++) {
      const para = items[i];
      const text = (para.text ?? "").trim();
      const level = getHeadingLevel(para.styleBuiltIn ?? "");

      if (level === 1) {
        flushBlock();
        h1 = text;
        h2 = null;
        current = {
          id: generateId(),
          headingLevel: 1,
          title: text,
          parentHeadings: [],
          bodyParts: [],
          paragraphStart: i,
          paragraphEnd: i,
        };
      } else if (level === 2) {
        flushBlock();
        h2 = text;
        current = {
          id: generateId(),
          headingLevel: 2,
          title: text,
          parentHeadings: h1 ? [h1] : [],
          bodyParts: [],
          paragraphStart: i,
          paragraphEnd: i,
        };
      } else if (level === 3) {
        flushBlock();
        current = {
          id: generateId(),
          headingLevel: 3,
          title: text,
          parentHeadings: [h1, h2].filter((x): x is string => x !== null),
          bodyParts: [],
          paragraphStart: i,
          paragraphEnd: i,
        };
      } else {
        if (current) {
          current.paragraphEnd = i;
          if (text) current.bodyParts.push(text);
        }
      }
    }
    flushBlock();

    // PASS 2: Queue getOoxml() for every block, then sync ONCE.
    // Word batches all the OOXML requests in parallel — one round-trip
    // for the entire file instead of one per paste.
    type Pair = { block: BlockIndex; result: OfficeExtension.ClientResult<string> };
    const pairs: Pair[] = [];

    for (const block of blocks) {
      const startPara = items[block.paragraphStart];
      const endPara = items[block.paragraphEnd];
      const range = startPara
        .getRange(Word.RangeLocation.whole)
        .expandTo(endPara.getRange(Word.RangeLocation.whole));
      pairs.push({ block, result: range.getOoxml() });
    }

    await context.sync(); // single round-trip fetches all block OOXML

    for (const { block, result } of pairs) {
      block.cachedOoxml = result.value;
    }
  });

  return { blocks, stylesXml };
}
