import JSZip from "jszip";

export interface DebateBlock {
  id: string;
  sourceFile: string;
  headingLevel: 1 | 2 | 3;
  title: string;
  bodyText: string;
  parentHeadings: string[];
  paragraphStart: number; // 0-indexed position among <w:p> children of <w:body>
  paragraphEnd: number;   // inclusive
}

function generateId(): string {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function getHeadingLevelFromElement(el: Element): 1 | 2 | 3 | null {
  const pStyleEls = el.getElementsByTagNameNS(W_NS, "pStyle");
  if (!pStyleEls.length) return null;
  const pStyle = pStyleEls[0];
  const val = (
    pStyle.getAttributeNS(W_NS, "val") ??
    pStyle.getAttribute("w:val") ??
    ""
  ).toLowerCase().replace(/\s/g, "");
  if (val === "heading1") return 1;
  if (val === "heading2") return 2;
  if (val === "heading3") return 3;
  return null;
}

function extractTextFromElement(el: Element): string {
  const tNodes = el.getElementsByTagNameNS(W_NS, "t");
  return Array.from(tNodes)
    .map(t => t.textContent ?? "")
    .join("");
}

/**
 * Extract only the direct <w:p> children of <w:body>, 0-indexed.
 * This mirrors how Word.js body.paragraphs.items is indexed, so
 * paragraphStart/paragraphEnd values can be used interchangeably
 * in paster.ts when re-opening the same document via createDocument().
 */
async function extractParagraphElements(file: File): Promise<Element[]> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) {
    throw new Error(`Could not find word/document.xml inside ${file.name}`);
  }
  const docXml = await docFile.async("string");

  const domParser = new DOMParser();
  const xmlDoc = domParser.parseFromString(docXml, "application/xml");

  const parseError = xmlDoc.querySelector("parsererror");
  if (parseError) {
    throw new Error(`Malformed document.xml in ${file.name}: ${parseError.textContent}`);
  }

  const bodyEl = xmlDoc.getElementsByTagNameNS(W_NS, "body")[0];
  if (!bodyEl) {
    throw new Error(`Could not find <w:body> in ${file.name}`);
  }

  const results: Element[] = [];
  for (const child of Array.from(bodyEl.childNodes)) {
    if (child.nodeType === Node.ELEMENT_NODE && (child as Element).localName === "p") {
      results.push(child as Element);
    }
  }
  return results;
}

export async function parseDebateFile(file: File): Promise<DebateBlock[]> {
  const elements = await extractParagraphElements(file);
  const blocks: DebateBlock[] = [];

  let h1: string | null = null;
  let h2: string | null = null;

  let currentBlock: {
    headingLevel: 1 | 2 | 3;
    title: string;
    parentHeadings: string[];
    bodyParts: string[];
    paragraphStart: number;
    paragraphEnd: number;
  } | null = null;

  const flushBlock = () => {
    if (!currentBlock || !currentBlock.title.trim()) return;
    blocks.push({
      id: generateId(),
      sourceFile: file.name,
      headingLevel: currentBlock.headingLevel,
      title: currentBlock.title,
      bodyText: currentBlock.bodyParts.join(" ").trim(),
      parentHeadings: currentBlock.parentHeadings,
      paragraphStart: currentBlock.paragraphStart,
      paragraphEnd: currentBlock.paragraphEnd,
    });
    currentBlock = null;
  };

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    const level = getHeadingLevelFromElement(el);
    const text = extractTextFromElement(el).trim();

    if (level === 1) {
      flushBlock();
      h1 = text;
      h2 = null;
      currentBlock = {
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
      currentBlock = {
        headingLevel: 2,
        title: text,
        parentHeadings: h1 ? [h1] : [],
        bodyParts: [],
        paragraphStart: i,
        paragraphEnd: i,
      };
    } else if (level === 3) {
      flushBlock();
      currentBlock = {
        headingLevel: 3,
        title: text,
        parentHeadings: [h1, h2].filter((x): x is string => x !== null),
        bodyParts: [],
        paragraphStart: i,
        paragraphEnd: i,
      };
    } else {
      if (currentBlock) {
        currentBlock.paragraphEnd = i;
        if (text) currentBlock.bodyParts.push(text);
      }
    }
  }

  flushBlock();
  return blocks;
}
