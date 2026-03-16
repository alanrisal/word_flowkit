import JSZip from "jszip";

export interface DebateBlock {
  id: string;
  sourceFile: string;
  headingLevel: 1 | 2 | 3;
  title: string;
  bodyText: string;
  rawOoxml: string;
  parentHeadings: string[];
}

function generateId(): string {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID();
  }
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Extract all <w:p> nodes from document.xml as serialized XML strings.
 * Uses DOMParser so tag boundaries are always correct, then XMLSerializer
 * to produce a clean, self-contained string for each paragraph.
 */
async function extractParagraphXml(file: File): Promise<string[]> {
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

  const serializer = new XMLSerializer();
  const paragraphs: string[] = [];

  for (const child of Array.from(bodyEl.childNodes)) {
    if (child.nodeType === Node.ELEMENT_NODE && (child as Element).localName === "p") {
      paragraphs.push(serializer.serializeToString(child as Element));
    }
  }

  return paragraphs;
}

/**
 * Extract plain text from all <w:t> nodes in a <w:p> XML string.
 */
function extractText(wpNode: string): string {
  const textRe = /<w:t(?:[^>]*)>([\s\S]*?)<\/w:t>/g;
  const parts: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = textRe.exec(wpNode)) !== null) {
    parts.push(m[1]);
  }
  return parts.join("");
}

/**
 * Detect the heading level (1, 2, or 3) from a <w:p> node's style.
 * Returns null for body paragraphs.
 */
function getHeadingLevel(wpNode: string): 1 | 2 | 3 | null {
  // Match <w:pStyle w:val="Heading1"/> or "Heading 1" or "heading1" etc.
  const styleMatch = wpNode.match(/<w:pStyle\s+w:val="([^"]+)"/i);
  if (!styleMatch) return null;
  const val = styleMatch[1].toLowerCase().replace(/\s/g, "");
  if (val === "heading1") return 1;
  if (val === "heading2") return 2;
  if (val === "heading3") return 3;
  return null;
}

export async function parseDebateFile(file: File): Promise<DebateBlock[]> {
  const paragraphXmls = await extractParagraphXml(file);

  const blocks: DebateBlock[] = [];

  // Current heading context for breadcrumbs
  let h1: string | null = null;
  let h2: string | null = null;

  // Accumulator for the block currently being built
  let currentBlock: {
    headingLevel: 1 | 2 | 3;
    title: string;
    parentHeadings: string[];
    paragraphs: string[];
    bodyParts: string[];
  } | null = null;

  const flushBlock = () => {
    if (!currentBlock) return;
    if (!currentBlock.title.trim()) return; // skip empty heading blocks
    blocks.push({
      id: generateId(),
      sourceFile: file.name,
      headingLevel: currentBlock.headingLevel,
      title: currentBlock.title,
      bodyText: currentBlock.bodyParts.join(" ").trim(),
      rawOoxml: currentBlock.paragraphs.join("\n"),
      parentHeadings: currentBlock.parentHeadings,
    });
  };

  for (const pXml of paragraphXmls) {
    const level = getHeadingLevel(pXml);
    const text = extractText(pXml).trim();

    if (level === 1) {
      flushBlock();
      h1 = text;
      h2 = null;
      currentBlock = {
        headingLevel: 1,
        title: text,
        parentHeadings: [],
        paragraphs: [pXml],
        bodyParts: [],
      };
    } else if (level === 2) {
      flushBlock();
      h2 = text;
      currentBlock = {
        headingLevel: 2,
        title: text,
        parentHeadings: h1 ? [h1] : [],
        paragraphs: [pXml],
        bodyParts: [],
      };
    } else if (level === 3) {
      flushBlock();
      currentBlock = {
        headingLevel: 3,
        title: text,
        parentHeadings: [h1, h2].filter((x): x is string => x !== null),
        paragraphs: [pXml],
        bodyParts: [],
      };
    } else {
      // Body paragraph — append to current block if one is open
      if (currentBlock) {
        currentBlock.paragraphs.push(pXml);
        if (text) currentBlock.bodyParts.push(text);
      }
    }
  }

  // Flush the last block
  flushBlock();

  return blocks;
}
