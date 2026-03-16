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

/**
 * Extract all <w:p>...</w:p> nodes from document.xml as raw XML strings.
 * Uses a state-machine approach to handle nested tags correctly.
 */
async function extractParagraphXml(file: File): Promise<string[]> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) {
    throw new Error(`Could not find word/document.xml inside ${file.name}`);
  }
  const docXml = await docFile.async("string");

  // Extract <w:p> nodes using a regex that handles nested elements
  // We track depth to correctly pair opening and closing tags
  const paragraphs: string[] = [];
  let searchFrom = 0;

  while (searchFrom < docXml.length) {
    // Find the next opening <w:p> tag (with optional attributes or self-closing handled below)
    const openMatch = docXml.indexOf("<w:p", searchFrom);
    if (openMatch === -1) break;

    // Peek at what follows the tag name — could be " ", ">", "/>"
    const afterTag = docXml[openMatch + 4];
    if (afterTag !== " " && afterTag !== ">" && afterTag !== "/") {
      // e.g. <w:pPr> or <w:pStyle> — skip past this character
      searchFrom = openMatch + 5;
      continue;
    }

    // Check for self-closing <w:p ... />
    const closeBracket = docXml.indexOf(">", openMatch);
    if (closeBracket !== -1 && docXml[closeBracket - 1] === "/") {
      paragraphs.push(docXml.slice(openMatch, closeBracket + 1));
      searchFrom = closeBracket + 1;
      continue;
    }

    // Walk forward counting open/close <w:p> tags to find our matching </w:p>
    let depth = 0;
    let pos = openMatch;
    let endPos = -1;

    while (pos < docXml.length) {
      const nextOpen = docXml.indexOf("<w:p", pos);
      const nextClose = docXml.indexOf("</w:p>", pos);

      if (nextClose === -1) break; // malformed XML

      if (nextOpen !== -1 && nextOpen < nextClose) {
        // Check it's actually a <w:p tag (not <w:pPr etc.)
        const ch = docXml[nextOpen + 4];
        if (ch === " " || ch === ">" || ch === "/") {
          // Is it self-closing?
          const cb = docXml.indexOf(">", nextOpen);
          if (cb !== -1 && docXml[cb - 1] === "/") {
            pos = cb + 1;
            continue;
          }
          depth++;
          pos = nextOpen + 5;
        } else {
          pos = nextOpen + 5;
        }
      } else {
        // nextClose comes first
        depth--;
        if (depth === 0) {
          endPos = nextClose + 6; // length of "</w:p>"
          break;
        }
        pos = nextClose + 6;
      }
    }

    if (endPos === -1) break;

    paragraphs.push(docXml.slice(openMatch, endPos));
    searchFrom = endPos;
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
