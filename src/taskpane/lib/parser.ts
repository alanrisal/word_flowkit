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
 * Strip xmlns:* declarations that XMLSerializer adds to every serialized element.
 * These are already declared on the OOXML wrapper, so leaving them on each <w:p>
 * creates redundant re-declarations that Word's strict XML parser rejects.
 */
function cleanSerializedParagraph(xml: string): string {
  return xml.replace(/\s+xmlns(?::[A-Za-z0-9_.-]+)?="[^"]*"/g, "");
}

/**
 * Detect heading level from a <w:p> DOM element.
 * Using the DOM directly avoids regex failure if XMLSerializer renamed the w: prefix.
 */
function getHeadingLevelFromElement(el: Element): 1 | 2 | 3 | null {
  const pStyleEls = el.getElementsByTagNameNS(W_NS, "pStyle");
  if (!pStyleEls.length) return null;
  const pStyle = pStyleEls[0];
  // w:val may be an NS attribute or plain attribute depending on serializer
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

/**
 * Extract visible text from a <w:p> DOM element via <w:t> nodes.
 */
function extractTextFromElement(el: Element): string {
  const tNodes = el.getElementsByTagNameNS(W_NS, "t");
  return Array.from(tNodes)
    .map(t => t.textContent ?? "")
    .join("");
}

interface ParagraphEntry {
  element: Element;
  xml: string; // cleaned, xmlns-stripped serialization
}

/**
 * Parse document.xml from a .docx zip, returning each direct <w:p> child of
 * <w:body> as both a DOM Element (for querying) and a cleaned XML string (for storage).
 */
async function extractParagraphs(file: File): Promise<ParagraphEntry[]> {
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
  const results: ParagraphEntry[] = [];

  for (const child of Array.from(bodyEl.childNodes)) {
    if (child.nodeType === Node.ELEMENT_NODE && (child as Element).localName === "p") {
      const el = child as Element;
      const raw = serializer.serializeToString(el);
      results.push({ element: el, xml: cleanSerializedParagraph(raw) });
    }
  }

  return results;
}

export async function parseDebateFile(file: File): Promise<DebateBlock[]> {
  const paragraphs = await extractParagraphs(file);

  const blocks: DebateBlock[] = [];

  let h1: string | null = null;
  let h2: string | null = null;

  let currentBlock: {
    headingLevel: 1 | 2 | 3;
    title: string;
    parentHeadings: string[];
    xmlParts: string[];
    bodyParts: string[];
  } | null = null;

  const flushBlock = () => {
    if (!currentBlock || !currentBlock.title.trim()) return;
    blocks.push({
      id: generateId(),
      sourceFile: file.name,
      headingLevel: currentBlock.headingLevel,
      title: currentBlock.title,
      bodyText: currentBlock.bodyParts.join(" ").trim(),
      rawOoxml: currentBlock.xmlParts.join("\n"),
      parentHeadings: currentBlock.parentHeadings,
    });
  };

  for (const { element, xml } of paragraphs) {
    const level = getHeadingLevelFromElement(element);
    const text = extractTextFromElement(element).trim();

    if (level === 1) {
      flushBlock();
      h1 = text;
      h2 = null;
      currentBlock = {
        headingLevel: 1,
        title: text,
        parentHeadings: [],
        xmlParts: [xml],
        bodyParts: [],
      };
    } else if (level === 2) {
      flushBlock();
      h2 = text;
      currentBlock = {
        headingLevel: 2,
        title: text,
        parentHeadings: h1 ? [h1] : [],
        xmlParts: [xml],
        bodyParts: [],
      };
    } else if (level === 3) {
      flushBlock();
      currentBlock = {
        headingLevel: 3,
        title: text,
        parentHeadings: [h1, h2].filter((x): x is string => x !== null),
        xmlParts: [xml],
        bodyParts: [],
      };
    } else {
      if (currentBlock) {
        currentBlock.xmlParts.push(xml);
        if (text) currentBlock.bodyParts.push(text);
      }
    }
  }

  flushBlock();

  return blocks;
}
