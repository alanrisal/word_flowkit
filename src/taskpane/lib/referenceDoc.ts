/* global Word */

export interface BlockIndex {
  title: string;
  headingLevel: 1 | 2 | 3;
  parentHeadings: string[];
  paragraphStart: number; // 0-indexed position in body.paragraphs
  paragraphEnd: number;   // inclusive
  bodyText: string;       // plain text for search
  sourceFile: string;
}

function getHeadingLevel(styleBuiltIn: string): 1 | 2 | 3 | null {
  // styleBuiltIn returns locale-independent names: "Heading1", "Heading2", "Heading3"
  if (styleBuiltIn === "Heading1") return 1;
  if (styleBuiltIn === "Heading2") return 2;
  if (styleBuiltIn === "Heading3") return 3;
  return null;
}

/**
 * Convert a File to a base64 string and open it as a Word document object
 * (without displaying it) to build a block index with accurate paragraph
 * positions — positions that will match paster.ts when it re-opens the same
 * base64 via context.application.createDocument().
 */
export async function loadReferenceFile(file: File): Promise<{
  base64: string;
  blocks: BlockIndex[];
}> {
  // Step 1: Read file as base64
  const arrayBuffer = await file.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  const base64 = btoa(binary);

  // Step 2: Open the document via Word API and walk paragraphs
  const blocks: BlockIndex[] = [];

  await Word.run(async (context) => {
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
        title: current.title,
        headingLevel: current.headingLevel,
        parentHeadings: current.parentHeadings,
        paragraphStart: current.paragraphStart,
        paragraphEnd: current.paragraphEnd,
        bodyText: current.bodyParts.join(" ").trim(),
        sourceFile: file.name,
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
  });

  return { base64, blocks };
}
