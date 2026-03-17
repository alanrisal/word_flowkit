import { DebateBlock } from "./parser";
import { sanitizeOoxml } from "./ooxmlBuilder";
import { getDocumentStyles } from "./styleResolver";

/* global Word */

export async function pasteBlock(block: DebateBlock): Promise<void> {
  // 1. Get styles present in the destination document
  const knownStyles = await getDocumentStyles();

  // 2. Sanitize: strip tracking attrs, remap unknown styles, wrap in OOXML document
  const ooxml = sanitizeOoxml(block.rawOoxml, knownStyles);

  // 3. Validate the final XML before handing it to Word
  const check = new DOMParser().parseFromString(ooxml, "application/xml");
  if (check.querySelector("parsererror")) {
    console.error("[FlowKit] Invalid OOXML after sanitize:", ooxml);
    throw new Error("Block XML is malformed after sanitization — check console for details");
  }

  // 4 & 5. Insert at cursor, log and rethrow on failure
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (err) {
    console.error(`[FlowKit] Paste failed for block: ${block.title}`, err);
    throw err;
  }
}
