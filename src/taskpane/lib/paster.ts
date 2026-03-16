import { DebateBlock } from "./parser";
import { buildOoxmlDocument, validateOoxml } from "./ooxmlBuilder";

/* global Word */

export async function pasteBlockAtCursor(block: DebateBlock): Promise<void> {
  const ooxml = buildOoxmlDocument(block.rawOoxml);

  // Validate before sending to Word — gives a readable error instead of
  // "contents have a problem" with a cryptic column number.
  const check = validateOoxml(ooxml);
  if (!check.valid) {
    console.error("[FlowKit] OOXML failed validation before paste.");
    console.error("[FlowKit] Parse error:", check.error);
    // Log the raw paragraphs so you can inspect which block is broken
    console.error("[FlowKit] rawOoxml for block:", block.title);
    console.error("[FlowKit] rawOoxml content:", block.rawOoxml);
    throw new Error(`Paste failed — malformed OOXML: ${check.error}`);
  }

  console.log("[FlowKit] OOXML valid, inserting block:", block.title);

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (wordErr) {
    console.error("[FlowKit] Word.run failed for block:", block.title);
    console.error("[FlowKit] Word error:", wordErr);
    console.error("[FlowKit] rawOoxml that caused failure:", block.rawOoxml);
    throw wordErr;
  }
}
