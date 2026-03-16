import { DebateBlock } from "./parser";
import { buildOoxmlDocument } from "./ooxmlBuilder";

// Word is provided as an ambient global by @types/office-js.
// The eslint-disable comment suppresses "no-undef" if any linter is configured.
/* global Word */

export async function pasteBlockAtCursor(block: DebateBlock): Promise<void> {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const ooxml = buildOoxmlDocument(block.rawOoxml);
    selection.insertOoxml(ooxml, Word.InsertLocation.replace);
    await context.sync();
  });
}
