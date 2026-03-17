import { BlockIndex } from "./referenceDoc";

/* global Word */

/**
 * Paste a block at the current cursor position.
 * Fast path: block.cachedOoxml was extracted at file-load time,
 * so this is a single insertOoxml call with no document re-opening.
 */
export async function pasteBlock(block: BlockIndex): Promise<void> {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertOoxml(block.cachedOoxml, Word.InsertLocation.replace);
      await context.sync();
    });
    console.log(`[FlowKit] Paste succeeded: "${block.title}"`);
  } catch (err) {
    console.error(`[FlowKit] Paste failed: ${block.title}`, err);
    throw err;
  }
}
