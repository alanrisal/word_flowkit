import { BlockIndex } from "./referenceDoc";

/* global Word */

export interface PasteTarget {
  block: BlockIndex;
  base64: string; // the reference file as base64
}

export async function pasteBlock(target: PasteTarget): Promise<void> {
  try {
    await Word.run(async (context) => {

      // STEP 1: Open reference document (not displayed to user)
      const refDoc = context.application.createDocument(target.base64);
      context.load(refDoc, "body");
      await context.sync();

      // STEP 2: Get all paragraphs from reference doc
      const paragraphs = refDoc.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      // STEP 3: Build a range covering paragraphStart → paragraphEnd
      const items = paragraphs.items;
      const startPara = items[target.block.paragraphStart];
      const endPara = items[target.block.paragraphEnd];

      const blockRange = startPara
        .getRange(Word.RangeLocation.whole)
        .expandTo(endPara.getRange(Word.RangeLocation.whole));

      // STEP 4: Get OOXML from Word — Word generates this, not us
      const ooxmlResult = blockRange.getOoxml();
      await context.sync();
      const ooxml = ooxmlResult.value;

      // STEP 5: Get the target document (the one the user is editing)
      const targetDoc = context.document;

      // STEP 6: Import styles that exist in reference but not in target
      const targetStyles = targetDoc.getStyles();
      targetStyles.load("items/nameLocal");
      await context.sync();
      const existingStyleNames = new Set(
        targetStyles.items.map(s => s.nameLocal)
      );

      const blockParas = blockRange.paragraphs;
      blockParas.load("items/style");
      await context.sync();
      const neededStyles = new Set(
        blockParas.items.map(p => p.style)
      );

      for (const styleName of neededStyles) {
        if (!existingStyleNames.has(styleName)) {
          console.log(`[FlowKit] Importing missing style: "${styleName}"`);
          try {
            const refStyle = refDoc.getStyles().getByNameOrNullObject(styleName);
            refStyle.load("nameLocal,type");
            await context.sync();
            if (!refStyle.isNullObject) {
              targetDoc.addStyle(styleName, refStyle.type);
              await context.sync();
            }
          } catch (e) {
            console.warn(`[FlowKit] Could not import style "${styleName}":`, e);
            // Non-fatal — continue with paste even if style import fails
          }
        }
      }

      // STEP 7: Insert at cursor in target document
      const selection = targetDoc.getSelection();
      selection.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();

      console.log(`[FlowKit] Paste succeeded: "${target.block.title}"`);
    });
  } catch (err) {
    console.error(`[FlowKit] Paste failed: ${target.block.title}`, err);
    throw err;
  }
}
