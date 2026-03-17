/* global Word */

/**
 * Import any styles from the reference document that are missing in the
 * active (target) document. Called once when a file is loaded — never
 * during paste so it never blocks the hot path.
 *
 * Failures for individual styles are non-fatal; Word will fall back to
 * Normal when it encounters an unknown style name in the OOXML.
 */
export async function importStylesIntoActiveDocument(
  base64: string,
  styleNames: string[]
): Promise<void> {
  await Word.run(async (context) => {
    // Find which styles are already present in the target document
    const targetStyles = context.document.getStyles();
    targetStyles.load("items/nameLocal");
    await context.sync();

    const existing = new Set(targetStyles.items.map(s => s.nameLocal));
    const missing = styleNames.filter(n => !existing.has(n));

    if (missing.length === 0) {
      console.log("[FlowKit] All styles already present, skipping import");
      return;
    }

    console.log(`[FlowKit] Importing ${missing.length} style(s):`, missing);

    // Open the reference document once to read its style definitions
    const refDoc = context.application.createDocument(base64);
    context.load(refDoc);
    await context.sync();

    for (const styleName of missing) {
      try {
        const refStyle = refDoc.getStyles().getByNameOrNullObject(styleName);
        refStyle.load("type,nameLocal");
        await context.sync();
        if (!refStyle.isNullObject) {
          context.document.addStyle(styleName, refStyle.type);
          await context.sync();
          console.log(`[FlowKit] Imported style: "${styleName}"`);
        }
      } catch (e) {
        console.warn(`[FlowKit] Skipping style "${styleName}":`, e);
      }
    }
  });
}
