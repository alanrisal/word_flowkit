import { BlockIndex } from "./referenceDoc";

/* global Word */

/**
 * Paste a block at the current cursor position.
 * If stylesXml is provided (raw word/styles.xml from the source .docx), it is
 * injected into the pkg:package OOXML so Word uses the source style definitions
 * rather than plain Normal-based shells.
 */
export async function pasteBlock(block: BlockIndex, stylesXml?: string): Promise<void> {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const ooxml = stylesXml
        ? injectStylesIntoOoxml(block.cachedOoxml, stylesXml)
        : block.cachedOoxml;
      selection.insertOoxml(ooxml, Word.InsertLocation.replace);
      await context.sync();
    });
    console.log(`[FlowKit] Paste succeeded: "${block.title}"`);
  } catch (err) {
    console.error(`[FlowKit] Paste failed: ${block.title}`, err);
    throw err;
  }
}

/**
 * Inject word/styles.xml into a pkg:package OOXML string so that pasted
 * content carries its source style definitions.
 *
 * getOoxml() returns pkg:package format.  We add (or replace) the
 * /word/styles.xml part so Word resolves style names against the source
 * definitions instead of whatever the target document has.
 */
function injectStylesIntoOoxml(ooxml: string, stylesXml: string): string {
  if (!ooxml.includes("<pkg:package")) return ooxml; // not pkg format — leave untouched

  // Strip any standalone XML declaration so the content can be embedded inline
  const stylesBody = stylesXml.replace(/<\?xml[^?]*\?>\s*/i, "");

  const stylesPart =
    `<pkg:part pkg:name="/word/styles.xml" ` +
    `pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml">` +
    `<pkg:xmlData>${stylesBody}</pkg:xmlData>` +
    `</pkg:part>`;

  // Replace an existing styles part, or append before the closing tag
  if (/<pkg:part[^>]*pkg:name="\/word\/styles\.xml"/.test(ooxml)) {
    return ooxml.replace(
      /<pkg:part[^>]*pkg:name="\/word\/styles\.xml"[\s\S]*?<\/pkg:part>/,
      stylesPart
    );
  }
  return ooxml.replace("</pkg:package>", `${stylesPart}</pkg:package>`);
}
