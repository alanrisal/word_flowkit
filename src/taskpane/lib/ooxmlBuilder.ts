/**
 * Validates an OOXML string by parsing it with DOMParser.
 * Returns { valid: true } or { valid: false, error: string }.
 * Call this before every insertOoxml() to get a readable error instead of
 * Word's cryptic "line 1 column N" message.
 */
export function validateOoxml(ooxml: string): { valid: boolean; error?: string } {
  const parser = new DOMParser();
  const doc = parser.parseFromString(ooxml, "application/xml");
  const err = doc.querySelector("parsererror");
  if (err) {
    return { valid: false, error: err.textContent ?? "Unknown XML parse error" };
  }
  return { valid: true };
}

/**
 * Wraps cleaned <w:p> paragraph strings in a complete OOXML document fragment
 * suitable for Word.Selection.insertOoxml().
 *
 * Namespace coverage:
 * - Core WordprocessingML namespaces (w, r, m, v, wp, w10)
 * - Office namespaces (o, mc)
 * - Word 2010 extension namespaces (w14, wp14, wpc, wpg, wpi, wps, wne)
 * - Word 2012–2018 extension namespaces (w15, w16cex, w16se)
 *
 * The paragraphs passed in must have their own xmlns:* declarations stripped
 * (see cleanSerializedParagraph in parser.ts) so there are no re-declaration conflicts.
 */
export function buildOoxmlDocument(rawParagraphs: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:wordprocessingML
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
  xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  mc:Ignorable="w14 w15 w16cex w16se wp14">
  <w:body>
    ${rawParagraphs}
  </w:body>
</w:wordprocessingML>`;
}
