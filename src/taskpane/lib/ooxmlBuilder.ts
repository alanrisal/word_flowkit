const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

const HEADING_RE = /^heading\s*[1-6]$/i;

// Wraps a serialized paragraph string, stripping xmlns declarations from its opening tag only.
function cleanOpeningTag(xml: string): string {
  // Match the opening <w:p ...> tag (up to the first > or end of self-close)
  return xml.replace(/^(<w:p\b[^>]*>)/, (tag) =>
    tag.replace(/\s+xmlns:[a-zA-Z0-9]+="[^"]*"/g, "")
  );
}

export function sanitizeOoxml(rawOoxml: string, knownStyles: Set<string>): string {
  // Step 1 — parse inside a root element with all needed namespace declarations
  const wrapped = `<root
    xmlns:w="${W_NS}"
    xmlns:w14="${W14_NS}"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  >${rawOoxml}</root>`;

  const doc = new DOMParser().parseFromString(wrapped, "application/xml");
  const parseErr = doc.querySelector("parsererror");
  if (parseErr) {
    console.error("[FlowKit] sanitizeOoxml: could not parse rawOoxml:", parseErr.textContent);
    // Return rawOoxml as-is; paster.ts will catch the subsequent validation failure
    return rawOoxml;
  }

  const paragraphs = Array.from(doc.documentElement.getElementsByTagNameNS(W_NS, "p"));

  for (const p of paragraphs) {
    // Step 2 — remove revision/session tracking attributes from <w:p> AND all
    // descendant elements (e.g. <w:r w:rsidRPr="...">). Word rejects these when
    // they appear in an insertOoxml fragment from a different document session.
    const allEls = [p, ...Array.from(p.getElementsByTagName("*"))];
    for (const el of allEls) {
      el.removeAttributeNS(W14_NS, "paraId");
      el.removeAttributeNS(W14_NS, "textId");
      el.removeAttributeNS(W_NS, "rsidR");
      el.removeAttributeNS(W_NS, "rsidRPr");
      el.removeAttributeNS(W_NS, "rsidRDefault");
      el.removeAttributeNS(W_NS, "rsidP");
      el.removeAttributeNS(W_NS, "rsidDel");
      el.removeAttribute("w14:paraId");
      el.removeAttribute("w14:textId");
      el.removeAttribute("w:rsidR");
      el.removeAttribute("w:rsidRPr");
      el.removeAttribute("w:rsidRDefault");
      el.removeAttribute("w:rsidP");
      el.removeAttribute("w:rsidDel");
    }

    // Step 3 — remap unknown paragraph styles to Normal.
    // getAttributeNS is used to read (w:val is in the W namespace after DOM parsing).
    // setAttribute with the qualified name is used to write — using both setAttributeNS
    // AND setAttribute on the same attribute creates duplicate attributes in Edge's
    // DOM serializer and produces malformed XML.
    const pStyleEls = p.getElementsByTagNameNS(W_NS, "pStyle");
    for (const pStyle of Array.from(pStyleEls)) {
      const val = pStyle.getAttributeNS(W_NS, "val") ?? pStyle.getAttribute("w:val") ?? "";
      if (!HEADING_RE.test(val) && !knownStyles.has(val)) {
        console.log(`[FlowKit] Remapping unknown style "${val}" → "Normal"`);
        pStyle.setAttribute("w:val", "Normal");
      }
    }
  }

  // Step 3b — strip any <w:sectPr> elements carried over from the source document.
  // sectPr defines document/section layout and is invalid inside an insertOoxml fragment.
  const sectPrs = doc.getElementsByTagNameNS(W_NS, "sectPr");
  for (const s of Array.from(sectPrs)) {
    s.parentNode?.removeChild(s);
  }

  // Step 4 — serialize and strip redundant xmlns from opening tags
  const serializer = new XMLSerializer();
  const cleaned = paragraphs
    .map((p) => cleanOpeningTag(serializer.serializeToString(p)))
    .join("\n");

  // Step 5 — wrap and return
  return wrapOoxml(cleaned);
}

export function wrapOoxml(paragraphs: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:wordprocessingML
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  mc:Ignorable="w14 w15">
  <w:body>
    ${paragraphs}
    <w:sectPr/>
  </w:body>
</w:wordprocessingML>`;
}
