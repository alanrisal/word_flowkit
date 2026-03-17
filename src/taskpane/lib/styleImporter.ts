/* global Word */

/**
 * Import any styles from the reference document that are missing in the
 * active (target) document. Called once when a file is loaded — never
 * during paste so it never blocks the hot path.
 *
 * styleSource has two modes:
 *   { stylesJson } — fast path: styles JSON was pre-extracted at load time,
 *                    no createDocument() needed at all.
 *   { base64 }    — slow path: open the reference doc and export styles now.
 *                   Falls back to addStyle() if exportStylesFromJson is unavailable.
 */
export async function importStylesIntoActiveDocument(
  styleSource: { base64?: string; stylesJson?: string },
  styleNames: string[]
): Promise<void> {
  if (styleNames.length === 0) return;

  if (styleSource.stylesJson) {
    await importFromCachedJson(styleSource.stylesJson, styleNames);
  } else if (styleSource.base64) {
    await importFromReferenceDoc(styleSource.base64, styleNames);
  }
}

// ─── Fast path ───────────────────────────────────────────────────────────────

/**
 * stylesJson already extracted — no document re-opening needed.
 * Just filter to missing styles and call importStylesFromJson.
 */
async function importFromCachedJson(
  stylesJson: string,
  styleNames: string[]
): Promise<void> {
  await Word.run(async (context) => {
    const targetStyles = context.document.getStyles();
    targetStyles.load("items/nameLocal");
    await context.sync();

    const existing = new Set(targetStyles.items.map(s => s.nameLocal));
    const missing = styleNames.filter(n => !existing.has(n));

    if (missing.length === 0) {
      console.log("[FlowKit Styles] All styles present, nothing to import");
      return;
    }

    let allStyles: unknown[];
    try {
      allStyles = JSON.parse(stylesJson) as unknown[];
    } catch (e) {
      console.error("[FlowKit Styles] Failed to parse cached styles JSON:", e);
      return;
    }

    const missingSet = new Set(missing);
    const filtered = allStyles.filter(
      (s: any) => missingSet.has(s.name) || missingSet.has(s.nameLocal)
    );

    if (filtered.length === 0) {
      console.warn("[FlowKit Styles] None of the missing styles found in cached JSON:", missing);
      return;
    }

    console.log(`[FlowKit Styles] Importing ${filtered.length} style(s) from cache`);
    context.document.importStylesFromJson(JSON.stringify(filtered));
    await context.sync();
    console.log("[FlowKit Styles] Import complete");
  });
}

// ─── Slow path ────────────────────────────────────────────────────────────────

/**
 * No pre-extracted JSON available — open the reference document and try
 * exportStylesFromJson() (WordApiDesktop 1.1+).
 * Falls back to addStyle() for older Word versions.
 */
async function importFromReferenceDoc(
  base64: string,
  styleNames: string[]
): Promise<void> {
  await Word.run(async (context) => {
    const targetStyles = context.document.getStyles();
    targetStyles.load("items/nameLocal");
    await context.sync();

    const existing = new Set(targetStyles.items.map(s => s.nameLocal));
    const missing = styleNames.filter(n => !existing.has(n));

    if (missing.length === 0) {
      console.log("[FlowKit Styles] All styles present, nothing to import");
      return;
    }

    console.log(`[FlowKit Styles] Need to import: ${missing.join(", ")}`);

    const refDoc = context.application.createDocument(base64);
    context.load(refDoc);
    await context.sync();

    try {
      // exportStylesFromJson is not in @types/office-js — runtime-only in newer Word.
      // Cast to any; the sync will throw ApiNotAvailable on older versions.
      const exportResult = (refDoc as any).exportStylesFromJson() as
        OfficeExtension.ClientResult<string>;
      await context.sync();

      let allStyles: unknown[];
      try {
        allStyles = JSON.parse(exportResult.value) as unknown[];
      } catch (e) {
        console.error("[FlowKit Styles] Failed to parse exported styles JSON:", e);
        return;
      }

      const missingSet = new Set(missing);
      const filtered = allStyles.filter(
        (s: any) => missingSet.has(s.name) || missingSet.has(s.nameLocal)
      );

      if (filtered.length === 0) {
        console.warn("[FlowKit Styles] None of the missing styles found in reference doc:", missing);
        return;
      }

      console.log(`[FlowKit Styles] Importing ${filtered.length} complete style definition(s)`);
      context.document.importStylesFromJson(JSON.stringify(filtered));
      await context.sync();
      console.log("[FlowKit Styles] Import complete");

    } catch (e: unknown) {
      if (isApiUnavailable(e)) {
        console.warn("[FlowKit Styles] exportStylesFromJson not available, using addStyle fallback");
        await fallbackStyleImport(missing, refDoc, context);
      } else {
        throw e;
      }
    }
  });
}

// ─── Fallback ─────────────────────────────────────────────────────────────────

/**
 * addStyle() fallback for Word versions that don't support exportStylesFromJson.
 * Creates empty style shells that inherit from Normal — no font/color fidelity,
 * but prevents Word from stripping unknown style names from pasted OOXML.
 *
 * Note: getStyles() is not typed on DocumentCreated but exists at runtime;
 * we cast to any to call it.
 */
async function fallbackStyleImport(
  missing: string[],
  refDoc: Word.DocumentCreated,
  context: Word.RequestContext
): Promise<void> {
  for (const styleName of missing) {
    try {
      const refStyle = (refDoc as any).getStyles().getByNameOrNullObject(styleName) as Word.Style;
      refStyle.load("type");
      await context.sync();
      if (!refStyle.isNullObject) {
        context.document.addStyle(styleName, refStyle.type);
        await context.sync();
        console.log(`[FlowKit Styles] Fallback: added empty style shell "${styleName}"`);
      }
    } catch (e) {
      console.warn(`[FlowKit Styles] Fallback failed for "${styleName}":`, e);
    }
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function isApiUnavailable(e: unknown): boolean {
  const s = String(e);
  return s.includes("ApiNotAvailable") || s.includes("not implemented");
}
