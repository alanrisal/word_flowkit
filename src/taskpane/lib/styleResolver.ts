/* global Word */

export async function getDocumentStyles(): Promise<Set<string>> {
  return Word.run(async (context) => {
    const styles = context.document.getStyles();
    styles.load("items/nameLocal, items/builtIn");
    await context.sync();
    const names = new Set<string>();
    for (const style of styles.items) {
      if (style.nameLocal) names.add(style.nameLocal);
    }
    return names;
  });
}
