/**
 * DebateBlock is the unified type for both search and paste.
 * Re-exported from referenceDoc so all existing importers (searcher,
 * BlockList, BlockPreview, fileStore) continue to work unchanged.
 */
export type { BlockIndex as DebateBlock } from "./referenceDoc";
