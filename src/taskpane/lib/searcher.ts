import Fuse, { type IFuseOptions, type FuseResult } from "fuse.js";
import { DebateBlock } from "./parser";

export interface RankedResult {
  block: DebateBlock;
  score: number;
  matchedIn: "title" | "body" | "breadcrumb";
}

const fuseOptions: IFuseOptions<DebateBlock> = {
  keys: [
    { name: "title", weight: 0.7 },
    { name: "bodyText", weight: 0.2 },
    { name: "parentHeadings", weight: 0.1 },
  ],
  threshold: 0.4,
  includeScore: true,
  includeMatches: true,
  useExtendedSearch: true,
  ignoreLocation: true,
  minMatchCharLength: 2,
};

function inferMatchLocation(
  result: FuseResult<DebateBlock>
): "title" | "body" | "breadcrumb" {
  const matches = result.matches ?? [];
  for (const m of matches) {
    if (m.key === "title") return "title";
    if (m.key === "bodyText") return "body";
  }
  return "breadcrumb";
}

export function searchBlocks(query: string, blocks: DebateBlock[]): RankedResult[] {
  if (!query.trim() || blocks.length === 0) return [];
  const fuse = new Fuse(blocks, fuseOptions);
  return fuse
    .search(query)
    .slice(0, 20)
    .map(r => ({
      block: r.item,
      score: 1 - (r.score ?? 0),
      matchedIn: inferMatchLocation(r),
    }));
}
