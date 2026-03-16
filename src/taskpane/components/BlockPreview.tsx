import React from "react";
import { DebateBlock } from "../lib/parser";

interface Props {
  block: DebateBlock;
}

export default function BlockPreview({ block }: Props) {
  const PREVIEW_LIMIT = 300;
  const preview = block.bodyText.slice(0, PREVIEW_LIMIT);
  const isTruncated = block.bodyText.length > PREVIEW_LIMIT;

  if (!block.bodyText.trim()) {
    return (
      <div className="block-preview">
        <p className="block-preview-empty">(No body text)</p>
      </div>
    );
  }

  return (
    <div className="block-preview">
      <p>
        {preview}
        {isTruncated && <span className="preview-ellipsis">…</span>}
      </p>
    </div>
  );
}
