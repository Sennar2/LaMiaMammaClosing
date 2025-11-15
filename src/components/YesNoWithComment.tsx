// src/components/YesNoWithComment.tsx
import React, { useEffect, useState } from "react";

export type YesNo = "yes" | "no" | "";

type Props = {
  label: string;
  value: YesNo;
  comment?: string;
  onChange: (value: YesNo, comment?: string) => void;
  requireCommentOnNo?: boolean; // default true
  disabled?: boolean;
};

export default function YesNoWithComment({
  label,
  value,
  comment,
  onChange,
  requireCommentOnNo = true,
  disabled = false,
}: Props) {
  const [localComment, setLocalComment] = useState(comment ?? "");
  const showComment = value === "no";
  const mustComment = requireCommentOnNo && value === "no" && !localComment.trim();

  useEffect(() => setLocalComment(comment ?? ""), [comment]);

  return (
    <div className="ynwc">
      <div className="ynwc-row">
        <label className="ynwc-label">{label}</label>
        <div className="ynwc-buttons">
          <button
            type="button"
            className={`ynwc-btn ${value === "yes" ? "selected" : ""}`}
            onClick={() => onChange("yes", "")}
            disabled={disabled}
          >
            Yes
          </button>
          <button
            type="button"
            className={`ynwc-btn ${value === "no" ? "selected" : ""}`}
            onClick={() => onChange("no", localComment)}
            disabled={disabled}
          >
            No
          </button>
        </div>
      </div>

      {showComment && (
        <div className="ynwc-comment">
          <label>
            <span>Comment (required for “No”)</span>
            <textarea
              value={localComment}
              onChange={(e) => setLocalComment(e.target.value)}
              onBlur={() => onChange("no", localComment)}
              placeholder="Brief reason…"
              rows={2}
              disabled={disabled}
            />
          </label>
          {mustComment && <div className="ynwc-error">Please add a comment to continue.</div>}
        </div>
      )}
    </div>
  );
}
