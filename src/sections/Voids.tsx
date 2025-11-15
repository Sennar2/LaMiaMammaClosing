// src/sections/Voids.tsx
import React, { useMemo, useState } from "react";
import YesNoWithComment, { YesNo } from "../components/YesNoWithComment";

type Props = {
  onCanProceed?: (ok: boolean) => void;
};

export default function VoidsSection({ onCanProceed }: Props) {
  const [value, setValue] = useState<YesNo>("");
  const [comment, setComment] = useState("");

  const ok = useMemo(() => !(value === "no" && !comment.trim()), [value, comment]);

  React.useEffect(() => {
    onCanProceed?.(ok);
  }, [ok, onCanProceed]);

  return (
    <section className="card">
      <h2>Voids</h2>
      <YesNoWithComment
        label="Any voids without manager approval?"
        value={value}
        comment={comment}
        onChange={(v, c) => {
          setValue(v);
          setComment(c ?? "");
        }}
        requireCommentOnNo
      />
      {value === "no" && !comment.trim() && (
        <div className="blocker">Add a comment to continue.</div>
      )}
    </section>
  );
}
