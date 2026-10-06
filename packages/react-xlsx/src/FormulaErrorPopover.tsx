import { useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { ViewerPalette } from "./viewer-palette";

export function FormulaErrorPopover({ text, x, y, palette }: {
  text: string;
  x: number;
  y: number;
  palette: ViewerPalette;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const [position, setPosition] = useState({ left: x, top: y });
  useLayoutEffect(() => {
    const rect = ref.current?.getBoundingClientRect();
    if (!rect) return;
    setPosition({
      left: Math.max(8, Math.min(x + 4, window.innerWidth - rect.width - 8)),
      top: Math.max(8, y + 4 + rect.height < window.innerHeight - 8 ? y + 4 : y - rect.height - 4)
    });
  }, [x, y, text]);

  return createPortal(
    <div ref={ref} role="tooltip" data-xlsx-formula-tooltip="true" style={{
      position: "fixed", ...position, zIndex: 1000, pointerEvents: "none",
      maxWidth: "min(220px, calc(100vw - 16px))", maxHeight: "calc(100vh - 16px)", overflow: "hidden",
      padding: "6px 8px", borderRadius: 5, border: `1px solid ${palette.border}`,
      background: palette.surface, color: palette.text, boxShadow: "0 4px 16px rgba(0,0,0,0.15)",
      fontFamily: "system-ui, sans-serif", fontSize: 11, lineHeight: 1.4,
      whiteSpace: "pre-line", overflowWrap: "anywhere"
    }}>{text}</div>, document.body
  );
}
