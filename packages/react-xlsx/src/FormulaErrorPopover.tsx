import { useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import type { ViewerPalette } from "./viewer-palette";

export function FormulaErrorPopover({ text, x, y, palette, className }: {
  text: string;
  x: number;
  y: number;
  palette: ViewerPalette;
  className?: string;
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
    <div ref={ref} role="tooltip" data-xlsx-formula-tooltip="true" className={className} style={{
      position: "fixed", ...position, zIndex: 1000, pointerEvents: "none",
      maxWidth: "min(256px, calc(100vw - 16px))", maxHeight: "calc(100vh - 16px)", overflow: "hidden",
      overflowWrap: "anywhere",
      ...(!className ? {
        padding: "4px 8px", borderRadius: 6, border: `1px solid ${palette.border}`,
        background: palette.surface, color: palette.text, boxShadow: palette.shadow,
        fontFamily: "system-ui, sans-serif", fontSize: 12, lineHeight: "16px"
      } : {})
    }}>{text}</div>, document.body
  );
}
