import * as React from "react";
import { createPortal } from "react-dom";
import type { ViewerPalette } from "./viewer-palette";

type Item = string | number | boolean;

export type ValidationDropdownHandle = { open: () => void };

export const ValidationDropdown = React.forwardRef<ValidationDropdownHandle, {
  label: string; items: Item[] | null; value: string; palette: ViewerPalette; zoom: number;
  onChoose: (value: Item) => void; focusGrid: () => void;
}>(function ValidationDropdown({ label, items, value, palette, zoom, onChoose, focusGrid }, ref) {
  const trigger = React.useRef<HTMLButtonElement>(null);
  const menu = React.useRef<HTMLDivElement>(null);
  const id = React.useId();
  const [position, setPosition] = React.useState<React.CSSProperties | null>(null);
  const [active, setActive] = React.useState(0);
  const fontSize = 12 * zoom;
  const rowHeight = 26 * zoom;
  const close = () => { setPosition(null); focusGrid(); };
  const choose = (index: number) => {
    const item = items?.[index];
    if (item !== undefined) onChoose(item);
    close();
  };
  const open = () => {
    const cell = trigger.current?.parentElement?.getBoundingClientRect();
    if (!cell || !items?.length) return;
    const context = document.createElement("canvas").getContext("2d");
    if (context) context.font = `${fontSize}px system-ui, sans-serif`;
    const longest = items.reduce<number>((width, item) => Math.max(width, context?.measureText(String(item) || "(blank)").width ?? 0), 0);
    const width = Math.min(Math.max(cell.width, longest + 48 * zoom), 340 * zoom, window.innerWidth - 16);
    const below = window.innerHeight - cell.bottom - 8;
    const above = cell.top - 8;
    const height = Math.min(items.length * rowHeight + 8, 240 * zoom, Math.max(below, above));
    setActive(Math.max(0, items.findIndex(item => String(item) === value)));
    setPosition({
      left: Math.max(8, Math.min(cell.left, window.innerWidth - width - 8)),
      top: below >= height ? cell.bottom : Math.max(8, cell.top - height),
      width, maxHeight: height
    });
  };
  React.useImperativeHandle(ref, () => ({ open }));
  React.useLayoutEffect(() => {
    if (!position) return;
    menu.current?.focus({ preventScroll: true });
    const outside = (event: PointerEvent) => {
      if (event.target instanceof Node && !menu.current?.contains(event.target) && !trigger.current?.contains(event.target)) setPosition(null);
    };
    const dismiss = (event: Event) => {
      if (event.target instanceof Node && menu.current?.contains(event.target)) return;
      setPosition(null);
    };
    document.addEventListener("pointerdown", outside, true);
    document.addEventListener("scroll", dismiss, true);
    window.addEventListener("resize", dismiss);
    return () => {
      document.removeEventListener("pointerdown", outside, true);
      document.removeEventListener("scroll", dismiss, true);
      window.removeEventListener("resize", dismiss);
    };
  }, [position]);
  React.useLayoutEffect(() => {
    menu.current?.querySelector(`[data-index="${active}"]`)?.scrollIntoView({ block: "nearest" });
  }, [active, position]);
  return <>
    <button ref={trigger} type="button" aria-label={label} aria-haspopup="listbox"
      aria-expanded={Boolean(position)} aria-controls={position ? id : undefined}
      data-testid="xlsx-validation-dropdown" disabled={!items?.length}
      aria-keyshortcuts="Alt+ArrowDown"
      title={items?.length ? "Choose a value (Alt/Option+↓)" : "This list source requires Excel"}
      onPointerDown={event => event.stopPropagation()}
      onDoubleClick={event => event.stopPropagation()}
      onClick={event => { event.stopPropagation(); if (position) close(); else open(); }}
      onKeyDown={event => {
        if (["ArrowDown", "ArrowUp"].includes(event.key)) { event.preventDefault(); event.stopPropagation(); open(); }
        else if (!["Delete", "Backspace"].includes(event.key)) event.stopPropagation();
      }}
      style={{ position: "absolute", right: -20 * zoom, top: 1, bottom: 1, width: 20 * zoom,
        boxSizing: "border-box", padding: 0, margin: 0,
        border: `1px solid ${palette.strongBorder}`, borderRadius: 0,
        background: palette.surface, color: palette.mutedText, fontSize: 11 * zoom,
        display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", zIndex: 6 }}
    ><span aria-hidden="true">▾</span></button>
    {position && items && createPortal(<div ref={menu} id={id} role="listbox" aria-label={label}
      aria-activedescendant={`${id}-${active}`} tabIndex={-1} data-testid="xlsx-validation-list"
      onPointerDown={event => { event.preventDefault(); event.stopPropagation(); }}
      onClick={event => event.stopPropagation()}
      onDoubleClick={event => event.stopPropagation()}
      onBlur={event => { if (!event.currentTarget.contains(event.relatedTarget)) setPosition(null); }}
      onKeyDown={event => {
        event.stopPropagation();
        if (event.key === "Tab") { setPosition(null); trigger.current?.focus(); return; }
        event.preventDefault();
        if (event.key === "Escape") close();
        else if (event.key === "Enter" || event.key === " ") choose(active);
        else if (event.key === "ArrowDown") setActive(index => (index + 1) % items.length);
        else if (event.key === "ArrowUp") setActive(index => (index - 1 + items.length) % items.length);
        else if (event.key === "Home") setActive(0);
        else if (event.key === "End") setActive(items.length - 1);
        else if (event.key.length === 1 && !event.metaKey && !event.ctrlKey) {
          const index = items.findIndex((_, offset) => String(items[(active + offset + 1) % items.length]).toLowerCase().startsWith(event.key.toLowerCase()));
          if (index >= 0) setActive((active + index + 1) % items.length);
        }
      }}
      style={{ ...position, position: "fixed", zIndex: 1000, boxSizing: "border-box", overflowY: "auto",
        padding: 3, border: `1px solid ${palette.strongBorder}`, borderRadius: 3,
        background: palette.surface, color: palette.text, boxShadow: "0 4px 12px rgba(0,0,0,0.16)",
        fontFamily: "system-ui, sans-serif", fontSize, outline: "none" }}>
      {items.map((item, index) => <div key={index} id={`${id}-${index}`} role="option"
        aria-selected={String(item) === value} data-index={index} title={String(item) || "(blank)"}
        onPointerMove={() => setActive(index)} onClick={() => choose(index)}
        style={{ height: rowHeight, boxSizing: "border-box", display: "flex", alignItems: "center",
          padding: "0 7px", gap: 6, borderRadius: 1, cursor: "pointer",
          background: active === index ? palette.headerSurface : "transparent" }}>
        <span aria-hidden="true" style={{ width: 12 * zoom, flexShrink: 0 }}>{String(item) === value ? "✓" : ""}</span>
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{String(item) || "(blank)"}</span>
      </div>)}
    </div>, document.body)}
  </>;
});
