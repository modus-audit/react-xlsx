/** Excel literals; an apostrophe explicitly preserves text. Formula entry is handled by the caller. */
export function coerceUserEnteredValue(value: string): unknown {
  const trimmed = value.trim();
  if (!trimmed) return "";
  if (trimmed.startsWith("'")) return trimmed.slice(1);
  if (/^(true|false)$/i.test(trimmed)) return trimmed.toLowerCase() === "true";
  const numeric = /^[+-]?(?:\d+\.?\d*|\.\d+)(%)?$/.exec(trimmed);
  if (numeric) {
    const parsed = Number(numeric[1] ? trimmed.slice(0, -1) : trimmed);
    if (Number.isFinite(parsed)) return numeric[1] ? parsed / 100 : parsed;
  }
  return value;
}

/** Existing percentage formats keep their precision; new percent entries get a matching format. */
export function enteredPercentageFormat(value: string, style?: unknown): string | null {
  const numberFormat = typeof style === "object" && style !== null && "numberFormat" in style ? style.numberFormat : null;
  const formatString = typeof numberFormat === "object" && numberFormat !== null && "formatString" in numberFormat && typeof numberFormat.formatString === "string" ? numberFormat.formatString : undefined;
  const match = /^[+-]?(?:\d+(?:\.(\d*))?|\.(\d+))%$/.exec(value.trim());
  if (!match || formatString?.includes("%")) return null;
  const digits = Math.min((match[1] ?? match[2] ?? "").length, 15);
  return `0${digits ? `.${"0".repeat(digits)}` : ""}%`;
}
