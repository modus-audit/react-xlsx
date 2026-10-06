const descriptions: Record<string, string> = {
  "#DIV/0!": "The formula divides by zero or an empty cell.",
  "#REF!": "The formula refers to a cell or range that is unavailable.",
  "#VALUE!": "A value has the wrong type for this formula.",
  "#NAME?": "A name or function could not be recognized.",
  "#N/A": "A value needed by this formula is unavailable.",
  "#NUM!": "The formula could not calculate a valid number.",
  "#NULL!": "The referenced ranges do not intersect.",
  "#SPILL!": "The formula cannot place its results in the required cells.",
  "#CALC!": "The formula could not be calculated.",
};

/** Preserve any calculator error code, including ones introduced by a future engine. */
export function formulaErrorTooltip(error: string, formula?: string | null) {
  return [error, Object.hasOwn(descriptions, error) ? descriptions[error] : "The calculator returned an error for this formula.",
    formula ? `=${formula.replace(/^=/, "")}` : null].filter(Boolean).join("\n");
}
