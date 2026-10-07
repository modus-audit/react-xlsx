const descriptions: Record<string, string> = {
  "#DIV/0!": "Divides by zero or an empty cell.",
  "#REF!": "The formula refers to a cell or range that is unavailable.",
  "#VALUE!": "A value has the wrong type for this formula.",
  "#NAME?": "A name or function could not be recognized.",
  "#N/A": "A value needed by this formula is unavailable.",
  "#NUM!": "The formula could not calculate a valid number.",
  "#NULL!": "The referenced ranges do not intersect.",
  "#SPILL!": "The formula cannot place its results in the required cells.",
  "#CALC!": "The formula could not be calculated.",
};

export function formulaErrorTooltip(error: string) {
  return Object.hasOwn(descriptions, error) ? descriptions[error] : "This formula could not be calculated.";
}
