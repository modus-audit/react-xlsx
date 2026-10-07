import { strFromU8, unzipSync } from "fflate";

const spreadsheetNamespaces = new Set([
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://purl.oclc.org/ooxml/spreadsheetml/main"
]);
const relationshipNamespaces = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships"
]);
const packageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
type Element = { name: string; namespace: string; namespaces: Record<string, string>; attributes: Record<string, string> };

function decodeAttribute(value: string): string {
  const entities: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
  return value.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, entity: string) =>
    entity.startsWith("#") ? String.fromCodePoint(entity[1] === "x"
      ? parseInt(entity.slice(2), 16) : Number(entity.slice(1))) : entities[entity]!);
}

// DOM-free for workers; track namespaces and paths to exclude extension formulas.
function scanXml(xml: string, visit: (path: Element[]) => void): void {
  const stack: Element[] = [];
  const tags = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<[^>"']*(?:"[^"]*"[^>"']*|'[^']*'[^>"']*)*>/g;
  for (const match of xml.matchAll(tags)) {
    const tag = match[0];
    if (tag.startsWith("<!") || tag.startsWith("<?")) continue;
    if (tag.startsWith("</")) { stack.pop(); continue; }
    const qualifiedName = /^<([^\s/>]+)/.exec(tag)?.[1];
    if (!qualifiedName) continue;
    const attributes = Object.fromEntries(Array.from(tag.matchAll(/\s([^\s=]+)\s*=\s*(["'])(.*?)\2/g),
      attribute => [attribute[1]!, decodeAttribute(attribute[3]!)]));
    const namespaces = Object.create(stack[stack.length - 1]?.namespaces ?? null) as Record<string, string>;
    for (const [name, value] of Object.entries(attributes)) {
      if (name === "xmlns") namespaces[""] = value;
      else if (name.startsWith("xmlns:")) namespaces[name.slice(6)] = value;
    }
    const separator = qualifiedName.indexOf(":");
    const name = separator < 0 ? qualifiedName : qualifiedName.slice(separator + 1);
    const prefix = separator < 0 ? "" : qualifiedName.slice(0, separator);
    const element = { name, namespace: namespaces[prefix] ?? "", namespaces, attributes };
    stack.push(element);
    visit(stack);
    if (/\/\s*>$/.test(tag)) stack.pop();
  }
}

function spreadsheetPath(path: Element[], names: string[]): boolean {
  return path.length === names.length && path.every((element, index) =>
    element.name === names[index] && spreadsheetNamespaces.has(element.namespace));
}

/** Count only worksheets linked by the workbook, never orphaned ZIP parts. */
export function countSourceWorkbookFormulas(bytes: Uint8Array): number | null {
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return null;
  try {
    const metadata = unzipSync(bytes, { filter: entry =>
      entry.name === "xl/workbook.xml" || entry.name === "xl/_rels/workbook.xml.rels" });
    if (!metadata["xl/workbook.xml"] || !metadata["xl/_rels/workbook.xml.rels"]) return null;
    const sheetIds: string[] = [];
    scanXml(strFromU8(metadata["xl/workbook.xml"]!), path => {
      if (!spreadsheetPath(path, ["workbook", "sheets", "sheet"])) return;
      const element = path[path.length - 1]!;
      const id = Object.entries(element.attributes).find(([name]) => {
        const [prefix, local] = name.split(":");
        return local === "id" && relationshipNamespaces.has(element.namespaces[prefix!] ?? "");
      })?.[1];
      if (!id) throw new Error("Missing worksheet relationship");
      sheetIds.push(id);
    });
    const relationships = new Map<string, Record<string, string>>();
    scanXml(strFromU8(metadata["xl/_rels/workbook.xml.rels"]!), path => {
      if (path.length === 2 && path[0]?.name === "Relationships" && path[1]?.name === "Relationship"
        && path.every(element => element.namespace === packageRelationshipNamespace)) {
        relationships.set(path[1]!.attributes.Id!, path[1]!.attributes);
      }
    });
    const worksheetPaths = new Set<string>();
    for (const id of sheetIds) {
      const relationship = relationships.get(id);
      if (!relationship?.Type || !relationship.Target || relationship.TargetMode === "External") return null;
      if (!relationship.Type.endsWith("/worksheet")) continue;
      const target = new URL(relationship.Target, "https://xlsx.invalid/xl/workbook.xml");
      if (target.origin !== "https://xlsx.invalid") return null;
      worksheetPaths.add(decodeURIComponent(target.pathname.slice(1)));
    }
    const worksheets = unzipSync(bytes, { filter: entry => worksheetPaths.has(entry.name) });
    let count = 0;
    for (const path of worksheetPaths) {
      if (!worksheets[path]) return null;
      scanXml(strFromU8(worksheets[path]!), elements => {
        if (spreadsheetPath(elements, ["worksheet", "sheetData", "row", "c", "f"])) count += 1;
      });
    }
    return count;
  } catch {
    return null;
  }
}
