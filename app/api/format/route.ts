import JSZip from "jszip";
import { XMLBuilder, XMLParser } from "fast-xml-parser";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

function ensureArray<T>(val: T | T[] | undefined): T[] {
  if (val === undefined) return [];
  return Array.isArray(val) ? val : [val];
}

function setDocDefaultsToTimesNewRoman(stylesRoot: any) {
  if (!stylesRoot["w:docDefaults"]) {
    stylesRoot["w:docDefaults"] = {};
  }
  const docDefaults = stylesRoot["w:docDefaults"]; // object

  if (!docDefaults["w:rPrDefault"]) docDefaults["w:rPrDefault"] = {};
  if (!docDefaults["w:rPrDefault"]["w:rPr"]) docDefaults["w:rPrDefault"]["w:rPr"] = {};

  const rPr = docDefaults["w:rPrDefault"]["w:rPr"];
  if (!rPr["w:rFonts"]) rPr["w:rFonts"] = {};
  rPr["w:rFonts"]["@_w:ascii"] = "Times New Roman";
  rPr["w:rFonts"]["@_w:hAnsi"] = "Times New Roman";
  rPr["w:rFonts"]["@_w:cs"] = "Times New Roman";

  // 24 half-points = 12pt
  rPr["w:sz"] = { "@_w:val": 24 };
  rPr["w:szCs"] = { "@_w:val": 24 };

  if (!docDefaults["w:pPrDefault"]) docDefaults["w:pPrDefault"] = {};
  if (!docDefaults["w:pPrDefault"]["w:pPr"]) docDefaults["w:pPrDefault"]["w:pPr"] = {};
  const pPr = docDefaults["w:pPrDefault"]["w:pPr"];
  // spacing: line 276 (=1.15), before/after 0
  pPr["w:spacing"] = { "@_w:line": 276, "@_w:lineRule": "auto", "@_w:before": 0, "@_w:after": 0 };
}

function normalizeStyle(style: any, target: { size?: number; bold?: boolean }) {
  if (!style["w:rPr"]) style["w:rPr"] = {};
  const rPr = style["w:rPr"];
  if (!rPr["w:rFonts"]) rPr["w:rFonts"] = {};
  rPr["w:rFonts"]["@_w:ascii"] = "Times New Roman";
  rPr["w:rFonts"]["@_w:hAnsi"] = "Times New Roman";
  rPr["w:rFonts"]["@_w:cs"] = "Times New Roman";
  if (target.size) {
    rPr["w:sz"] = { "@_w:val": target.size };
    rPr["w:szCs"] = { "@_w:val": target.size };
  }
  if (target.bold) rPr["w:b"] = {};
}

function editStylesXml(stylesXml: string): string {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
  const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: "@_", suppressEmptyNode: true });
  const stylesDoc = parser.parse(stylesXml);

  const stylesRoot = stylesDoc["w:styles"] ?? stylesDoc.styles ?? stylesDoc;

  setDocDefaultsToTimesNewRoman(stylesRoot);

  const styles = ensureArray(stylesRoot["w:style"]);
  for (const s of styles) {
    const styleId = s["@_w:styleId"];
    const type = s["@_w:w:type"] || s["@_w:type"]; // tolerate variants
    if (!styleId) continue;

    if (styleId === "Normal") {
      normalizeStyle(s, { size: 24 });
      if (!s["w:pPr"]) s["w:pPr"] = {};
      s["w:pPr"]["w:spacing"] = { "@_w:line": 276, "@_w:lineRule": "auto", "@_w:before": 0, "@_w:after": 0 };
    }
    if (type === "paragraph" || true) {
      if (styleId === "Heading1") normalizeStyle(s, { size: 32, bold: true });
      if (styleId === "Heading2") normalizeStyle(s, { size: 28, bold: true });
      if (styleId === "Heading3") normalizeStyle(s, { size: 24, bold: true });
    }
  }

  return builder.build(stylesDoc);
}

function editDocumentXml(documentXml: string): string {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
  const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: "@_", suppressEmptyNode: true });
  const doc = parser.parse(documentXml);

  const wDoc = doc["w:document"] ?? doc.document ?? doc;
  if (!wDoc["w:body"]) wDoc["w:body"] = {};
  const body = wDoc["w:body"];

  // Ensure a section properties block exists with 1 inch margins (1440 twips)
  if (!body["w:sectPr"]) body["w:sectPr"] = {};
  const sect = body["w:sectPr"];
  sect["w:pgMar"] = {
    "@_w:top": 1440,
    "@_w:right": 1440,
    "@_w:bottom": 1440,
    "@_w:left": 1440,
    "@_w:header": 720,
    "@_w:footer": 720,
    "@_w:gutter": 0
  };

  return builder.build(doc);
}

async function formatDocx(input: ArrayBuffer): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(input);

  // Modify styles.xml
  const stylesPath = "word/styles.xml";
  if (zip.file(stylesPath)) {
    const stylesXml = await zip.file(stylesPath)!.async("string");
    const newStyles = editStylesXml(stylesXml);
    zip.file(stylesPath, newStyles);
  }

  // Modify document.xml (for margins/sectPr)
  const docPath = "word/document.xml";
  if (zip.file(docPath)) {
    const docXml = await zip.file(docPath)!.async("string");
    const newDoc = editDocumentXml(docXml);
    zip.file(docPath, newDoc);
  }

  const output = await zip.generateAsync({ type: "uint8array" });
  return output;
}

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const file = form.get("file");
    if (!(file instanceof Blob)) {
      return new Response("No file provided", { status: 400 });
    }
    // Basic guard: require DOCX signature by extension only (content sniff would be heavy)
    const arrBuf = await file.arrayBuffer();
    const formatted = await formatDocx(arrBuf);

    return new Response(Buffer.from(formatted), {
      headers: new Headers({
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="formatted.docx"`
      })
    });
  } catch (err: any) {
    console.error(err);
    return new Response(err?.message || "Internal Error", { status: 500 });
  }
}
