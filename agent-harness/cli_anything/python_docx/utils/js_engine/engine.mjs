import fs from "node:fs/promises";
import path from "node:path";

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import JSZip from "jszip";
import xpath from "xpath";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS = "http://purl.org/dc/terms/";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";

const FOOTNOTES_RELTYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
const FOOTNOTES_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";

const FOOTNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="${W_NS}">
  <w:footnote w:type="separator" w:id="-1">
    <w:p>
      <w:r>
        <w:separator/>
      </w:r>
    </w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p>
      <w:r>
        <w:continuationSeparator/>
      </w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
`;

const BIBLIOGRAPHY_HEADING = "Bibliography";
const BIBLIOGRAPHY_ENTRY_PATTERN = /^\[(\d+)\]\s+([A-Za-z0-9._:-]+)\s*:\s*(.+)$/;
const BIBLIOGRAPHY_KEY_PATTERN = /^[A-Za-z0-9._:-]+$/;
const PALETTE_MAIN = "05206E";
const PALETTE_SECONDARY = "357AE9";
const PALETTE_WHITE = "FFFFFF";
const FRONTPAGE_TEMPLATES = ["clean", "corporate", "academic"];

const CORE_PROPERTY_MAP = {
  author: { namespace: DC_NS, prefix: "dc", element: "creator", type: "string" },
  category: { namespace: CP_NS, prefix: "cp", element: "category", type: "string" },
  comments: { namespace: DC_NS, prefix: "dc", element: "description", type: "string" },
  content_status: { namespace: CP_NS, prefix: "cp", element: "contentStatus", type: "string" },
  created: { namespace: DCTERMS_NS, prefix: "dcterms", element: "created", type: "datetime" },
  identifier: { namespace: DC_NS, prefix: "dc", element: "identifier", type: "string" },
  keywords: { namespace: CP_NS, prefix: "cp", element: "keywords", type: "string" },
  language: { namespace: DC_NS, prefix: "dc", element: "language", type: "string" },
  last_modified_by: { namespace: CP_NS, prefix: "cp", element: "lastModifiedBy", type: "string" },
  last_printed: { namespace: CP_NS, prefix: "cp", element: "lastPrinted", type: "datetime" },
  modified: { namespace: DCTERMS_NS, prefix: "dcterms", element: "modified", type: "datetime" },
  revision: { namespace: CP_NS, prefix: "cp", element: "revision", type: "int" },
  subject: { namespace: DC_NS, prefix: "dc", element: "subject", type: "string" },
  title: { namespace: DC_NS, prefix: "dc", element: "title", type: "string" },
  version: { namespace: CP_NS, prefix: "cp", element: "version", type: "string" },
};

function fail(message) {
  throw new Error(message);
}

function parseXml(xml) {
  return new DOMParser().parseFromString(xml, "application/xml");
}

function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function makeSelector(namespaces) {
  return xpath.useNamespaces(namespaces);
}

function getAttribute(node, primary, fallback = null) {
  const direct = node.getAttribute(primary);
  if (direct !== null && direct !== undefined) return direct;
  if (fallback) return node.getAttribute(fallback);
  return null;
}

function setWAttr(node, name, value) {
  node.setAttributeNS(W_NS, `w:${name}`, String(value));
}

function paragraphText(selectW, paragraphNode) {
  const texts = selectW(".//w:t", paragraphNode);
  return texts.map((node) => node.textContent ?? "").join("");
}

function paragraphStyle(selectW, paragraphNode) {
  const attrs = selectW("w:pPr/w:pStyle/@w:val | w:pPr/w:pStyle/@val", paragraphNode);
  if (!attrs.length) return null;
  return attrs[0].value ?? null;
}

function createRunWithText(doc, text) {
  const run = doc.createElementNS(W_NS, "w:r");
  const t = doc.createElementNS(W_NS, "w:t");
  if (/^\s|\s$/.test(text)) {
    t.setAttributeNS(XML_NS, "xml:space", "preserve");
  }
  t.appendChild(doc.createTextNode(text));
  run.appendChild(t);
  return run;
}

function createFootnoteReferenceRun(doc, footnoteId) {
  const run = doc.createElementNS(W_NS, "w:r");
  const ref = doc.createElementNS(W_NS, "w:footnoteReference");
  setWAttr(ref, "id", footnoteId);
  run.appendChild(ref);
  return run;
}

function createParagraph(doc, text, options = {}) {
  const paragraph = doc.createElementNS(W_NS, "w:p");
  const styleName = options.styleName ?? null;
  if (styleName) {
    const pPr = doc.createElementNS(W_NS, "w:pPr");
    const pStyle = doc.createElementNS(W_NS, "w:pStyle");
    setWAttr(pStyle, "val", styleName);
    pPr.appendChild(pStyle);
    paragraph.appendChild(pPr);
  }

  const run = createRunWithText(doc, text);
  if (options.bold || options.colorHex) {
    const rPr = doc.createElementNS(W_NS, "w:rPr");
    if (options.bold) {
      const bold = doc.createElementNS(W_NS, "w:b");
      rPr.appendChild(bold);
    }
    if (options.colorHex) {
      const color = doc.createElementNS(W_NS, "w:color");
      setWAttr(color, "val", options.colorHex);
      rPr.appendChild(color);
    }
    run.insertBefore(rPr, run.firstChild);
  }
  paragraph.appendChild(run);
  return paragraph;
}

function createFrontpageParagraph(doc, text, options = {}) {
  const paragraph = doc.createElementNS(W_NS, "w:p");
  const pPr = doc.createElementNS(W_NS, "w:pPr");
  const jc = doc.createElementNS(W_NS, "w:jc");
  setWAttr(jc, "val", "center");
  pPr.appendChild(jc);
  if (options.spaceAfter && options.spaceAfter > 0) {
    const spacing = doc.createElementNS(W_NS, "w:spacing");
    setWAttr(spacing, "after", options.spaceAfter * 20);
    pPr.appendChild(spacing);
  }
  if (options.backgroundColor) {
    const shd = doc.createElementNS(W_NS, "w:shd");
    setWAttr(shd, "val", "clear");
    setWAttr(shd, "color", "auto");
    setWAttr(shd, "fill", options.backgroundColor);
    pPr.appendChild(shd);
  }
  paragraph.appendChild(pPr);

  const run = doc.createElementNS(W_NS, "w:r");
  const rPr = doc.createElementNS(W_NS, "w:rPr");
  if (options.bold) {
    rPr.appendChild(doc.createElementNS(W_NS, "w:b"));
  }
  if (options.italic) {
    rPr.appendChild(doc.createElementNS(W_NS, "w:i"));
  }
  if (options.size) {
    const sz = doc.createElementNS(W_NS, "w:sz");
    setWAttr(sz, "val", options.size * 2);
    rPr.appendChild(sz);
  }
  const color = doc.createElementNS(W_NS, "w:color");
  setWAttr(color, "val", options.colorHex ?? PALETTE_MAIN);
  rPr.appendChild(color);
  run.appendChild(rPr);

  const t = doc.createElementNS(W_NS, "w:t");
  t.appendChild(doc.createTextNode(text));
  run.appendChild(t);
  paragraph.appendChild(run);
  return paragraph;
}

function createPageBreakParagraph(doc) {
  const paragraph = doc.createElementNS(W_NS, "w:p");
  const pPr = doc.createElementNS(W_NS, "w:pPr");
  const jc = doc.createElementNS(W_NS, "w:jc");
  setWAttr(jc, "val", "center");
  pPr.appendChild(jc);
  paragraph.appendChild(pPr);

  const run = doc.createElementNS(W_NS, "w:r");
  const br = doc.createElementNS(W_NS, "w:br");
  setWAttr(br, "type", "page");
  run.appendChild(br);
  paragraph.appendChild(run);
  return paragraph;
}

function appendParagraphToBody(selectW, documentDoc, paragraphNode) {
  const body = selectW("/w:document/w:body", documentDoc)[0];
  if (!body) fail("Document body missing.");
  const sectPr = selectW("w:sectPr", body)[0];
  if (sectPr) {
    body.insertBefore(paragraphNode, sectPr);
    return;
  }
  body.appendChild(paragraphNode);
}

function insertParagraphBefore(paragraphNode, newParagraphNode) {
  const parent = paragraphNode.parentNode;
  if (!parent) fail("Paragraph has no parent.");
  parent.insertBefore(newParagraphNode, paragraphNode);
}

function appendBodyChild(selectW, documentDoc, node) {
  const body = selectW("/w:document/w:body", documentDoc)[0];
  if (!body) fail("Document body missing.");
  const sectPr = selectW("w:sectPr", body)[0];
  if (sectPr) {
    body.insertBefore(node, sectPr);
    return;
  }
  body.appendChild(node);
}

function headingStyleFromLevel(level) {
  if (level === 0) return "Title";
  return `Heading${level}`;
}

function normalizePaletteColor(value, fieldName) {
  const normalized = String(value ?? "").trim().toLowerCase().replace(/^#/, "");
  if (normalized === "main" || normalized === PALETTE_MAIN.toLowerCase()) return PALETTE_MAIN;
  if (normalized === "secondary" || normalized === PALETTE_SECONDARY.toLowerCase()) return PALETTE_SECONDARY;
  fail(`${fieldName} must be one of: main, secondary, ${PALETTE_MAIN}, ${PALETTE_SECONDARY}.`);
}

function ensureNamespaceDeclaration(root, prefix, namespace) {
  const attrName = `xmlns:${prefix}`;
  if (!root.getAttribute(attrName)) {
    root.setAttribute(attrName, namespace);
  }
}

function ensureCoreElement(selectCore, coreDoc, root, descriptor) {
  const nodes = selectCore(`${descriptor.prefix}:${descriptor.element}`, root);
  if (nodes.length) return nodes[0];
  const element = coreDoc.createElementNS(
    descriptor.namespace,
    `${descriptor.prefix}:${descriptor.element}`,
  );
  root.appendChild(element);
  return element;
}

function normalizedDatetimeValue(value, key) {
  const datetime = new Date(String(value));
  if (Number.isNaN(datetime.getTime())) {
    fail(`Invalid datetime for core property '${key}': ${value}`);
  }
  return datetime.toISOString().replace(/\.\d{3}Z$/, "Z");
}

function extractBibliography(selectW, documentDoc) {
  const paragraphs = selectW("/w:document/w:body/w:p", documentDoc);
  let headingIndex = -1;
  for (let index = 0; index < paragraphs.length; index += 1) {
    if (paragraphText(selectW, paragraphs[index]).trim().toLowerCase() === BIBLIOGRAPHY_HEADING.toLowerCase()) {
      headingIndex = index;
      break;
    }
  }

  if (headingIndex === -1) {
    return {
      paragraphs,
      headingIndex: -1,
      nextHeadingIndex: -1,
      entries: [],
    };
  }

  let nextHeadingIndex = -1;
  for (let index = headingIndex + 1; index < paragraphs.length; index += 1) {
    const styleName = paragraphStyle(selectW, paragraphs[index]);
    if (styleName && styleName.startsWith("Heading")) {
      nextHeadingIndex = index;
      break;
    }
  }

  const endIndex = nextHeadingIndex === -1 ? paragraphs.length : nextHeadingIndex;
  const entries = [];
  for (let index = headingIndex + 1; index < endIndex; index += 1) {
    const text = paragraphText(selectW, paragraphs[index]).trim();
    if (!text) continue;
    const match = text.match(BIBLIOGRAPHY_ENTRY_PATTERN);
    if (!match) continue;
    entries.push({
      number: Number.parseInt(match[1], 10),
      key: match[2],
      reference: match[3],
      paragraph: paragraphs[index],
    });
  }
  entries.sort((a, b) => a.number - b.number);

  return {
    paragraphs,
    headingIndex,
    nextHeadingIndex,
    entries,
  };
}

function normalizeBibliographyKey(key) {
  const normalized = String(key ?? "").trim();
  if (!normalized) fail("Bibliography key cannot be empty.");
  if (!BIBLIOGRAPHY_KEY_PATTERN.test(normalized)) {
    fail("Bibliography key must only contain letters, numbers, dot, underscore, colon, or dash.");
  }
  return normalized;
}

async function loadZip(docPath) {
  const fileBuffer = await fs.readFile(docPath);
  return JSZip.loadAsync(fileBuffer);
}

async function getXml(zip, zipPath) {
  const file = zip.file(zipPath);
  if (!file) return null;
  return file.async("string");
}

async function getRequiredXml(zip, zipPath) {
  const xml = await getXml(zip, zipPath);
  if (!xml) fail(`Missing required XML part: ${zipPath}`);
  return xml;
}

function findNextRelationshipId(selectRels, relsDoc) {
  const rels = selectRels("/rels:Relationships/rels:Relationship", relsDoc);
  let highest = 0;
  rels.forEach((rel) => {
    const relId = getAttribute(rel, "Id");
    const match = relId ? relId.match(/^rId(\d+)$/) : null;
    if (match) {
      highest = Math.max(highest, Number.parseInt(match[1], 10));
    }
  });
  return `rId${highest + 1}`;
}

function ensureContentTypeOverride(selectCT, contentTypesDoc) {
  const existing = selectCT(
    `/ct:Types/ct:Override[@PartName='/word/footnotes.xml' and @ContentType='${FOOTNOTES_CONTENT_TYPE}']`,
    contentTypesDoc,
  );
  if (existing.length) return false;
  const types = selectCT("/ct:Types", contentTypesDoc)[0];
  if (!types) fail("[Content_Types].xml missing <Types> root.");
  const override = contentTypesDoc.createElementNS(CONTENT_TYPES_NS, "Override");
  override.setAttribute("PartName", "/word/footnotes.xml");
  override.setAttribute("ContentType", FOOTNOTES_CONTENT_TYPE);
  types.appendChild(override);
  return true;
}

function ensureFootnotesRelationship(selectRels, relsDoc) {
  const existing = selectRels(
    `/rels:Relationships/rels:Relationship[@Type='${FOOTNOTES_RELTYPE}']`,
    relsDoc,
  );
  if (existing.length) return false;
  const root = selectRels("/rels:Relationships", relsDoc)[0];
  if (!root) fail("document.xml.rels missing <Relationships> root.");
  const relationship = relsDoc.createElementNS(REL_NS, "Relationship");
  relationship.setAttribute("Id", findNextRelationshipId(selectRels, relsDoc));
  relationship.setAttribute("Type", FOOTNOTES_RELTYPE);
  relationship.setAttribute("Target", "footnotes.xml");
  root.appendChild(relationship);
  return true;
}

function ensureFootnotesPart(selectW, footnotesDoc) {
  const root = selectW("/w:footnotes", footnotesDoc)[0];
  if (!root) fail("footnotes.xml missing <w:footnotes> root.");
}

function createFootnoteBodyParagraph(doc, text) {
  const paragraph = doc.createElementNS(W_NS, "w:p");
  paragraph.appendChild(createRunWithText(doc, text));
  return paragraph;
}

function upsertFootnote(selectW, footnotesDoc, footnoteId, text) {
  const byId = selectW(`/w:footnotes/w:footnote[@w:id='${footnoteId}']`, footnotesDoc);
  let footnoteNode = null;
  byId.forEach((candidate) => {
    const footnoteType = getAttribute(candidate, "w:type", "type");
    if (!footnoteType) {
      footnoteNode = candidate;
    }
  });

  if (!footnoteNode) {
    const root = selectW("/w:footnotes", footnotesDoc)[0];
    if (!root) fail("footnotes.xml missing <w:footnotes> root.");
    footnoteNode = footnotesDoc.createElementNS(W_NS, "w:footnote");
    setWAttr(footnoteNode, "id", footnoteId);
    root.appendChild(footnoteNode);
  } else {
    while (footnoteNode.firstChild) {
      footnoteNode.removeChild(footnoteNode.firstChild);
    }
  }

  footnoteNode.appendChild(createFootnoteBodyParagraph(footnotesDoc, text));
}

function citationMarker(citationNumbers) {
  if (citationNumbers.length === 1) {
    return `[${citationNumbers[0]}]`;
  }
  return `[${citationNumbers.join(", ")}]`;
}

function resolveCitationEntries(entries, sourceKeys) {
  if (!sourceKeys.length) fail("At least one source key is required.");
  const keyToEntry = new Map(entries.map((entry) => [entry.key.toLowerCase(), entry]));
  const resolved = [];
  const missing = [];
  sourceKeys.forEach((key) => {
    const normalized = normalizeBibliographyKey(key);
    const match = keyToEntry.get(normalized.toLowerCase());
    if (!match) {
      missing.push(normalized);
      return;
    }
    if (!resolved.find((entry) => entry.number === match.number)) {
      resolved.push(match);
    }
  });
  if (missing.length) {
    fail(`Unknown bibliography key(s): ${missing.join(", ")}`);
  }
  return resolved;
}

function appendCitationRuns(documentDoc, paragraphNode, citationEntries) {
  paragraphNode.appendChild(createRunWithText(documentDoc, " "));
  citationEntries.forEach((entry) => {
    paragraphNode.appendChild(createFootnoteReferenceRun(documentDoc, entry.number));
  });
}

async function loadState(docPath) {
  const absoluteDocPath = path.resolve(docPath);
  const zip = await loadZip(absoluteDocPath);

  const documentXml = await getRequiredXml(zip, "word/document.xml");
  const relsXml = await getRequiredXml(zip, "word/_rels/document.xml.rels");
  const contentTypesXml = await getRequiredXml(zip, "[Content_Types].xml");
  const existingFootnotesXml = await getXml(zip, "word/footnotes.xml");

  const documentDoc = parseXml(documentXml);
  const relsDoc = parseXml(relsXml);
  const contentTypesDoc = parseXml(contentTypesXml);
  const footnotesDoc = parseXml(existingFootnotesXml ?? FOOTNOTES_XML);

  const selectW = makeSelector({ w: W_NS });
  const selectRels = makeSelector({ rels: REL_NS });
  const selectCT = makeSelector({ ct: CONTENT_TYPES_NS });

  ensureFootnotesPart(selectW, footnotesDoc);

  return {
    absoluteDocPath,
    zip,
    documentDoc,
    relsDoc,
    contentTypesDoc,
    footnotesDoc,
    selectW,
    selectRels,
    selectCT,
    footnotesFileExisted: Boolean(existingFootnotesXml),
  };
}

async function saveState(state) {
  state.zip.file("word/document.xml", serializeXml(state.documentDoc));
  state.zip.file("word/_rels/document.xml.rels", serializeXml(state.relsDoc));
  state.zip.file("[Content_Types].xml", serializeXml(state.contentTypesDoc));
  state.zip.file("word/footnotes.xml", serializeXml(state.footnotesDoc));

  const outputBuffer = await state.zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(state.absoluteDocPath, outputBuffer);
}

async function loadCoreState(docPath) {
  const absoluteDocPath = path.resolve(docPath);
  const zip = await loadZip(absoluteDocPath);
  const coreXml = await getRequiredXml(zip, "docProps/core.xml");
  const coreDoc = parseXml(coreXml);
  const selectCore = makeSelector({
    cp: CP_NS,
    dc: DC_NS,
    dcterms: DCTERMS_NS,
    xsi: XSI_NS,
  });
  const root = selectCore("/cp:coreProperties", coreDoc)[0];
  if (!root) fail("docProps/core.xml missing <cp:coreProperties> root.");
  return {
    absoluteDocPath,
    zip,
    coreDoc,
    selectCore,
    root,
  };
}

async function saveCoreState(state) {
  state.zip.file("docProps/core.xml", serializeXml(state.coreDoc));
  const outputBuffer = await state.zip.generateAsync({ type: "nodebuffer" });
  await fs.writeFile(state.absoluteDocPath, outputBuffer);
}

function ensureBibliographySection(state) {
  const before = extractBibliography(state.selectW, state.documentDoc);
  if (before.headingIndex !== -1) return before;

  const headingParagraph = createParagraph(state.documentDoc, BIBLIOGRAPHY_HEADING, {
    styleName: "Heading1",
    colorHex: "05206E",
  });
  appendParagraphToBody(state.selectW, state.documentDoc, headingParagraph);

  return extractBibliography(state.selectW, state.documentDoc);
}

function bibliographyFootnoteText(entry) {
  return `[${entry.number}] ${entry.key}: ${entry.reference}`;
}

async function addBibliographyEntry(request) {
  const state = await loadState(request.docPath);
  const key = normalizeBibliographyKey(request.key);
  const reference = String(request.reference ?? "").trim();
  if (!reference) fail("Bibliography reference cannot be empty.");
  const url = request.url ? String(request.url).trim() : null;

  const bib = ensureBibliographySection(state);
  const existingKey = bib.entries.find((entry) => entry.key.toLowerCase() === key.toLowerCase());
  if (existingKey) fail(`Bibliography key already exists: ${key}`);

  const citationNumber = (bib.entries[bib.entries.length - 1]?.number ?? 0) + 1;
  let entryText = `[${citationNumber}] ${key}: ${reference}`;
  if (url) {
    entryText = `${entryText} (URL: ${url})`;
  }

  const paragraph = createParagraph(state.documentDoc, entryText);
  if (bib.nextHeadingIndex !== -1) {
    const targetParagraph = bib.paragraphs[bib.nextHeadingIndex];
    insertParagraphBefore(targetParagraph, paragraph);
  } else {
    appendParagraphToBody(state.selectW, state.documentDoc, paragraph);
  }

  ensureFootnotesRelationship(state.selectRels, state.relsDoc);
  ensureContentTypeOverride(state.selectCT, state.contentTypesDoc);
  upsertFootnote(
    state.selectW,
    state.footnotesDoc,
    citationNumber,
    bibliographyFootnoteText({
      number: citationNumber,
      key,
      reference: url ? `${reference} (URL: ${url})` : reference,
    }),
  );
  await saveState(state);

  return {
    entry: {
      number: citationNumber,
      key,
      reference,
      url,
      text: entryText,
    },
  };
}

async function summary(request) {
  const state = await loadState(request.docPath);
  const paragraphs = state.selectW("/w:document/w:body/w:p", state.documentDoc);
  const tables = state.selectW("/w:document/w:body/w:tbl", state.documentDoc);

  let headingCount = 0;
  paragraphs.forEach((paragraph) => {
    const styleName = paragraphStyle(state.selectW, paragraph);
    if (styleName && styleName.startsWith("Heading")) {
      headingCount += 1;
    }
  });

  return {
    summary: {
      path: state.absoluteDocPath,
      paragraph_count: paragraphs.length,
      table_count: tables.length,
      heading_count: headingCount,
      undo_depth: 0,
      redo_depth: 0,
    },
  };
}

async function listParagraphs(request) {
  const state = await loadState(request.docPath);
  const paragraphs = state.selectW("/w:document/w:body/w:p", state.documentDoc);
  const limit = request.limit ?? null;

  const rows = paragraphs.map((paragraph, index) => {
    let normalizedStyle = paragraphStyle(state.selectW, paragraph);
    if (normalizedStyle && /^Heading\d$/.test(normalizedStyle)) {
      normalizedStyle = normalizedStyle.replace("Heading", "Heading ");
    }
    return {
      index,
      text: paragraphText(state.selectW, paragraph),
      style: normalizedStyle,
    };
  });

  return {
    paragraphs: limit === null ? rows : rows.slice(0, Math.max(limit, 0)),
  };
}

async function addParagraph(request) {
  const state = await loadState(request.docPath);
  const text = String(request.text ?? "");
  const style = request.style ? String(request.style) : null;
  const paragraph = createParagraph(state.documentDoc, text, {
    styleName: style,
  });
  appendParagraphToBody(state.selectW, state.documentDoc, paragraph);
  await saveState(state);

  return {
    paragraph: {
      text,
      style,
    },
  };
}

async function addHeading(request) {
  const state = await loadState(request.docPath);
  const text = String(request.text ?? "");
  const level = Number.parseInt(String(request.level ?? 1), 10);
  if (Number.isNaN(level) || level < 0 || level > 9) {
    fail(`level must be in range 0-9, got ${request.level}`);
  }

  const paragraph = createParagraph(state.documentDoc, text, {
    styleName: headingStyleFromLevel(level),
    colorHex: "05206E",
  });
  appendParagraphToBody(state.selectW, state.documentDoc, paragraph);
  await saveState(state);

  return {
    heading: {
      text,
      level,
      style: headingStyleFromLevel(level),
    },
  };
}

function createTableCell(doc, text, options = {}) {
  const tc = doc.createElementNS(W_NS, "w:tc");
  const tcPr = doc.createElementNS(W_NS, "w:tcPr");
  if (options.backgroundColor) {
    const shd = doc.createElementNS(W_NS, "w:shd");
    setWAttr(shd, "val", "clear");
    setWAttr(shd, "color", "auto");
    setWAttr(shd, "fill", options.backgroundColor);
    tcPr.appendChild(shd);
  }
  tc.appendChild(tcPr);
  tc.appendChild(
    createParagraph(doc, text, {
      bold: options.bold ?? false,
      colorHex: options.colorHex ?? null,
    }),
  );
  return tc;
}

function applyTableBorders(selectW, tableNode, options) {
  const tblPr = selectW("w:tblPr", tableNode)[0] ?? tableNode.insertBefore(tableNode.ownerDocument.createElementNS(W_NS, "w:tblPr"), tableNode.firstChild);
  let tblBorders = selectW("w:tblBorders", tblPr)[0];
  if (!tblBorders) {
    tblBorders = tableNode.ownerDocument.createElementNS(W_NS, "w:tblBorders");
    tblPr.appendChild(tblBorders);
  }

  const edges = [];
  if (options.outerBorder) edges.push("top", "left", "bottom", "right");
  if (options.rowLines) edges.push("insideH");
  if (options.columnLines) edges.push("insideV");

  edges.forEach((edge) => {
    let edgeNode = selectW(`w:${edge}`, tblBorders)[0];
    if (!edgeNode) {
      edgeNode = tableNode.ownerDocument.createElementNS(W_NS, `w:${edge}`);
      tblBorders.appendChild(edgeNode);
    }
    setWAttr(edgeNode, "val", options.lineStyle);
    setWAttr(edgeNode, "sz", options.lineSize);
    setWAttr(edgeNode, "space", 0);
    setWAttr(edgeNode, "color", options.lineColor);
  });
}

function createTableNode(doc, rowCount, colCount) {
  const tbl = doc.createElementNS(W_NS, "w:tbl");
  const tblPr = doc.createElementNS(W_NS, "w:tblPr");
  const tblW = doc.createElementNS(W_NS, "w:tblW");
  setWAttr(tblW, "w", 0);
  setWAttr(tblW, "type", "auto");
  tblPr.appendChild(tblW);
  tbl.appendChild(tblPr);

  const tblGrid = doc.createElementNS(W_NS, "w:tblGrid");
  for (let index = 0; index < colCount; index += 1) {
    const gridCol = doc.createElementNS(W_NS, "w:gridCol");
    setWAttr(gridCol, "w", 2880);
    tblGrid.appendChild(gridCol);
  }
  tbl.appendChild(tblGrid);

  for (let row = 0; row < rowCount; row += 1) {
    const tr = doc.createElementNS(W_NS, "w:tr");
    for (let col = 0; col < colCount; col += 1) {
      tr.appendChild(createTableCell(doc, ""));
    }
    tbl.appendChild(tr);
  }
  return tbl;
}

function tableRows(selectW, tableNode) {
  return selectW("w:tr", tableNode);
}

function tableCells(selectW, rowNode) {
  return selectW("w:tc", rowNode);
}

async function addTable(request) {
  const state = await loadState(request.docPath);
  const headers = Array.isArray(request.headers) ? request.headers : [];
  const records = Array.isArray(request.records) ? request.records : [];
  const rows = request.rows ?? null;
  const cols = request.cols ?? null;

  const hasStructuredData = headers.length > 0 || records.length > 0;
  if (!hasStructuredData && (rows === null || cols === null)) {
    fail("rows and cols are required when no headers/records are provided.");
  }
  if (cols !== null && cols < 1) fail("cols must be >= 1 when provided.");
  if (rows !== null && rows < 1) fail("rows must be >= 1 when provided.");

  let tableRowsCount = 0;
  let tableColsCount = 0;

  if (!hasStructuredData) {
    tableRowsCount = rows;
    tableColsCount = cols;
  } else {
    let inferredCols = cols ?? 0;
    inferredCols = Math.max(inferredCols, headers.length);
    records.forEach((rowValues) => {
      inferredCols = Math.max(inferredCols, Array.isArray(rowValues) ? rowValues.length : 0);
    });
    if (inferredCols < 1) fail("Unable to infer table column count from headers/records.");

    tableColsCount = inferredCols;
    tableRowsCount = (headers.length ? 1 : 0) + records.length;
    if (rows !== null) {
      tableRowsCount = Math.max(tableRowsCount, rows);
    }
  }

  const table = createTableNode(state.documentDoc, tableRowsCount, tableColsCount);
  const rowNodes = tableRows(state.selectW, table);

  let nextRow = 0;
  if (headers.length) {
    const headerBgColor = request.headerBgColor
      ? normalizePaletteColor(request.headerBgColor, "header_bg_color")
      : PALETTE_MAIN;
    const headerRowCells = tableCells(state.selectW, rowNodes[0]);
    headers.forEach((value, index) => {
      if (index >= headerRowCells.length) return;
      const cell = createTableCell(state.documentDoc, String(value), {
        bold: request.headerBold ?? false,
        colorHex: PALETTE_WHITE,
        backgroundColor: headerBgColor,
      });
      rowNodes[0].replaceChild(cell, headerRowCells[index]);
    });
    nextRow = 1;
  }

  records.forEach((recordValues) => {
    if (nextRow >= rowNodes.length) return;
    const currentCells = tableCells(state.selectW, rowNodes[nextRow]);
    (Array.isArray(recordValues) ? recordValues : []).forEach((value, colIndex) => {
      if (colIndex >= currentCells.length) return;
      const cell = createTableCell(state.documentDoc, String(value));
      rowNodes[nextRow].replaceChild(cell, currentCells[colIndex]);
    });
    nextRow += 1;
  });

  if (request.rowLines || request.columnLines || request.outerBorder) {
    applyTableBorders(state.selectW, table, {
      rowLines: Boolean(request.rowLines),
      columnLines: Boolean(request.columnLines),
      outerBorder: Boolean(request.outerBorder),
      lineStyle: String(request.lineStyle ?? "single"),
      lineSize: Number.parseInt(String(request.lineSize ?? 8), 10),
      lineColor: normalizePaletteColor(request.lineColor ?? "main", "line_color"),
    });
  }

  appendBodyChild(state.selectW, state.documentDoc, table);
  await saveState(state);

  return {
    table: {
      rows: tableRowsCount,
      cols: tableColsCount,
    },
  };
}

function bodyContentBlocks(selectW, documentDoc) {
  return selectW("/w:document/w:body/*[not(self::w:sectPr)]", documentDoc);
}

function prependBodyBlocks(selectW, documentDoc, blocks) {
  const body = selectW("/w:document/w:body", documentDoc)[0];
  if (!body) fail("Document body missing.");
  const existingBlocks = bodyContentBlocks(selectW, documentDoc);
  const firstBlock = existingBlocks.length ? existingBlocks[0] : null;
  if (!firstBlock) {
    blocks.forEach((block) => appendBodyChild(selectW, documentDoc, block));
    return;
  }
  blocks.forEach((block) => {
    body.insertBefore(block, firstBlock);
  });
}

async function addFrontpage(request) {
  const state = await loadState(request.docPath);
  const template = String(request.template ?? "clean").trim().toLowerCase();
  const title = String(request.title ?? "").trim();
  const subtitle = request.subtitle ? String(request.subtitle) : null;
  const author = request.author ? String(request.author) : null;
  const organization = request.organization ? String(request.organization) : null;
  const dateText = request.dateText ? String(request.dateText) : new Date().toISOString().slice(0, 10);
  const includePageBreak = request.includePageBreak !== false;
  const setCoreTitle = request.setCoreTitle !== false;

  if (!FRONTPAGE_TEMPLATES.includes(template)) {
    fail(`Unsupported frontpage template: ${template}. Available: ${FRONTPAGE_TEMPLATES.join(", ")}`);
  }
  if (!title) {
    fail("Frontpage title cannot be empty.");
  }

  const blocks = [];
  if (template === "clean") {
    blocks.push(createFrontpageParagraph(state.documentDoc, title, { size: 30, bold: true, spaceAfter: 18 }));
    if (subtitle) {
      blocks.push(
        createFrontpageParagraph(state.documentDoc, subtitle, {
          size: 16,
          italic: true,
          spaceAfter: 24,
          colorHex: PALETTE_SECONDARY,
        }),
      );
    }
    if (author) {
      blocks.push(createFrontpageParagraph(state.documentDoc, author, { size: 12, spaceAfter: 6 }));
    }
    if (organization) {
      blocks.push(createFrontpageParagraph(state.documentDoc, organization, { size: 12, spaceAfter: 6 }));
    }
    blocks.push(
      createFrontpageParagraph(state.documentDoc, dateText, {
        size: 11,
        colorHex: PALETTE_SECONDARY,
      }),
    );
  } else if (template === "corporate") {
    if (organization) {
      blocks.push(
        createFrontpageParagraph(state.documentDoc, organization.toUpperCase(), {
          size: 12,
          bold: true,
          spaceAfter: 24,
          colorHex: PALETTE_WHITE,
          backgroundColor: PALETTE_MAIN,
        }),
      );
    }
    blocks.push(
      createFrontpageParagraph(state.documentDoc, title, {
        size: 28,
        bold: true,
        spaceAfter: 12,
        colorHex: PALETTE_WHITE,
        backgroundColor: PALETTE_MAIN,
      }),
    );
    if (subtitle) {
      blocks.push(
        createFrontpageParagraph(state.documentDoc, subtitle, {
          size: 14,
          spaceAfter: 24,
          colorHex: PALETTE_WHITE,
          backgroundColor: PALETTE_MAIN,
        }),
      );
    }
    if (author) {
      blocks.push(
        createFrontpageParagraph(state.documentDoc, `Prepared by ${author}`, {
          size: 12,
          spaceAfter: 6,
          colorHex: PALETTE_WHITE,
          backgroundColor: PALETTE_MAIN,
        }),
      );
    }
    blocks.push(
      createFrontpageParagraph(state.documentDoc, dateText, {
        size: 11,
        colorHex: PALETTE_WHITE,
        backgroundColor: PALETTE_MAIN,
      }),
    );
  } else {
    blocks.push(createFrontpageParagraph(state.documentDoc, title, { size: 26, bold: true, spaceAfter: 12 }));
    if (subtitle) {
      blocks.push(
        createFrontpageParagraph(state.documentDoc, subtitle, {
          size: 14,
          italic: true,
          spaceAfter: 18,
          colorHex: PALETTE_SECONDARY,
        }),
      );
    }
    if (author) {
      blocks.push(createFrontpageParagraph(state.documentDoc, `Author: ${author}`, { size: 12, spaceAfter: 6 }));
    }
    if (organization) {
      blocks.push(
        createFrontpageParagraph(state.documentDoc, `Institution: ${organization}`, {
          size: 12,
          spaceAfter: 6,
        }),
      );
    }
    blocks.push(
      createFrontpageParagraph(state.documentDoc, dateText, {
        size: 11,
        colorHex: PALETTE_SECONDARY,
      }),
    );
  }

  if (includePageBreak) {
    blocks.push(createPageBreakParagraph(state.documentDoc));
  }
  prependBodyBlocks(state.selectW, state.documentDoc, blocks);
  await saveState(state);

  if (setCoreTitle) {
    await setCore({
      docPath: request.docPath,
      key: "title",
      value: title,
    });
  }

  return {
    frontpage: {
      template,
      title,
      subtitle,
      author,
      organization,
      date_text: dateText,
      page_break: includePageBreak,
      set_core_title: setCoreTitle,
    },
  };
}

async function setCore(request) {
  const key = String(request.key ?? "").trim();
  const value = String(request.value ?? "");
  const descriptor = CORE_PROPERTY_MAP[key];
  if (!descriptor) {
    fail(`Unsupported core property: ${key}`);
  }

  const state = await loadCoreState(request.docPath);
  ensureNamespaceDeclaration(state.root, descriptor.prefix, descriptor.namespace);
  if (descriptor.type === "datetime") {
    ensureNamespaceDeclaration(state.root, "xsi", XSI_NS);
    ensureNamespaceDeclaration(state.root, "dcterms", DCTERMS_NS);
  }

  const target = ensureCoreElement(state.selectCore, state.coreDoc, state.root, descriptor);

  if (descriptor.type === "int") {
    const numeric = Number.parseInt(value, 10);
    if (!Number.isInteger(numeric) || numeric < 1) {
      fail(`Invalid integer for core property '${key}': ${value}`);
    }
    target.textContent = String(numeric);
  } else if (descriptor.type === "datetime") {
    target.textContent = normalizedDatetimeValue(value, key);
    target.setAttributeNS(XSI_NS, "xsi:type", "dcterms:W3CDTF");
  } else {
    if (value.length > 255) {
      fail(`Exceeded 255 char limit for core property '${key}'.`);
    }
    target.textContent = value;
  }

  await saveCoreState(state);
  return {
    core: {
      key,
      value,
    },
  };
}

async function listBibliography(request) {
  const state = await loadState(request.docPath);
  const bib = extractBibliography(state.selectW, state.documentDoc);
  const entries = bib.entries.map((entry) => ({
    number: entry.number,
    key: entry.key,
    reference: entry.reference,
  }));
  return { entries };
}

async function addCitation(request) {
  const state = await loadState(request.docPath);
  const text = String(request.text ?? "").trim();
  if (!text) fail("Citation text cannot be empty.");

  const bib = extractBibliography(state.selectW, state.documentDoc);
  if (!bib.entries.length) fail("No bibliography entries found. Add bibliography entries before citing.");
  const resolved = resolveCitationEntries(bib.entries, request.sourceKeys ?? []);
  const marker = citationMarker(resolved.map((entry) => entry.number));

  const paragraph = createParagraph(state.documentDoc, text, {
    styleName: request.style ? String(request.style) : null,
  });
  appendCitationRuns(state.documentDoc, paragraph, resolved);
  appendParagraphToBody(state.selectW, state.documentDoc, paragraph);

  ensureFootnotesRelationship(state.selectRels, state.relsDoc);
  ensureContentTypeOverride(state.selectCT, state.contentTypesDoc);
  resolved.forEach((entry) => {
    upsertFootnote(state.selectW, state.footnotesDoc, entry.number, bibliographyFootnoteText(entry));
  });
  await saveState(state);

  return {
    citation: {
      text: `${text} ${marker}`,
      source_keys: resolved.map((entry) => entry.key),
      citation_numbers: resolved.map((entry) => entry.number),
      citation_marker: marker,
      style: request.style ?? null,
      format: "footnote",
    },
  };
}

async function citeParagraph(request) {
  const state = await loadState(request.docPath);
  const paragraphIndex = Number.parseInt(String(request.paragraphIndex), 10);
  if (Number.isNaN(paragraphIndex) || paragraphIndex < 0) {
    fail(`paragraph_index out of range: ${request.paragraphIndex}`);
  }

  const paragraphs = state.selectW("/w:document/w:body/w:p", state.documentDoc);
  if (paragraphIndex >= paragraphs.length) {
    fail(`paragraph_index out of range: ${paragraphIndex}`);
  }

  const bib = extractBibliography(state.selectW, state.documentDoc);
  if (!bib.entries.length) fail("No bibliography entries found. Add bibliography entries before citing.");
  const resolved = resolveCitationEntries(bib.entries, request.sourceKeys ?? []);
  const marker = citationMarker(resolved.map((entry) => entry.number));

  const paragraph = paragraphs[paragraphIndex];
  const textBefore = paragraphText(state.selectW, paragraph);
  appendCitationRuns(state.documentDoc, paragraph, resolved);

  ensureFootnotesRelationship(state.selectRels, state.relsDoc);
  ensureContentTypeOverride(state.selectCT, state.contentTypesDoc);
  resolved.forEach((entry) => {
    upsertFootnote(state.selectW, state.footnotesDoc, entry.number, bibliographyFootnoteText(entry));
  });
  await saveState(state);

  return {
    citation: {
      paragraph_index: paragraphIndex,
      text_before: textBefore,
      text_after: `${textBefore} ${marker}`.trim(),
      source_keys: resolved.map((entry) => entry.key),
      citation_numbers: resolved.map((entry) => entry.number),
      citation_marker: marker,
      format: "footnote",
    },
  };
}

async function execute(request) {
  switch (request.command) {
    case "summary":
      return summary(request);
    case "listParagraphs":
      return listParagraphs(request);
    case "addParagraph":
      return addParagraph(request);
    case "addHeading":
      return addHeading(request);
    case "addTable":
      return addTable(request);
    case "addFrontpage":
      return addFrontpage(request);
    case "setCore":
      return setCore(request);
    case "addBibliographyEntry":
      return addBibliographyEntry(request);
    case "listBibliography":
      return listBibliography(request);
    case "addCitation":
      return addCitation(request);
    case "citeParagraph":
      return citeParagraph(request);
    default:
      fail(`Unsupported command: ${request.command}`);
  }
}

async function main() {
  const payload = process.argv[2];
  if (!payload) fail("Missing request payload.");
  const request = JSON.parse(payload);
  const result = await execute(request);
  process.stdout.write(JSON.stringify({ ok: true, ...result }));
}

main().catch((error) => {
  process.stdout.write(
    JSON.stringify({
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    }),
  );
  process.exitCode = 1;
});
