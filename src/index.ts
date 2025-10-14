import {XMLParser} from 'fast-xml-parser';
import {ZipReader} from "./ZipReader";

export type ContentComplex =
    | { type: 'text'; data: string }
    | { type: 'image'; mimeType: string; data: string } // data is base64
    | { type: 'note'; data: string }
    | { type: 'comment'; data: string };

function getMimeTypeFromPath(filePath: string): string {
    const ext = filePath.split('.').pop()?.toLowerCase();
    const mimeMap: Record<string, string> = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'svg': 'image/svg+xml',
        'webp': 'image/webp',
        'tiff': 'image/tiff',
        'tif': 'image/tiff',
        'ico': 'image/x-icon'
    };
    return mimeMap[ext || ''] || 'application/octet-stream';
}

export class Page {

    constructor(public content: ContentComplex[] = [], public metadata: Record<string, any> = {}) {
    }

    get text() {
        return this.content.filter(x => x.type === 'text').reduce((p, c) => p + c.data, '');
    }

    get note() {
        return this.content.filter(x => x.type === 'note').reduce((p, c) => p + c.data, '');
    }

    get images() {
        return this.content.filter(x => x.type === 'image');
    }

    get comments() {
        return this.content.filter(x => x.type === 'comment');
    }
}

export class Document {
    readonly type: 'docx' | 'pptx' | 'xlsx' | 'unknown';
    readonly pages: Page[] = [];
    metadata: Record<string, any> = {};

    constructor(type: 'docx' | 'pptx' | 'xlsx' | 'unknown') {
        this.type = type;
    }
}

// --- XML Parser (secure config) ---
const xmlParser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseAttributeValue: true,
    trimValues: true,
    removeNSPrefix: true,
    allowBooleanAttributes: false,
    parseTagValue: true,
    stopNodes: ['*'], // prevents XXE
});

// --- Main function ---
export function parseOoxmlToDocument(zipPath: string): Document {
    // Step 1: Use a ZIP reader (we'll use a minimal sync version for clarity)
    const zip = new ZipReader(zipPath);
    const entries: Record<string, string> = {};

    // Only load XML/rels needed for parsing
    for (const name of zip.listEntries()) {
        if (/\.(xml|rels)$/i.test(name)) {
            const buf = zip.getEntry(name);
            if (buf) entries[name] = buf.toString('utf-8');
        }
    }

    // Step 2: Detect type
    let type: 'docx' | 'pptx' | 'xlsx' | 'unknown' = 'unknown';
    if (entries['word/document.xml']) type = 'docx';
    else if (entries['ppt/presentation.xml']) type = 'pptx';
    else if (entries['xl/workbook.xml']) type = 'xlsx';

    const doc = new Document(type);

    // Step 3: Extract core properties (metadata)
    const coreProps = entries['docProps/core.xml'];
    if (coreProps) {
        try {
            const parsed = xmlParser.parse(coreProps);
            const cp = parsed?.['cp:coreProperties'] || parsed?.['coreProperties'];
            if (cp) {
                doc.metadata = {
                    title: getString(cp, 'title'),
                    subject: getString(cp, 'subject'),
                    creator: getString(cp, 'creator'),
                    keywords: getString(cp, 'keywords'),
                    description: getString(cp, 'description'),
                    lastModifiedBy: getString(cp, 'lastModifiedBy'),
                    created: getString(cp, 'created'),
                    modified: getString(cp, 'modified'),
                };
            }
        } catch (e) {
            // ignore malformed core props
        }
    }
    doc.metadata.zipPath = zipPath;


    // Step 4: Parse content by type
    if (type === 'docx') {
        parseDocx(entries, doc, zip);
    } else if (type === 'pptx') {
        parsePptx(entries, doc, zip);
    } else if (type === 'xlsx') {
        parseXlsx(entries, doc);
    }


    return doc;
}

// --- DOCX (with images, page breaks, and structured content) ---
function parseDocx(entries: Record<string, string>, doc: Document, zip: ZipReader): void {
    const xml = entries['word/document.xml'];
    if (!xml) return;

    // --- 1. Load document relationships (for images) ---
    const relsXml = entries['word/_rels/document.xml.rels'];
    const imageRelsMap = new Map<string, string>(); // rId → media path
    if (relsXml) {
        const rels = xmlParser.parse(relsXml);
        const relList = arrayify(rels?.Relationships?.Relationship);
        for (const rel of relList) {
            const type = rel?.['@_Type'] || '';
            if (type.includes('/relationships/image')) {
                const target = rel?.['@_Target'];
                if (target) {
                    // Normalize: "media/image1.png" (already relative to word/)
                    const normalized = target.replace(/\\/g, '/');
                    imageRelsMap.set(rel['@_Id'], normalized);
                }
            }
        }
    }

    // --- 2. Parse main document ---
    const parsed = xmlParser.parse(xml);
    const body = parsed?.['document']?.['body'];
    if (!body) return;

    const paragraphs = arrayify(body['p']);
    let currentPageContent: ContentComplex[] = [];
    let pageIndex = 1;

    for (const p of paragraphs) {
        const blockContent: ContentComplex[] = [];

        // Extract text and images from this paragraph
        extractParagraphContent(p, imageRelsMap, zip, blockContent);

        // Check for page break
        let hasPageBreak = false;
        if (p?.r) {
            const runs = arrayify(p.r);
            for (const r of runs) {
                if (r?.br) {
                    const brElements = arrayify(r.br);
                    for (const br of brElements) {
                        if (br?.['@_type'] === 'page') {
                            hasPageBreak = true;
                            break;
                        }
                    }
                    if (hasPageBreak) break;
                }
            }
        }

        // Add this paragraph's content (text + images)
        currentPageContent.push(...blockContent);

        // Finalize page on break
        if (hasPageBreak) {
            doc.pages.push(new Page(currentPageContent, {index: pageIndex}));
            currentPageContent = [];
            pageIndex++;
        }
    }

    // Push final page
    if (currentPageContent.length > 0) {
        doc.pages.push(new Page(currentPageContent, {index: pageIndex}));
    }
}

// --- Extract mixed content (text + images) from a paragraph ---
function extractParagraphContent(
    p: any,
    imageRelsMap: Map<string, string>,
    zip: ZipReader,
    out: ContentComplex[]
): void {
    if (!p) return;

    // Helper to process runs and drawings
    function processElement(element: any): void {
        if (!element) return;

        // Handle text runs: <w:r> → <w:t>
        if (element.t != null) {
            let text: string | undefined;
            if (typeof element.t === 'string') {
                text = element.t;
            } else if (typeof element.t === 'object' && element.t['#text'] != null) {
                text = element.t['#text'];
            }
            if (text != null) {
                const trimmed = text.trim();
                if (trimmed) {
                    out.push({type: 'text', data: trimmed});
                }
            }
        }

        // Handle image references in drawings: <w:drawing> → <wp:inline> → <a:graphic> → <a:graphicData> → <pic:pic> → <pic:blipFill> → <a:blip r:embed="rIdX">
        if (element.drawing) {
            const drawings = arrayify(element.drawing);
            for (const drawing of drawings) {
                // Look for r:embed or r:id in blip
                const blip = findBlipInDrawing(drawing);
                if (blip) {
                    const embedId =
                        blip['@_r:embed'] ||
                        blip['@_embed'] ||
                        blip['@_r:id'] ||
                        blip['@_id'];

                    if (embedId && typeof embedId === 'string') {
                        const mediaPath = imageRelsMap.get(embedId);
                        if (mediaPath) {
                            const fullPath = `word/${mediaPath}`;
                            try {
                                const buffer = zip.getEntry(fullPath);
                                if (buffer) {
                                    const mimeType = getMimeTypeFromPath(mediaPath);
                                    const base64 = buffer.toString('base64');
                                    out.push({type: 'image', mimeType, data: base64});
                                }
                            } catch (e) {
                                console.warn(`[DOCX] Image not found: ${fullPath}`);
                            }
                        }
                    }
                }
            }
        }

        // Recurse into children (for complex structures)
        if (typeof element === 'object' && !Array.isArray(element)) {
            Object.values(element).forEach(val => {
                if (Array.isArray(val)) {
                    val.forEach(processElement);
                } else {
                    processElement(val);
                }
            });
        }
    }

    // Start processing from the paragraph root
    processElement(p);
}

// --- Find <a:blip> inside a <w:drawing> structure ---
function findBlipInDrawing(drawing: any): any | null {
    if (!drawing) return null;

    // DFS to find 'blip' key
    function dfs(obj: any): any {
        if (!obj || typeof obj !== 'object') return null;
        if (Array.isArray(obj)) {
            for (const item of obj) {
                const found = dfs(item);
                if (found) return found;
            }
            return null;
        }
        if ('blip' in obj) {
            return obj.blip;
        }
        for (const key in obj) {
            const found = dfs(obj[key]);
            if (found) return found;
        }
        return null;
    }

    return dfs(drawing);
}

// --- PPTX ---
function parsePptx(entries: Record<string, string>, doc: Document, zip: ZipReader): void {
    const presRelsXml = entries['ppt/_rels/presentation.xml.rels'];
    if (!presRelsXml) return;

    const presRels = xmlParser.parse(presRelsXml);
    const relList = arrayify(presRels?.Relationships?.Relationship);

    // Collect all slide and notes relationships
    const slideEntries: { rId: string; target: string; slideNumber: number }[] = [];
    const notesRelsMap = new Map<string, string>(); // rId → notes target

    for (const rel of relList) {
        const type = rel?.['@_Type'] || '';
        const rId = rel?.['@_Id'];
        const target = rel?.['@_Target'];
        if (!rId || !target) continue;

        if (type.includes('/notesSlide')) {
            notesRelsMap.set(rId, target);
        } else if (type.includes('/slide') && !type.includes('/slideLayout') && !type.includes('/slideMaster')) {
            // Extract number from target like "slides/slide5.xml"
            const match = target.match(/slide(\d+)\.xml$/i);
            const slideNumber = match ? parseInt(match[1], 10) : Infinity;
            slideEntries.push({rId, target, slideNumber});
        }
    }

    // ✅ SORT SLIDES BY SLIDE NUMBER
    slideEntries.sort((a, b) => a.slideNumber - b.slideNumber);

    // Process in correct order
    for (let i = 0; i < slideEntries.length; i++) {
        const {rId, target: slideTarget} = slideEntries[i];
        const slidePath = `ppt/${slideTarget}`;
        const slideXml = entries[slidePath];
        if (!slideXml) continue;

        const content: ContentComplex[] = [];

        // --- Text ---
        const slideObj = xmlParser.parse(slideXml);
        collectPptxTextStructured(slideObj, content);

        // --- Notes ---
        const notesTarget = notesRelsMap.get(rId);
        if (notesTarget) {
            const notesPath = `ppt/${notesTarget}`;
            const notesXml = entries[notesPath];
            if (notesXml) {
                const notesObj = xmlParser.parse(notesXml);
                const noteText = extractNotesText(notesObj);
                if (noteText) {
                    content.push({type: 'note', data: noteText});
                }
            }
        }

        // --- Images (from slide's .rels) ---
        const slideDir = slideTarget.substring(0, slideTarget.lastIndexOf('/') + 1);
        const slideFileName = slideTarget.substring(slideDir.length);
        const relsPath = `ppt/${slideDir}_rels/${slideFileName}.rels`;

        const relsXml = entries[relsPath];
        if (relsXml) {
            const rels = xmlParser.parse(relsXml);
            const slideRelList = arrayify(rels?.Relationships?.Relationship);

            for (const rel of slideRelList) {
                const relType = rel?.['@_Type'] || '';
                const relTarget = rel?.['@_Target'];
                if (!relTarget) continue;

                if (relType.includes('/relationships/image')) {
                    const normalizedTarget = relTarget
                        .replace(/^(\.\.\/)+/, '')
                        .replace(/\\/g, '/');
                    const imagePath = `ppt/${normalizedTarget}`;

                    try {
                        const imageBuffer = zip.getEntry(imagePath);
                        if (imageBuffer) {
                            const mimeType = getMimeTypeFromPath(imagePath);
                            const base64 = imageBuffer.toString('base64');
                            content.push({
                                type: 'image',
                                mimeType,
                                data: base64
                            });
                        }
                    } catch (e) {
                        console.warn(`[PPTX] Image not found: ${imagePath}`);
                    }
                }
            }
        }

        // Use 1-based index from sorted order (or use slideNumber if preferred)
        const pageIndex = i + 1;
        doc.pages.push(new Page(content, {
            index: pageIndex,
            slideId: slideTarget,
            slideNumber: slideEntries[i].slideNumber,
            hasNotes: !!notesTarget
        }));
    }
}

// --- Text extraction (structured) ---
function collectPptxTextStructured(obj: any, out: ContentComplex[]): void {
    if (obj == null) return;

    if (Array.isArray(obj)) {
        for (const item of obj) {
            collectPptxTextStructured(item, out);
        }
        return;
    }

    if (typeof obj === 'object') {
        // Handle <a:t> text nodes
        if (obj.t != null) {
            let text: string | undefined;
            if (typeof obj.t === 'string') {
                text = obj.t;
            } else if (typeof obj.t === 'object' && obj.t['#text'] != null) {
                text = obj.t['#text'];
            }
            if (text != null) {
                const trimmed = text.trim();
                if (trimmed) {
                    out.push({type: 'text', data: trimmed});
                }
            }
        }

        // Recurse into all values
        for (const value of Object.values(obj)) {
            collectPptxTextStructured(value, out);
        }
    }
}


// --- Notes extraction ---
function extractNotesText(notesObj: any): string {
    const textItems: ContentComplex[] = [];
    collectPptxTextStructured(notesObj, textItems);
    return textItems
        .filter(item => item.type === 'text')
        .map(item => item.data)
        .join(' ')
        .trim();
}

// --- XLSX ---
function parseXlsx(entries: Record<string, string>, doc: Document): void {
    const wbXml = entries['xl/workbook.xml'];
    const relsXml = entries['xl/_rels/workbook.xml.rels'];
    if (!wbXml || !relsXml) return;

    const wb = xmlParser.parse(wbXml);
    const rels = xmlParser.parse(relsXml);

    const sheetList = arrayify(wb?.workbook?.sheets?.sheet);
    const relList = arrayify(rels?.Relationships?.Relationship);

    for (let i = 0; i < sheetList.length; i++) {
        const sheet = sheetList[i];
        const rid = sheet?.['@_r:id'];
        const rel = relList.find((r: any) => r?.['@_Id'] === rid);
        if (!rel) continue;

        const sheetPath = `xl/${rel['@_Target']}`;
        const sheetXml = entries[sheetPath];
        if (!sheetXml) continue;

        const sheetName = sheet?.['@_name'] || `Sheet${i + 1}`;
        const rows = parseXlsxSheet(sheetXml);
        const lines = rows.map((row) => row.filter((cell) => cell).join('\t')).filter((line) => line);
        const content = lines.join('\n');

        doc.pages.push(new Page([{type: 'text', data: content}], {sheetName, index: i + 1}));
    }
}

function parseXlsxSheet(xml: string): string[][] {
    const parsed = xmlParser.parse(xml);
    const rowList = arrayify(parsed?.worksheet?.sheetData?.row);
    const result: string[][] = [];

    for (const row of rowList) {
        const cellList = arrayify(row?.c);
        const cells: string[] = [];
        for (const cell of cellList) {
            const v = cell?.v;
            cells.push(v ? String(v).trim() : '');
        }
        result.push(cells);
    }
    return result;
}

// --- Helpers ---
function getString(obj: any, key: string): string | undefined {
    const val = obj?.[key];
    if (typeof val === 'string') return val;
    else if (typeof val === 'object' && '#text' in val && val?.['#text']) return val['#text'];
    return undefined;
}

function arrayify<T>(item: T | T[]): T[] {
    if (!item) return [];
    return Array.isArray(item) ? item : [item];
}
