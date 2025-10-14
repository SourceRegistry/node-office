# OOXML Document Parser (Work in Progress)

A TypeScript library for parsing Office Open XML (`.docx`, `.pptx`, `.xlsx`) documents into structured, page-oriented content.

> ⚠️ **This is a work in progress** — APIs and behavior may change. Not yet suitable for production use.

## Features

- **DOCX**: Extracts text split by explicit page breaks (`<w:br w:type="page"/>`) and embedded images (as Base64 with MIME type).
- **PPTX**: Extracts slide text, speaker notes, and images (Base64 + MIME type), with slides correctly ordered by filename (`slide1.xml`, `slide2.xml`, etc.).
- **XLSX**: Basic sheet-by-sheet text extraction (tab-separated values).
- Structured output using `ContentComplex` for rich content handling:
  ```ts
  type ContentComplex =
    | { type: 'text'; data: string }
    | { type: 'image'; mimeType: string; data: string } // Base64
    | { type: 'note'; data: string }
    | { type: 'comment'; data: string };
  ```

## Usage

```ts
import { parseOoxmlToDocument } from './parser';

const doc = parseOoxmlToDocument('path/to/document.docx');
console.log(doc.pages[0].content);
```

## Limitations

- **DOCX**: Only splits on *explicit* page breaks (not layout-based pagination).
- **Images**: Relies on relationship files; may include unused media if present.
- **Comments/Tables/Headers**: Not fully supported yet.
- **Security**: Assumes trusted input; not hardened against malicious OOXML.

## Roadmap

- [ ] Add table support (DOCX/PPTX)
- [ ] Extract image alt text and captions
- [ ] Support headers/footers (DOCX)
- [ ] Improve comment extraction
- [ ] Add tests and validation

---

**Use at your own risk. Contributions and feedback welcome!**
