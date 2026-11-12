# Phase 3 — Step 3 Architecture Audit: Document Editing & Transformation

## 1. Executive Summary

This architecture audit evaluates the requirements, existing codebase, and target technical design for **Phase 3, Step 3: Document Editing & Transformation** in SmartShift.
SmartShift is a Windows desktop utility built on C#, .NET 10, WPF, PDFsharp (6.1.1), DocumentFormat.OpenXml (3.5.1), Docnet.Core (2.6.0), and SkiaSharp (4.153.1).

Phase 1, Phase 2, Phase 3 Step 1 (Advanced PDF Conversion Engine), and Phase 3 Step 2 (Advanced Office Conversion Engine) are fully completed and verified.

### Baseline Status:
- **Automated Tests:** 483 passed, 0 failed, 0 skipped across all test assemblies (`SmartShift.Core.Tests`, `SmartShift.Processing.Tests`, `SmartShift.Windows.Tests`, `SmartShift.Infrastructure.Tests`).
- **Build Quality:** 0 warnings, 0 errors in both **Debug** and **Release** configurations (`net10.0` / `net10.0-windows`).

---

## 2. Existing Relevant Functionality & Reusable Services

### 2.1 Existing PDF Architecture
- **PDFsharp Integration:** `PdfDocument`, `PdfReader`, `PdfPage`, `XGraphics`, `XFont`, `XSolidBrush`, `XStringFormats` are actively used in `PdfMergeEngine`, `PdfSplitEngine`, `PdfPageOrganizerEngine`, and `DocxToPdfConverter`.
- **Metadata Support in PDFsharp:** `PdfDocument.Info` natively supports `Title`, `Author`, `Subject`, `Keywords`, `Creator`, and `Producer`.
- **Watermarking Support in PDFsharp:** `XGraphics` can draw semi-transparent text with arbitrary rotation angle (e.g. diagonal 45° or horizontal 0°), custom font size, color, alpha transparency, and page ranges.
- **Atomic File Handling & Integrity:** `IOutputPathAllocator` and `.partial` promotion pattern ensures zero data loss. SHA-256 hash checks verify input file immutability.

### 2.2 Existing OpenXML (DOCX) Architecture
- **DocumentFormat.OpenXml 3.5.1:** Utilized across `DocxToPdfConverter`, `DocxTextExtractor`, `DocxCompressor`, `PdfToDocxConverter`, `MultiImageDocumentBuilder`.
- **Package Properties:** `WordprocessingDocument.PackageProperties` provides direct, standards-compliant read/write access to Core File Properties (`Title`, `Subject`, `Creator`, `Keywords`, `Description`, `Category`).
- **MainDocumentPart & Header/Footer:** Open XML allows manipulating `MainDocumentPart`, adding `HeaderPart` or `FooterPart`, defining `HeaderReference` / `FooterReference` in `SectionProperties`, and embedding dynamic field codes (`w:fldSimple w:instr="PAGE"` and `w:fldSimple w:instr="NUMPAGES"`).
- **Watermarks in WordprocessingML:** Text watermarks can be added via `HeaderPart` containing VML shapes (`<v:shape type="#_x0000_t136"><v:textpath string="..."/></v:shape>`) or DrawingML WordArt shapes with semi-transparency and rotation.

### 2.3 Catalog, Capabilities, and Workflow Infrastructure
- **`ActionIds` & `GoalIds`:** Central string constants defining all operations.
- **`FileCapabilityRegistry`:** Exposes format capabilities, available action IDs, and supported conversion targets.
- **`ActionOptionValidator`:** Validates user-supplied parameters prior to processor execution.
- **`ProcessorCatalog`:** Central registry mapping `ActionId` to `IFileProcessor`.
- **`IProcessingHistoryStore`:** Persists operation records (`OperationRecord`) categorized by `OperationCategory.Document` and `OperationCategory.Pdf`.

---

## 3. Analysis of Required Workflows & Missing Capabilities

### 3.1 DOCX Document Merger
- **Current State:** Only `PdfMergeEngine` exists for PDF merging. No DOCX merging service exists.
- **Technical Challenge:** Naive XML concatenation breaks documents because:
  1. `SectionProperties` (`w:sectPr`): Word only uses the body-level `w:sectPr` for the last section; preceding sections must have their `w:sectPr` placed inside a paragraph (`w:p/w:pPr/w:sectPr`) to preserve page dimensions, margins, and orientation across merged parts.
  2. Image & Media Relationships (`rId`): Part IDs collision causes missing images or corrupt packages. Target document must copy `ImagePart` streams from source documents and assign fresh unique relationship IDs (`rId`), updating `a:blip r:embed="rIdX"` references.
  3. Styles & Numbering: Normal styles and font definitions should be reconciled or retained without crashing Word.
  4. Non-Destructive: Source files must remain untouched (pre/post SHA-256 verification).
- **Solution:** Implement `DocxMergeService` / `IDocxMergeService` with relationship re-mapping, section break retention, and structured validation reporting.

### 3.2 DOCX Page Numbering
- **Current State:** None exists.
- **Technical Challenge:** Page numbers in Word are dynamic fields evaluated by the layout engine (Word / LibreOffice). They must be placed in `FooterPart` or `HeaderPart` linked to the document's `SectionProperties`.
- **Solution:** Implement `DocxPageNumberService` / `IDocxPageNumberService`:
  - Placements: `Footer` or `Header`.
  - Alignments: `Left`, `Center`, `Right`.
  - Formats: `NumberOnly` ("1"), `PageNumber` ("Page 1"), `PageNumberTotal` ("Page 1 of N").
  - Preserves existing section headers/footers when present or creates default section references.

### 3.3 DOCX Watermark
- **Current State:** None exists.
- **Technical Challenge:** Word watermarks are stored inside headers so they appear on all pages of a section without interfering with main body text flow.
- **Solution:** Implement `DocxWatermarkService` / `IDocxWatermarkService`:
  - Inserts a watermark element into the `HeaderPart` with configurable text, font size, color, opacity, and rotation (diagonal ~315°/45° or horizontal 0°).
  - Supports non-destructive replacement/removal of prior watermarks.

### 3.4 DOCX Metadata Editor
- **Current State:** `InspectFileProcessor` can read basic file info, but no dedicated metadata editor exists.
- **Solution:** Implement `DocxMetadataService` / `IDocxMetadataService`:
  - Reads and updates Open XML `PackageProperties`: `Title`, `Subject`, `Author` (`Creator`), `Keywords`, `Comments` (`Description`), and `Category`.
  - Non-destructive atomic write.

### 3.5 PDF Watermark
- **Current State:** None exists.
- **Technical Challenge:** Drawing text across multi-page PDF documents with accurate rotation, font size, opacity, and page filtering ("all", "odd", "even", range "1-5").
- **Solution:** Implement `PdfWatermarkService` / `IPdfWatermarkService`:
  - Uses PDFsharp `XGraphics` and `XFont`.
  - Draws semi-transparent text with alpha channel rotation at page center.
  - Page selection engine supporting all, odd, even, or range syntax.

### 3.6 PDF Metadata Editor
- **Current State:** None exists.
- **Solution:** Implement `PdfMetadataService` / `IPdfMetadataService`:
  - Reads and updates `PdfDocument.Info` fields (`Title`, `Author`, `Subject`, `Keywords`, `Creator`, `Producer`).
  - Preserves all PDF pages and content streams intact.

---

## 4. Proposed Implementation Architecture

```
SmartShift.Core
  ├── Actions
  │    ├── ActionIds.cs (MergeDocx, AddDocxPageNumbers, AddDocxWatermark, EditDocxMetadata, AddPdfWatermark, EditPdfMetadata)
  │    ├── GoalIds.cs
  │    └── ActionOptionValidator.cs (Option validation for page numbers, watermarks, metadata)
  └── Capabilities
       └── FileCapabilityRegistry.cs (Register actions under Docx and Pdf)

SmartShift.Processing
  ├── Documents
  │    ├── Editing
  │    │    ├── IDocxEditingServices.cs (IDocxMergeService, IDocxPageNumberService, IDocxWatermarkService, IDocxMetadataService)
  │    │    ├── DocxEditingModels.cs (Options, Results, Formats, Positions)
  │    │    ├── DocxMergeService.cs
  │    │    ├── DocxPageNumberService.cs
  │    │    ├── DocxWatermarkService.cs
  │    │    ├── DocxMetadataService.cs
  │    │    └── Processors
  │    │         ├── DocxMergeProcessor.cs
  │    │         ├── DocxPageNumberProcessor.cs
  │    │         ├── DocxWatermarkProcessor.cs
  │    │         └── DocxMetadataProcessor.cs
  ├── Pdf
  │    ├── Editing
  │    │    ├── IPdfEditingServices.cs (IPdfWatermarkService, IPdfMetadataService)
  │    │    ├── PdfEditingModels.cs
  │    │    ├── PdfWatermarkService.cs
  │    │    ├── PdfMetadataService.cs
  │    │    └── Processors
  │    │         ├── PdfWatermarkProcessor.cs
  │    │         └── PdfMetadataProcessor.cs
  ├── ProcessorCatalog.cs (Registers new processors)
  └── AppServices.cs (DI container registration)
```

---

## 5. Dependencies & Licensing Assessment

No new third-party dependencies are required. All proposed functionality relies on existing production libraries:
- **`DocumentFormat.OpenXml` (3.5.1):** MIT License. Industry standard for Open XML documents.
- **`PDFsharp` (6.1.1):** MIT License. Existing vector PDF engine used throughout SmartShift.
- **.NET 10 BCL:** `System.IO.Packaging`, `System.Security.Cryptography`, `System.Text.Json`.

All processing is 100% local, offline, and free from cloud or online API requirements.

---

## 6. Risks and Mitigation Strategies

1. **Risk:** DOCX merging losing images or cross-references.
   - *Mitigation:* Explicitly iterate and deep-copy `ImagePart` elements to the target package and re-point `rId` in DrawingML.
2. **Risk:** DOCX section break formatting loss.
   - *Mitigation:* Convert `w:sectPr` of all preceding documents into paragraph-level section breaks `w:p/w:pPr/w:sectPr`.
3. **Risk:** Word rejecting corrupt VML or drawing watermark markup.
   - *Mitigation:* Use validated standard WordprocessingML Drawing / VML header structures compatible with Microsoft Word and LibreOffice.
4. **Risk:** Accidental source document modification.
   - *Mitigation:* Enforce input immutability by computing SHA-256 pre and post operation and asserting equality.
5. **Risk:** Unhandled PDF encryption.
   - *Mitigation:* Validate `SecuritySettings.IsEncrypted` and return structured failure reports without crashing.

---

## 7. Baseline Verification Sign-Off

- **Baseline Tests:** 483 passed, 0 failed, 0 skipped.
- **Baseline Build:** 0 warnings, 0 errors (Debug & Release).
- **Audit Decision:** APPROVED to proceed with Phase 3 Step 3 implementation.
