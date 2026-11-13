<p align="center">
  <img src="src/SmartShift.App/Assets/logo.png" alt="ZUNO Logo" width="128" height="128" />
</p>

<h1 align="center">ZUNO</h1>

<p align="center">
  <strong>Universal, Privacy-First Modern File Workspace for Windows</strong><br>
  <em>Convert, edit, organize, and transform documents and media with zero cloud dependencies.</em>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows%2010%20%7C%2011-0078D6?style=flat-square" alt="Platform" />
  <img src="https://img.shields.io/badge/.NET-10.0-512BD4?style=flat-square" alt=".NET 10" />
  <img src="https://img.shields.io/badge/Theme-%23cc4778%20%26%20White-cc4778?style=flat-square" alt="Theme" />
  <img src="https://img.shields.io/badge/Privacy-100%25%20Local%20%7C%20Zero%20Uploads-success?style=flat-square" alt="Privacy" />
  <img src="https://img.shields.io/badge/License-MIT-blue?style=flat-square" alt="License" />
</p>

---

## 🌟 Overview

**ZUNO** is a high-performance Windows desktop application built for professionals who need reliable, privacy-first document and media workflows. Designed with a signature **`#cc4778` Berry & Pure White** aesthetic, ZUNO provides an intuitive, non-destructive workspace where your files never leave your machine.

Everything in ZUNO runs **100% locally on your computer**:
- 🚫 **No cloud uploads**
- 🚫 **No subscriptions or accounts**
- 🚫 **No hidden telemetries**
- 🛡️ **Zero risk to originals** (all writes are collision-safe, atomic, and non-destructive)

---

## ✨ Features & Capabilities

### 1. 🔄 Universal Document Conversion
Convert between industry-standard document formats with structure preservation:
- **PDF to Office**: High-fidelity PDF to DOCX and PDF to PPTX conversion.
- **Office to PDF**: DOCX to PDF and PPTX to PDF rendering.
- **PDF to Raster Images**: Export PDF pages to PNG, JPEG, and WebP with configurable DPI (72, 100, 150, 300 DPI).
- **PDF Text Extraction**: Fast, accurate text layout extraction with reading order detection.
- **Conversion Fidelity & Quality Validator**: Automated pre/post-conversion quality reporting verifying page counts, table structures, and text completeness.

### 2. 📑 Document Editing & Organization
- **PDF Page Organizer**: Visually reorder, rotate (90° / 180° / 270°), extract, or delete individual pages with instantaneous thumbnail previews.
- **PDF Splitter**: Split by page ranges, single pages, or bookmarks.
- **PDF Merger**: Combine multiple PDF files in custom order with atomic output generation.
- **Watermarking & Metadata**: Non-destructive text watermarking and metadata sanitation.

### 3. 🎨 Interactive Image Studio & Batch Engine
- **Interactive Editor**: Crop with aspect-ratio locking (Free, 1:1, 4:3, 16:9, Original), 90° rotation, horizontal/vertical flipping, and Mitchell cubic resampling.
- **Batch Processing**: Convert formats (PNG, JPEG, WebP), strip EXIF/GPS metadata, and batch compress with live multi-file progress tracking.
- **Multi-Image Document Builder**: Stitch multiple images into clean, standardized multi-page PDF documents or PPTX presentations.

### 4. 📊 Presentation Studio
- Full-featured deck builder and outline generator for PowerPoint (`.pptx`).
- Dynamic layout presets (Title, Bullets, Comparison, Showcase, Blank).
- Cohesive typography and color palettes tailored for executive presentations.

### 5. 🎯 Contextual Radial Action Menu
- Lightning-fast radial wheel orbiting your active file.
- Action hotkeys (`1` through `7`) for muscle-memory productivity.
- Instant keyboard shortcuts (`Ctrl+O` to open, `Ctrl+,` for preferences, `F1` for quick help, `Enter` to run, `Esc` to cancel).

---

## 🎨 Design Philosophy

ZUNO breaks away from generic dark neon AI themes with a fresh, focused interface:
- **Berry Accent (`#cc4778`)**: Balanced, energetic accent color used for primary actions, focus rings, and status indicators.
- **Crisp Light Surfaces**: High-contrast, clean surfaces (`#FFFFFF` and `#F8F9FA`) with slate typography (`#0F172A`) ensuring readability.
- **Smooth Feedback**: Responsive transitions, informative hover states, and clear quality diagnostics.

---

## 🚀 Getting Started

### Prerequisites
- **Operating System**: Windows 10 (version 1809+) or Windows 11
- **Runtime**: [.NET 10.0 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)

### Building from Source

Clone the repository and build using the .NET CLI:

```powershell
# Restore dependencies
dotnet restore SmartShift.sln

# Build Release binary
dotnet build SmartShift.sln -c Release
```

### Running the Application

```powershell
dotnet run --project src\SmartShift.App\SmartShift.App.csproj -c Release
```

Optional CLI flags:
- `--diagnostics`: Enables developer diagnostics window and inspection tools.

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
| --- | --- |
| `Ctrl + O` | Open file browser to add files |
| `Ctrl + ,` | Open application preferences / settings |
| `F1` | Open quick start guide & keyboard reference |
| `1` – `7` | Select radial action directly |
| `Enter` | Execute currently selected action |
| `Delete` | Remove selected file from queue |
| `Esc` | Cancel ongoing operation or close active overlay |

---

## 🔒 Safety & Non-Destructive Guarantee

1. **Original Protection**: ZUNO strictly treats all input files as read-only. Source files are never modified, renamed, or overwritten in place.
2. **Collision-Free Outputs**: All processed outputs are written into dedicated output storage with automatic collision avoidance (`file (1).ext`).
3. **Atomic Writes**: Files are written to temporary `.partial` files and atomically renamed only upon 100% successful verification, preventing corrupted outputs.
4. **Offline Isolation**: ZUNO makes zero external network requests. Your sensitive legal, financial, and personal documents never touch the internet.

---

## 📄 License

Licensed under the [MIT License](LICENSE).
