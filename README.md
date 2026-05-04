# Precision Converter

A minimalist desktop application for converting **DOCX ↔ PDF**.

Built with Electron. Uses Microsoft Word's COM automation engine for maximum precision.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| **Microsoft Word** | Office 2013 or later. Both directions require it. |
| **Windows 10/11** | x64 only. |
| **Node.js** | For development only. Not needed by end users. |

---

## Development

```bash
# Install dependencies
npm install

# Run in development mode
npm start
```

---

## Build

```bash
# Produces both outputs in dist/
npm run dist
```



---

## Project Structure

```
converter-tool/
├── main.js                   # Electron main process — IPC, startup Word check
├── preload.js                # Context bridge between main and renderer
├── wordConverter.js          # Word COM automation via PowerShell
├── renderer/
│   ├── index.html            # UI
│   └── renderer.js           # UI logic — drag-drop, queue, progress
├── scripts/
│   ├── word_convert.ps1      # PowerShell script — drives Word.Application COM
│   └── gen-icon.js           # Generates assets/icon.ico (run by npm run dist)
└── assets/
    └── icon.ico              # Multi-size Windows icon (16, 32, 48, 256 px)
```

---

## How It Works

**DOCX → PDF**
Word opens the document and calls `ExportAsFixedFormat` with `OptimizeFor = Print` — embeds fonts, preserves margins, and respects all layout settings.

**PDF → DOCX**
Word opens the PDF using its built-in PDF import engine (Word 2013+) and saves the result as `.docx` via `SaveAs2`.

A single `Word.Application` COM instance handles the entire batch sequentially, then closes cleanly — no orphaned `WINWORD.EXE` processes.

---

## Notes

- If Microsoft Word is not detected at startup, a native error dialog appears before the window opens.
- Conversion errors are logged to `%APPDATA%\ectd-converter\conversion-errors.log`.
- For strict **PDF/A-1b** compliance (required by some eCTD submissions), set `UseISO19005_1 = $true` in `scripts/word_convert.ps1`.
