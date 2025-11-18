# WARP.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview

A VS Code extension that converts Microsoft Word HTML content into clean, semantic HTML when pasting. The extension intercepts paste commands, detects Word-generated HTML (with its namespace tags, inline styles, and proprietary markup), and transforms it into minimal, standards-compliant HTML.

**Platform**: macOS only (uses AppleScript for clipboard access)

## Commands

### Development
```bash
# Install dependencies
npm install

# Run linting
npm run lint

# Package extension for local installation
npx vsce package
```

### Testing & Debugging
```bash
# Debug extension in VS Code
# Press F5 in VS Code to launch Extension Development Host

# Test clipboard reading (debug utility)
./debug-clipboard.sh

# View debug logs
# Raw HTML is saved to /tmp/word-html-debug-*.html during paste operations
```

### Installation
```bash
# Install packaged extension
# In VS Code: Cmd+Shift+P → "Extensions: Install from VSIX..."
# Select word-to-html-paste-0.0.1.vsix
```

## Architecture

### Core Components

**extension.js** - Single-file architecture containing:
- Extension activation/deactivation lifecycle
- Command registration (toggle and paste commands)
- Clipboard reading via macOS AppleScript
- HTML cleaning pipeline

### HTML Processing Pipeline

The paste operation follows this flow:

1. **Clipboard Reading** (`pasteAsHtml`)
   - Attempts to read HTML format via AppleScript (`clipboard as «class HTML»`)
   - Falls back to RTF conversion using `textutil -convert html` if HTML unavailable
   - Falls back to plain text if neither HTML nor RTF available

2. **Word HTML Detection** (`isWordHtml`)
   - Checks for Word namespace markers (`xmlns:w`, `xmlns:o`, `class="Mso"`, etc.)
   - Detects excessive inline styles/classes typical of Word documents
   - Returns boolean indicating if aggressive cleaning is needed

3. **HTML Cleaning** (`cleanWordHtml`)
   - Strips Word namespace tags (`w:`, `o:`, `v:`, `m:`)
   - Removes style blocks, scripts, conditional comments
   - Removes all attributes from semantic tags (except `href` on links)
   - Converts `<b>` to `<strong>`, `<i>` to `<em>`
   - Strips `<div>` and `<span>` tags while preserving content
   - Normalizes whitespace

4. **Header Detection** (`wrapHeaderPrefixes`)
   - Detects patterns like `H2:`, `H2 Title:`, `H2 Tag:` in content
   - Converts to proper heading tags (`<h1>` through `<h6>`)
   - Unwraps malformed nested structures

5. **List Conversion** (`convertBulletPointsToLists`)
   - Finds paragraphs starting with `•`, `·`, or `-`
   - Groups consecutive bullet paragraphs into `<ul><li>` lists

6. **Attribute Stripping** (`stripUnwantedAttributes`)
   - Removes `data-*`, `id`, `class`, `style`, `aria-*`, etc.
   - Preserves only `href` attribute on anchor tags
   - Applied as final cleanup pass

7. **Formatting** (`formatBlockElements`)
   - Adds newlines after block-level closing tags
   - Ensures readable output structure

### Extension State

- **isEnabled**: Boolean toggle for paste interception (default: `false`)
- **statusBarItem**: Visual indicator in VS Code status bar showing enabled/disabled state

### File Type Filtering

Only activates for:
- Language IDs: `html`, `plaintext`
- File extensions: `.html`, `.txt`

### Key Behaviors

- **Keybinding Override**: Cmd+V is intercepted when extension is enabled and file type matches
- **Graceful Fallback**: If any step fails, falls back to standard paste behavior
- **Debug Logging**: All processing steps log to console; raw HTML saved to `/tmp/` for inspection
- **Temporary Files**: Uses `/tmp/` for RTF conversion intermediate files (cleaned up after use)

## Development Notes

- Extension uses CommonJS (`require`/`module.exports`)
- No TypeScript or build step - plain JavaScript
- JSDOM used for robust HTML parsing and manipulation
- AppleScript commands executed via Node's `child_process.exec`
- The HTML cleaning uses regex-based transformations for performance (no DOM parsing in main flow)
