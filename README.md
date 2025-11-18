# Word to HTML Paste

A VS Code extension that intelligently converts content from Microsoft Word documents into clean, semantic HTML when pasting. No more messy inline styles, proprietary markup, or bloated HTML from Word!

## ✨ Features

- **Smart Conversion**: Automatically detects Word HTML and converts it to clean, semantic HTML
- **Toggle Control**: Easy on/off toggle via status bar or Command Palette
- **File Type Detection**: Only activates for HTML and text files
- **Attribute Stripping**: Removes all unwanted attributes (classes, IDs, data-*, styles, etc.)
- **Header Recognition**: Converts `H1:`, `H2:`, etc. prefixes to proper heading tags
- **Format Preservation**: Maintains semantic structure while removing Word-specific artifacts
- **Status Bar Indicator**: Visual feedback showing whether the feature is enabled or disabled

## 🎯 What Gets Cleaned

### Removed
- Microsoft Word namespace tags (`w:`, `o:`, `v:`, `m:`)
- All inline styles and CSS
- Class and ID attributes
- Data attributes
- ARIA attributes
- Meta tags and link tags
- Empty paragraphs and headings
- Font tags
- Excessive whitespace

### Preserved
- **Headings**: `<h1>` through `<h6>`
- **Paragraphs**: `<p>`
- **Lists**: `<ul>`, `<ol>`, `<li>`
- **Text formatting**: `<strong>`, `<em>`, `<u>`
- **Links**: `<a>` tags (with `href` attribute)
- **Tables**: `<table>`, `<thead>`, `<tbody>`, `<tr>`, `<th>`, `<td>`
- **Other semantic tags**: `<blockquote>`, `<pre>`, `<code>`, `<br>`

## 🚀 Installation

Since this is a personal extension not published to the VS Code Marketplace, install it manually:

1. Clone or download this repository
2. Open the project folder in your terminal
3. Run:
   ```bash
   npm install
   ```
4. Package the extension:
   ```bash
   npx vsce package
   ```
5. Install the `.vsix` file in VS Code:
   - Open VS Code
   - Go to Extensions (Cmd+Shift+X)
   - Click the "..." menu → "Install from VSIX..."
   - Select the generated `.vsix` file

## 📖 Usage

1. **Copy content** from a Microsoft Word document
2. **Open** a `.html` or `.txt` file in VS Code
3. **Enable the extension** by either:
   - Clicking the status bar item in the bottom right corner
   - Running "Toggle Word to HTML Paste" from the Command Palette (Cmd+Shift+P)
4. **Paste** with Cmd+V (Mac) or Ctrl+V (Windows/Linux)
5. The content will be automatically converted to clean HTML!

### Status Bar

The extension shows its status in the bottom right of VS Code:
- **$(clippy) HTML Paste: $(check)** - Feature is enabled
- **$(clippy) HTML Paste: $(x)** - Feature is disabled

Click the status bar item to quickly toggle the feature on or off.

## 💡 Examples

### Before (Word HTML)
```html
<p class="MsoNormal" style="margin:0in;font-size:11pt;font-family:Calibri,sans-serif">
  <span style="font-size:16.0pt;font-weight:bold">This is a heading</span>
</p>
<p class="MsoNormal" style="margin:0in;font-size:11pt;font-family:Calibri,sans-serif">
  This is some <strong style="color:#ff0000">bold</strong> text.
</p>
```

### After (Clean HTML)
```html
<h2>This is a heading</h2>
<p>This is some <strong>bold</strong> text.</p>
```

## ⚙️ Technical Details

- **Platform**: macOS (uses AppleScript for clipboard access)
- **VS Code Version**: Requires ^1.74.0 or higher
- **Dependencies**:
  - `jsdom` - HTML parsing and manipulation
  - `clipboardy` - Cross-platform clipboard access
- **Activation**: Activates on first use, low overhead

### macOS-Specific Features

The extension uses native macOS clipboard APIs via AppleScript to:
- Detect HTML content in the clipboard
- Fall back to RTF conversion if HTML is not available
- Handle multiple clipboard data types

## 🛠️ Development

### Setup
```bash
npm install
```

### Linting
```bash
npm run lint
```

### Debugging
1. Open the project in VS Code
2. Press F5 to launch Extension Development Host
3. Check the Debug Console for logs

A debug script is included for clipboard testing:
```bash
./debug-clipboard.sh
```

## 📁 Project Structure

```
word-to-html-paste/
├── extension.js          # Main extension logic
├── package.json          # Extension manifest
├── debug-clipboard.sh    # Debugging utility
├── LICENSE              # MIT License
└── README.md            # This file
```

## 🐛 Known Limitations

- **macOS only**: The clipboard HTML reading uses AppleScript, which is macOS-specific
- **File types**: Only activates for `.html` and `.txt` files (by design)
- **Complex tables**: Very complex Word tables may need manual adjustment

## 📝 License

MIT License - See [LICENSE](LICENSE) file for details

## 🤝 Contributing

This is a personal project, but suggestions and improvements are welcome! Feel free to open issues or submit pull requests.
---

**Note**: This extension is designed for personal use and is not published on the VS Code Marketplace. If you find it useful, feel free to fork and customize it for your needs!
