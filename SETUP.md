# Setup and Installation Guide

## How to Run the Extension

### Method 1: Run in Development Mode (Recommended for testing)

1. Open this folder in VS Code:
   ```bash
   code "path/to/word-to-html-paste"
   ```

2. Press `F5` or go to **Run > Start Debugging**
   - This will open a new VS Code window with the extension loaded
   - The new window title will say "[Extension Development Host]"

3. In the new window:
   - Create or open a `.html` or `.txt` file
   - Look for the status bar item in the bottom right (shows "HTML Paste: ✗")
   - Click it to enable (or use Command Palette: `Cmd+Shift+P` → "Toggle Word to HTML Paste")
   - Copy content from Word and paste it!

### Method 2: Install Locally (For permanent use)

1. Package the extension:
   ```bash
   npm install -g @vscode/vsce
   cd "path/to/word-to-html-paste"
   vsce package
   ```

2. Install the `.vsix` file:
   - In VS Code: `Cmd+Shift+P` → "Extensions: Install from VSIX..."
   - Select the generated `word-to-html-paste-0.0.1.vsix` file

3. Restart VS Code

## Usage

1. **Enable the extension:**
   - Click the status bar item (bottom right corner)
   - Or: `Cmd+Shift+P` → "Toggle Word to HTML Paste"
   - Status will show a checkmark when enabled

2. **Copy from Word:**
   - Select and copy content from any Word document

3. **Paste in VS Code:**
   - Open a `.html` or `.txt` file
   - Paste with `Cmd+V`
   - Content will be converted to clean HTML

4. **Format (optional):**
   - Run Prettier: `Cmd+Shift+P` → "Format Document"

## Troubleshooting

- **Paste not converting?** Make sure:
  - The extension is enabled (check status bar)
  - You're in a `.html` or `.txt` file
  - The clipboard contains HTML content from Word

- **Extension not loading?** 
  - Check the Output panel: View → Output → "Extension Host"
  - Look for any error messages

- **Want to see the original paste?**
  - Toggle the extension off before pasting
