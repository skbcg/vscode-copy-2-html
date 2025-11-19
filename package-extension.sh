#!/bin/bash

# Helper script to package the VS Code extension
# Handles node_modules.nosync symlink properly

set -e  # Exit on error

echo "📦 Packaging Word to HTML Paste extension..."

# Remove any duplicate node_modules files that iCloud might have created
rm -f "node_modules 2" "node_modules 3" 2>/dev/null || true

# Temporarily copy node_modules for packaging (vsce doesn't handle symlinks well)
if [ -L "node_modules" ]; then
    echo "🔗 Converting symlink to directory for packaging..."
    rm node_modules
    cp -R node_modules.nosync node_modules
    RESTORE_SYMLINK=true
else
    RESTORE_SYMLINK=false
fi

# Package the extension
npx vsce package --allow-star-activation

# Restore symlink
if [ "$RESTORE_SYMLINK" = true ]; then
    echo "🔗 Restoring symlink..."
    rm -rf node_modules
    ln -s node_modules.nosync node_modules
fi

echo "✅ Packaging complete!"
ls -lh word-to-html-paste-*.vsix | tail -1
