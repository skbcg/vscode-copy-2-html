#!/bin/bash

echo "=== Clipboard Info ==="
osascript -e "clipboard info"

echo -e "\n=== Trying to read HTML ==="
osascript -e "the clipboard as «class HTML»" | perl -ne 'print chr foreach unpack("C*",pack("H*",substr($_,11,-3)))' 2>&1 | head -c 500

echo -e "\n\n=== Trying to read RTF ==="
osascript -e "the clipboard as «class RTF »" | perl -ne 'print chr foreach unpack("C*",pack("H*",substr($_,11,-3)))' 2>&1 | head -c 500

echo -e "\n\n=== Plain text ==="
osascript -e "the clipboard as text" 2>&1 | head -c 500

echo -e "\n\n=== Using pbpaste ==="
pbpaste | head -c 500
