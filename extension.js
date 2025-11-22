const vscode = require('vscode');
const { JSDOM } = require('jsdom');
const { exec } = require('child_process');
const { promisify } = require('util');
const fs = require('fs');

const execAsync = promisify(exec);

let isEnabled = false;
let statusBarItem;

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
    console.log('Word to HTML Paste extension is now active');

    // Create status bar item
    statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Right, 100);
    statusBarItem.command = 'word-to-html-paste.toggle';
    updateStatusBar();
    statusBarItem.show();
    context.subscriptions.push(statusBarItem);

    // Register toggle command
    let toggleCommand = vscode.commands.registerCommand('word-to-html-paste.toggle', function () {
        isEnabled = !isEnabled;
        updateStatusBar();
        vscode.window.showInformationMessage(
            `Word to HTML Paste: ${isEnabled ? 'Enabled' : 'Disabled'}`
        );
    });

    // Register paste command that will be bound to Cmd+V when enabled
    let pasteCommand = vscode.commands.registerTextEditorCommand(
        'word-to-html-paste.paste',
        async (textEditor, edit) => {
            const languageId = textEditor.document.languageId;
            const fileName = textEditor.document.fileName;
            
            // Check if file type is supported
            const supportedLanguages = ['html', 'plaintext'];
            const supportedExtensions = ['.html', '.txt'];
            const isSupportedFile = supportedLanguages.includes(languageId) || 
                                   supportedExtensions.some(ext => fileName.endsWith(ext));

            if (isEnabled && isSupportedFile) {
                await pasteAsHtml(textEditor, edit);
            } else {
                // Do default paste by inserting clipboard text
                const clipboardContent = await vscode.env.clipboard.readText();
                if (clipboardContent) {
                    const selection = textEditor.selection;
                    await textEditor.edit(editBuilder => {
                        editBuilder.replace(selection, clipboardContent);
                    });
                }
            }
        }
    );

    context.subscriptions.push(toggleCommand, pasteCommand);
}

function updateStatusBar() {
    statusBarItem.text = `$(clippy) HTML Paste: ${isEnabled ? '$(check)' : '$(x)'}`;
    statusBarItem.tooltip = `Word to HTML Paste is ${isEnabled ? 'enabled' : 'disabled'}. Click to toggle.`;
}

async function pasteAsHtml(textEditor, edit) {
    try {
        // Try to read HTML from clipboard using AppleScript (macOS)
        let clipboardContent = '';
        let hadHtmlContent = false;
        
        try {
            // First check what's available on the clipboard
            const { stdout: classInfo } = await execAsync(
                'osascript -e "clipboard info"'
            );
            console.log('Clipboard info:', classInfo);
            
            // Try multiple approaches to get HTML content
            // Approach 1: Direct HTML class
            try {
                const { stdout } = await execAsync(
                    'osascript -e "the clipboard as «class HTML»" | perl -ne \'print chr foreach unpack("C*",pack("H*",substr($_,11,-3)))\' '
                );
                
                if (stdout && stdout.trim().length > 0) {
                    clipboardContent = stdout;
                    hadHtmlContent = true;
                    console.log('Read HTML from clipboard (method 1), length:', clipboardContent?.length);
                }
            } catch (e1) {
                console.log('Method 1 failed:', e1.message);
            }
            
            // Approach 2: Try reading as RTF if HTML failed
            if (!hadHtmlContent && classInfo.includes('RTF')) {
                try {
                    const { stdout: rtfContent } = await execAsync(
                        'osascript -e "the clipboard as «class RTF »" | perl -ne \'print chr foreach unpack("C*",pack("H*",substr($_,11,-3)))\' '
                    );
                    
                    if (rtfContent && rtfContent.trim().length > 0) {
                        // Convert RTF to HTML using textutil
                        const tempRtfPath = `/tmp/temp_${Date.now()}.rtf`;
                        
                        // Write RTF to temp file
                        fs.writeFileSync(tempRtfPath, rtfContent);
                        
                        // Convert using textutil
                        const { stdout: htmlContent } = await execAsync(`textutil -convert html -stdout "${tempRtfPath}"`);
                        
                        // Clean up temp file
                        fs.unlinkSync(tempRtfPath);
                        
                        if (htmlContent && htmlContent.trim().length > 0) {
                            clipboardContent = htmlContent;
                            hadHtmlContent = true;
                            console.log('Read HTML from RTF conversion, length:', clipboardContent?.length);
                        }
                    }
                } catch (e2) {
                    console.log('RTF conversion failed:', e2.message);
                }
            }
            
            // If still no HTML content, throw to use fallback
            if (!hadHtmlContent || !clipboardContent || clipboardContent.trim().length === 0) {
                throw new Error('No HTML or RTF content found');
            }
        } catch (error) {
            // Fall back to plain text if HTML not available
            console.log('No HTML in clipboard or error reading HTML:', error.message);
            clipboardContent = await vscode.env.clipboard.readText();
            console.log('Using plain text fallback, length:', clipboardContent?.length);
        }
        
        console.log('First 200 chars:', clipboardContent?.substring(0, 200));
        console.log('Contains HTML tags:', clipboardContent?.includes('<') && clipboardContent?.includes('>'));
        console.log('Had HTML content:', hadHtmlContent);
        
        if (!clipboardContent || clipboardContent.trim().length === 0) {
            console.log('Empty clipboard content, aborting');
            vscode.window.showWarningMessage('Clipboard is empty');
            return;
        }

        // Check if content looks like HTML
        if (!clipboardContent.includes('<') || !clipboardContent.includes('>')) {
            // Not HTML, do regular paste
            console.log('Not HTML, doing regular paste');
            const selection = textEditor.selection;
            await textEditor.edit(editBuilder => {
                editBuilder.replace(selection, clipboardContent);
            });
            return;
        }
        
        // Check if this is VS Code's syntax-highlighted copy format
        // VS Code wraps copied code in divs with syntax highlighting and HTML-encodes it
        // Detect early and use plain text to avoid processing overhead
        const isVSCodeFormat = clipboardContent.includes('&lt;') && 
                              clipboardContent.includes('&gt;') && 
                              (clipboardContent.includes('cascadia code') || 
                               clipboardContent.includes('monospace') || 
                               clipboardContent.includes('<meta charset'));
        
        if (isVSCodeFormat) {
            console.log('Detected VS Code format, using plain text (no processing needed)');
            const plainText = await vscode.env.clipboard.readText();
            const selection = textEditor.selection;
            await textEditor.edit(editBuilder => {
                editBuilder.replace(selection, plainText);
            });
            return;
        }
        
        console.log('Processing HTML...');
        
        // Count and log anchor tags for debugging
        const anchorCount = (clipboardContent?.match(/<a[^>]*>/gi) || []).length;
        const hrefCount = (clipboardContent?.match(/href\s*=\s*["'][^"']*["']/gi) || []).length;
        console.log(`Found ${anchorCount} anchor tags with ${hrefCount} href attributes in clipboard`);
        
        // Check if HTML is already clean (either from this extension or other sources)
        const isCleanHtml = looksLikeCleanHtml(clipboardContent);
        console.log('HTML is already clean:', isCleanHtml);
        
        // If HTML is already clean, skip all processing
        if (isCleanHtml) {
            console.log('Clean HTML detected, skipping all processing');
            const selection = textEditor.selection;
            await textEditor.edit(editBuilder => {
                editBuilder.replace(selection, clipboardContent);
            });
            return;
        }
        
        // Check if HTML needs cleaning (detect Word/Office artifacts)
        const needsCleaning = isWordHtml(clipboardContent);
        console.log('HTML needs cleaning:', needsCleaning);
        
        let finalHtml;
        if (needsCleaning) {
            // Save raw HTML to temp file for debugging
            const tempDebugPath = `/tmp/word-html-debug-${Date.now()}.html`;
            fs.writeFileSync(tempDebugPath, clipboardContent);
            console.log('Raw HTML saved to:', tempDebugPath);
            
            // Clean the HTML
            finalHtml = cleanWordHtml(clipboardContent);
            console.log('Cleaned HTML length:', finalHtml.length);
            console.log('Cleaned HTML (first 500 chars):', finalHtml.substring(0, 500));
        } else {
            // HTML may need other transformations but not Word cleaning
            console.log('HTML does not have Word markers, applying transformations');
            finalHtml = clipboardContent;
        }
        
        // Apply header tag detection (H1:, H2:, etc.)
        finalHtml = wrapHeaderPrefixes(finalHtml);
        console.log('After header wrapping (first 500 chars):', finalHtml.substring(0, 500));
        
        // Convert bullet points to proper lists
        finalHtml = convertBulletPointsToLists(finalHtml);
        console.log('After bullet list conversion (first 500 chars):', finalHtml.substring(0, 500));
        
        // Remove unwanted attributes (data-*, id, class, style, etc.)
        finalHtml = stripUnwantedAttributes(finalHtml);
        console.log('After stripping attributes (first 500 chars):', finalHtml.substring(0, 500));
        
        // Verify anchor links are still present
        const finalAnchorCount = (finalHtml?.match(/<a\s+href/gi) || []).length;
        console.log(`Final HTML contains ${finalAnchorCount} anchor links with href`);
        
        // Format with proper line breaks after block elements
        finalHtml = formatBlockElements(finalHtml);
        console.log('After formatting (first 500 chars):', finalHtml.substring(0, 500));
        
        // Final pass: normalize all remaining special whitespace characters
        finalHtml = normalizeWhitespace(finalHtml);
        console.log('After whitespace normalization (first 500 chars):', finalHtml.substring(0, 500));

        // Insert the final HTML
        const selection = textEditor.selection;
        await textEditor.edit(editBuilder => {
            editBuilder.replace(selection, finalHtml);
        });

    } catch (error) {
        console.error('Error pasting as HTML:', error);
        vscode.window.showErrorMessage(`Failed to paste as HTML: ${error.message}`);
        // Fall back to regular paste
        const clipboardContent = await vscode.env.clipboard.readText();
        if (clipboardContent) {
            const selection = textEditor.selection;
            await textEditor.edit(editBuilder => {
                editBuilder.replace(selection, clipboardContent);
            });
        }
    }
}

function normalizeWhitespace(html) {
    // Replace all special whitespace characters with regular spaces
    // This handles non-breaking spaces and other Unicode whitespace that editors flag
    let result = html;
    
    // Replace various Unicode whitespace characters with regular space
    result = result.replace(/\u00A0/g, ' ');   // Non-breaking space (most common)
    result = result.replace(/\u2007/g, ' ');   // Figure space
    result = result.replace(/\u202F/g, ' ');   // Narrow no-break space
    result = result.replace(/\u2009/g, ' ');   // Thin space
    result = result.replace(/\u200A/g, ' ');   // Hair space
    result = result.replace(/\u2003/g, ' ');   // Em space
    result = result.replace(/\u2002/g, ' ');   // En space
    
    // Remove zero-width characters completely
    result = result.replace(/\uFEFF/g, '');    // Zero-width no-break space (BOM)
    result = result.replace(/\u200B/g, '');    // Zero-width space
    result = result.replace(/\u200C/g, '');    // Zero-width non-joiner
    result = result.replace(/\u200D/g, '');    // Zero-width joiner
    
    // Clean up multiple consecutive spaces (but preserve single spaces)
    result = result.replace(/ {2,}/g, ' ');
    
    // Log if any special characters were found and replaced
    if (result !== html) {
        const nbspCount = (html.match(/\u00A0/g) || []).length;
        if (nbspCount > 0) {
            console.log(`Normalized ${nbspCount} non-breaking spaces and other special whitespace`);
        }
    }
    
    return result;
}

function formatBlockElements(html) {
    // Add newlines after block elements for readability
    // But keep inline elements (strong, em, a, etc.) on the same line
    let result = html;
    
    // Add newline after closing header tags
    result = result.replace(/<\/h[1-6]>/gi, '$&\n');
    
    // Add newline after closing paragraph tags
    result = result.replace(/<\/p>/gi, '$&\n');
    
    // Add newline after closing list tags
    result = result.replace(/<\/ul>/gi, '$&\n');
    result = result.replace(/<\/ol>/gi, '$&\n');
    
    // List items stay inline (no newline after </li>)
    
    // Clean up any double newlines
    result = result.replace(/\n\n+/g, '\n');
    
    return result.trim();
}

function stripUnwantedAttributes(html) {
    // Remove data-* attributes, id, class, style, and other unwanted attributes
    // Also remove meta tags and other document metadata
    // Keep only href for links
    
    let result = html;
    
    // Remove meta tags
    result = result.replace(/<meta[^>]*>/gi, '');
    result = result.replace(/<\/meta>/gi, '');
    
    // Remove link tags (stylesheets, etc.)
    result = result.replace(/<link[^>]*>/gi, '');
    result = result.replace(/<\/link>/gi, '');
    
    // Remove data-* attributes (like data-renderer-start-pos, data-id, etc.)
    result = result.replace(/\s+data-[a-z0-9-]+="[^"]*"/gi, '');
    result = result.replace(/\s+data-[a-z0-9-]+='[^']*'/gi, '');
    result = result.replace(/\s+data-[a-z0-9-]+=[^\s>]*/gi, '');
    
    // Remove class attributes
    result = result.replace(/\s+class="[^"]*"/gi, '');
    result = result.replace(/\s+class='[^']*'/gi, '');
    
    // Remove id attributes
    result = result.replace(/\s+id="[^"]*"/gi, '');
    result = result.replace(/\s+id='[^']*'/gi, '');
    
    // Remove style attributes
    result = result.replace(/\s+style="[^"]*"/gi, '');
    result = result.replace(/\s+style='[^']*'/gi, '');
    
    // Remove other common unwanted attributes (but preserve href on <a> tags)
    const unwantedAttrs = [
        'aria-[a-z0-9-]+',
        'role',
        'title',
        'rel',
        'target',
        'contenteditable',
        'spellcheck',
        'tabindex',
        'dir',
        'lang'
    ];
    
    unwantedAttrs.forEach(attr => {
        result = result.replace(new RegExp(`\\s+${attr}="[^"]*"`, 'gi'), '');
        result = result.replace(new RegExp(`\\s+${attr}='[^']*'`, 'gi'), '');
        result = result.replace(new RegExp(`\\s+${attr}=[^\\s>]*`, 'gi'), '');
    });
    
    // Ensure anchor tags with href are preserved - fix any broken hrefs
    // This ensures <a href="url"> format is maintained
    result = result.replace(/<a\s+href\s*=\s*["']([^"']+)["']\s*>/gi, '<a href="$1">');
    result = result.replace(/<a\s+href\s*=\s*([^\s>"'][^\s>]*)\s*>/gi, '<a href="$1">');
    
    // Clean up any double spaces left behind
    result = result.replace(/\s{2,}/g, ' ');
    
    // Log if changes were made
    if (result !== html) {
        console.log('Stripped unwanted attributes (data-*, class, id, style, etc.)');
    }
    
    return result;
}

function convertBulletPointsToLists(html) {
    // Convert paragraphs starting with bullets/dashes to proper <ul><li> lists
    let result = html;
    
    // Use regex to find all <p> tags with bullet points and group them into lists
    // This approach keeps everything inline without introducing line breaks
    
    // First, mark bullet points for processing
    result = result.replace(/<p>\s*(?:-|•|·)\s*(.+?)<\/p>\s*/gi, '##BULLET##$1##ENDBULLET##');
    
    // Now convert consecutive bullets into a list
    result = result.replace(/((?:##BULLET##.+?##ENDBULLET##)+)/g, (match) => {
        // Extract all bullet items
        const items = match.match(/##BULLET##(.+?)##ENDBULLET##/g);
        if (!items) return match;
        
        // Convert to list items
        const listItems = items.map(item => {
            const content = item.replace(/##BULLET##(.+?)##ENDBULLET##/, '$1').trim();
            return `<li>${content}</li>`;
        }).join('');
        
        return `<ul>${listItems}</ul>`;
    });
    
    if (result !== html) {
        console.log('Converted bullet points to proper <ul><li> lists');
    }
    
    return result;
}

function wrapHeaderPrefixes(html) {
    // Match paragraphs or plain text lines that start with H1:, H2:, etc.
    // Handles variations like "H2:", "H2 Title:", "H2 Tag:", "h2:", etc.
    
    let result = html;
    
    // Pattern 1: <p><strong>H2 Tag: Title text</strong></p>
    // First extract the H2 prefix and unwrap the strong tag
    result = result.replace(
        /<p>\s*<strong>\s*H([1-4])(?:\s+(?:Title|Tag))?\s*:\s*([^<]+)<\/strong>\s*<\/p>/gi,
        (match, level, text) => `<h${level}>${text.trim()}</h${level}>`
    );
    
    // Pattern 2: <p>H2: Title text</p> or <p>H2 Title: text</p> or <p>H2 Tag: text</p>
    result = result.replace(
        /<p>\s*H([1-4])(?:\s+(?:Title|Tag))?\s*:\s*([^<]+)<\/p>/gi,
        (match, level, text) => `<h${level}>${text.trim()}</h${level}>`
    );
    
    // Pattern 3: More complex - <p> with nested tags where title is in a tag
    // Example: <p><h3></h3><strong>Title</strong></p> (malformed)
    result = result.replace(
        /<p>\s*<h([1-4])>\s*<\/h\1>\s*<(?:strong|em|b|i|u)>([^<]+)<\/(?:strong|em|b|i|u)>\s*<\/p>/gi,
        (match, level, text) => `<h${level}>${text.trim()}</h${level}>`
    );
    
    // Pattern 4: Headers already wrapped in <p> tags
    // Example: <p><h2>Title</h2></p> → <h2>Title</h2>
    result = result.replace(
        /<p>\s*(<h[1-6]>[^<]*<\/h[1-6]>)\s*<\/p>/gi,
        '$1'
    );
    
    // Pattern 5: Bare text lines (no tags) like "H2: Title text" or "H2 Tag: text"
    result = result.replace(
        /^H([1-4])(?:\s+(?:Title|Tag))?\s*:\s*(.+)$/gim,
        (match, level, text) => `<h${level}>${text.trim()}</h${level}>`
    );
    
    // Log if any replacements were made
    if (result !== html) {
        console.log('Wrapped header prefixes (H1:, H2:, etc.) and cleaned up formatting');
    }
    
    return result;
}

function looksLikeCleanHtml(html) {
    // Check if HTML is already clean (minimal, semantic HTML without messy attributes)
    // This detects HTML that was already processed by this extension or is naturally clean
    
    console.log('Checking if HTML is clean...');
    console.log('HTML length:', html.length);
    console.log('First 300 chars:', html.substring(0, 300));
    
    // Check for signs of dirty HTML
    const dirtyMarkers = [
        // Word/Office markers
        'xmlns:w=',
        'xmlns:o=',
        'xmlns:m=',
        '<w:',
        '<o:',
        'class="Mso',
        'urn:schemas-microsoft-com',
        'StartFragment',
        'EndFragment',
        '/* Style Definitions */',
        
        // HTML document structure (indicates a full document, not a snippet)
        '<html',
        '<head',
        '<body',
        '<!DOCTYPE',
        
        // Messy attributes (typical of Word/CMS output)
        'style=',
        'class=',
        'id=',
        'data-',
        
        // Metadata tags
        '<meta',
        '<link',
        '<style',
        '<script'
    ];
    
    // If any dirty markers found, it's not clean
    for (const marker of dirtyMarkers) {
        if (html.includes(marker)) {
            console.log('Found dirty marker:', marker);
            return false;
        }
    }
    
    // Check if HTML only contains clean semantic tags
    // Clean tags: p, h1-h6, strong, em, ul, ol, li, a (with href only), br, span (if no attributes), div (if no attributes)
    const tagPattern = /<\/?([a-z][a-z0-9]*)[^>]*>/gi;
    const matches = Array.from(html.matchAll(tagPattern));
    const allowedTags = ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'strong', 'em', 'ul', 'ol', 'li', 'a', 'br', 'u', 'blockquote', 'pre', 'code', 'span', 'div'];
    
    console.log(`Found ${matches.length} HTML tags`);
    
    for (const match of matches) {
        const tagName = match[1].toLowerCase();
        if (!allowedTags.includes(tagName)) {
            console.log('Found non-semantic tag:', tagName);
            return false;
        }
    }
    
    // Check that anchor tags only have href attribute (no other attributes)
    const anchorPattern = /<a\s+([^>]*)>/gi;
    const anchorMatches = Array.from(html.matchAll(anchorPattern));
    
    for (const match of anchorMatches) {
        const attributes = match[1].trim();
        // Should only contain href="..." and nothing else
        if (!attributes.match(/^href="[^"]*"$/)) {
            console.log('Found anchor with non-href attributes:', attributes);
            return false;
        }
    }
    
    console.log('HTML appears to be clean (semantic tags only, minimal attributes)');
    return true;
}

function isWordHtml(html) {
    // Check for Word/Office-specific markers
    const wordMarkers = [
        'xmlns:w="urn:schemas-microsoft-com:office:word"',
        'xmlns:o="urn:schemas-microsoft-com:office:office"',
        'xmlns:m="http://schemas.microsoft.com/office',
        '<w:',
        '<o:',
        'class="Mso',
        'urn:schemas-microsoft-com',
        'StartFragment',
        'EndFragment',
        '<style>\n<!--',
        '/* Style Definitions */'
    ];
    
    // If any Word markers are found, it needs cleaning
    for (const marker of wordMarkers) {
        if (html.includes(marker)) {
            console.log('Found Word marker:', marker);
            return true;
        }
    }
    
    // Check if it's a complete HTML document (has <html>, <head>, <body>)
    // If it has these but no Word markers, it might still be messy HTML from Word
    if (html.includes('<html') && html.includes('<head') && html.includes('<body')) {
        // Check for excessive inline styles or class attributes (typical of Word)
        const styleCount = (html.match(/style="[^"]*"/g) || []).length;
        const classCount = (html.match(/class="[^"]*"/g) || []).length;
        
        // If there are many style/class attributes, it's probably from Word
        if (styleCount > 5 || classCount > 5) {
            console.log('Found excessive styling, likely Word HTML');
            return true;
        }
    }
    
    console.log('HTML appears to be clean');
    return false;
}

function cleanWordHtml(html) {
    try {
        console.log('Starting regex-based cleaning...');
        
        // Extract body content first
        let bodyContent = html;
        const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
        if (bodyMatch) {
            bodyContent = bodyMatch[1];
            console.log('Extracted body, length:', bodyContent.length);
        }
        
        console.log('Body content (first 500 chars):', bodyContent.substring(0, 500));
        
        // Aggressive cleaning with regex
        let cleanedHtml = bodyContent
            // Remove conditional comments first (Word uses these heavily)
            .replace(/<!--\[if[^\]]*\]>[\s\S]*?<!\[endif\]-->/gi, '')
            // Remove Word-specific CDATA-style conditional comments like <![if !supportLists]>
            .replace(/<!\[if[^\]]*\]>/gi, '')
            .replace(/<!\[endif\]>/gi, '')
            // Remove style blocks
            .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
            // Remove script blocks
            .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
            // Remove xml blocks
            .replace(/<xml[^>]*>[\s\S]*?<\/xml>/gi, '')
            // Remove Word namespace tags (w:, o:, v:, m:)
            .replace(/<\/?w:[^>]*>/gi, '')
            .replace(/<\/?o:[^>]*>/gi, '')
            .replace(/<\/?v:[^>]*>/gi, '')
            .replace(/<\/?m:[^>]*>/gi, '')
            // Remove HTML comments (including StartFragment/EndFragment)
            .replace(/<!--[^>]*-->/g, '')
            // Remove meta, link tags
            .replace(/<\/?meta[^>]*>/gi, '')
            .replace(/<\/?link[^>]*>/gi, '')
            // Remove div/span tags but keep their content
            .replace(/<div[^>]*>/gi, '')
            .replace(/<\/div>/gi, '')
            .replace(/<span[^>]*>/gi, '')
            .replace(/<\/span>/gi, '')
            // Remove font tags
            .replace(/<\/?font[^>]*>/gi, '')
            // Convert b to strong, i to em
            .replace(/<b(\s[^>]*)?>([\s\S]*?)<\/b>/gi, '<strong>$2</strong>')
            .replace(/<i(\s[^>]*)?>([\s\S]*?)<\/i>/gi, '<em>$2</em>')
            // Clean attributes from tags (but preserve href on links)
            .replace(/<(p|h[1-6]|ul|ol|li|strong|em|u|br|table|thead|tbody|tr|th|td|blockquote|pre|code)\s+[^>]*?>/gi, '<$1>')
            // Handle links - preserve href and clean other attributes
            // Match <a> tags with href attribute (with or without other attributes)
            .replace(/<a\s+[^>]*?href\s*=\s*["']([^"']+)["'][^>]*?>/gi, '<a href="$1">')
            // Also handle href without quotes (less common but valid HTML)
            .replace(/<a\s+[^>]*?href\s*=\s*([^\s>"'][^\s>]*)[^>]*?>/gi, '<a href="$1">')
            // Clean up entities and special whitespace characters
            .replace(/&nbsp;/g, ' ')  // HTML entity
            .replace(/\u00A0/g, ' ')   // Non-breaking space (Unicode U+00A0)
            .replace(/\u2007/g, ' ')   // Figure space
            .replace(/\u202F/g, ' ')   // Narrow no-break space
            .replace(/\u2009/g, ' ')   // Thin space
            .replace(/\u200A/g, ' ')   // Hair space
            .replace(/\uFEFF/g, '')    // Zero-width no-break space (BOM)
            .replace(/\u200B/g, '')    // Zero-width space
            .replace(/\u200C/g, '')    // Zero-width non-joiner
            .replace(/\u200D/g, '')    // Zero-width joiner
            // Remove empty paragraphs and headings (multiple passes)
            .replace(/<(p|h[1-6]|li|strong|em|u|th|td)>\s*<\/\1>/gi, '')
            .replace(/<(p|h[1-6]|li|strong|em|u|th|td)>\s*<\/\1>/gi, '')
            .replace(/<(p|h[1-6]|li|strong|em|u|th|td)>\s*<\/\1>/gi, '')
            // Normalize whitespace inside tags (but preserve structure between tags)
            .replace(/[ \t\u00A0]+/g, ' ')
            // Remove all newlines first to clean slate
            .replace(/\n+/g, ' ')
            // Add space after inline closing tags to prevent words running together
            .replace(/<\/(strong|em|u|a)>/gi, '</$1> ')
            // Remove space before punctuation that follows inline tags
            .replace(/\s+([.,;:!?])/g, '$1')
            // Clean up multiple spaces
            .replace(/  +/g, ' ')
            .trim();
        
        console.log('After regex cleaning, length:', cleanedHtml.length);
        console.log('After regex cleaning (first 500 chars):', cleanedHtml.substring(0, 500));
        
        return cleanedHtml;

    } catch (error) {
        console.error('Error cleaning HTML:', error);
        return html; // Return original on error
    }
}

function deactivate() {
    if (statusBarItem) {
        statusBarItem.dispose();
    }
}

module.exports = {
    activate,
    deactivate
};
