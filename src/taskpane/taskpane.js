import { processText } from '../lib/llm-providers.js';
import { promptManager } from '../lib/prompt-manager.js';
import { storage } from '../lib/storage.js';

// DOM Elements
let elements = {};
let currentEmailBody = '';
let currentEmailBodyHtml = '';
let currentThreadContent = '';
let initialBodyTemplate = null; // Captured on first load to identify signature
let currentResult = { subject: '', body: '' };
let refreshInterval = null;

/**
 * Initialize the add-in when Office is ready
 */
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initializeApp();
    }
});

/**
 * Initialize the application
 */
function initializeApp() {
    // Cache DOM elements
    elements = {
        statusIndicator: document.getElementById('statusIndicator'),
        statusText: document.getElementById('statusText'),
        contextPreviewContent: document.getElementById('contextPreviewContent'),
        savedPrompts: document.getElementById('savedPrompts'),
        instruction: document.getElementById('instruction'),
        includeThreadToggle: document.getElementById('includeThreadToggle'),
        processBtn: document.getElementById('processBtn'),
        resultSection: document.getElementById('resultSection'),
        subjectText: document.getElementById('subjectText'),
        insertSubjectBtn: document.getElementById('insertSubjectBtn'),
        bodyPreview: document.getElementById('bodyPreview'),
        copyToClipboardBtn: document.getElementById('copyToClipboardBtn'),
        replaceBodyBtn: document.getElementById('replaceBodyBtn'),
        settingsBtn: document.getElementById('settingsBtn'),
        errorMessage: document.getElementById('errorMessage'),
        providerIndicator: document.getElementById('providerIndicator')
    };

    // Initialize default prompts
    promptManager.initDefaults();

    // Load saved prompts into dropdown
    loadSavedPrompts();

    // Update provider indicator
    updateProviderIndicator();

    // Set up event listeners
    setupEventListeners();

    // Start auto-capture
    captureEmailBody();
    startAutoRefresh();
}

/**
 * Set up all event listeners
 */
function setupEventListeners() {
    elements.savedPrompts.addEventListener('change', onSavedPromptChange);
    elements.instruction.addEventListener('input', updateProcessButtonState);
    elements.processBtn.addEventListener('click', handleProcess);
    elements.insertSubjectBtn.addEventListener('click', handleInsertSubject);
    elements.replaceBodyBtn.addEventListener('click', handleReplaceBody);
    elements.copyToClipboardBtn.addEventListener('click', handleCopyToClipboard);
    elements.settingsBtn.addEventListener('click', openSettings);

    // Toggle context preview
    elements.statusIndicator.addEventListener('click', toggleContextPreview);

    // Refresh status when thread toggle changes
    elements.includeThreadToggle.addEventListener('change', () => {
        captureEmailBody();
    });
}

/**
 * Start auto-refresh interval
 */
function startAutoRefresh() {
    // Refresh every 1 second
    refreshInterval = setInterval(() => {
        captureEmailBody();
    }, 1000);
}

/**
 * Capture the email body, separating current message from thread content
 * Uses HTML parsing to reliably identify boundaries
 */
function captureEmailBody() {
    // Get HTML version for proper parsing
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        (htmlResult) => {
            if (htmlResult.status === Office.AsyncResultStatus.Succeeded) {
                const fullHtml = htmlResult.value;
                currentEmailBodyHtml = fullHtml;

                // Parse HTML to separate current message from thread
                const parsed = parseEmailHtml(fullHtml);

                // On first run, store the initial message part as the signature template
                // This is very effective because Outlook inserts the signature before the user starts typing
                if (initialBodyTemplate === null) {
                    initialBodyTemplate = parsed.currentMessage.trim();
                }

                // Current body is the message content (stripped of signature)
                // We pass the template for smart comparison
                currentEmailBody = stripSignature(parsed.currentMessage, initialBodyTemplate);

                // Thread content is cleaned previous emails (when toggle is on)
                if (elements.includeThreadToggle.checked) {
                    currentThreadContent = cleanThreadContent(parsed.threadContent);
                } else {
                    currentThreadContent = '';
                }

                updateStatusIndicator(currentEmailBody.length > 0);
                updateProcessButtonState();
            }
        }
    );
}

/**
 * Parse email HTML to separate current message from thread content
 * @param {string} html - Full email HTML
 * @returns {{currentMessage: string, threadContent: string}}
 */
function parseEmailHtml(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;

    // Find thread boundary markers (Outlook uses divRplyFwdMsg, Gmail uses gmail_quote)
    const threadMarkers = [
        '#divRplyFwdMsg',           // Outlook's reply/forward marker
        '#x_divRplyFwdMsg',         // Sometimes prefixed with x_
        '.gmail_quote',             // Gmail's quote marker
        '#appendonsend',            // Outlook appendonsend marker
        'hr[tabindex="-1"]',        // Common Outlook separator
        '.ms-MessageBody',          // Outlook Web separator
        'blockquote',               // Standard blockquote
    ];

    let threadElement = null;
    for (const selector of threadMarkers) {
        threadElement = temp.querySelector(selector);
        if (threadElement) break;
    }

    // Also check for common patterns without specific IDs
    if (!threadElement) {
        // Look for elements containing "From:" followed by "Sent:" pattern
        const allElements = temp.querySelectorAll('div, p, font');
        for (const el of allElements) {
            const text = el.textContent || '';
            // More relaxed check for the start of a thread
            if (/From:.*Sent:/i.test(text) ||
                /De\s*:.*Envoyé/i.test(text) ||
                /Von:.*Gesendet/i.test(text) ||
                /---------- Original Message ----------/i.test(text)) {
                threadElement = el;
                break;
            }
        }
    }

    // If still no thread element, check for horizontal rules as separators
    if (!threadElement) {
        const hr = temp.querySelector('hr');
        if (hr) threadElement = hr;
    }

    let currentMessage = '';
    let threadContent = '';

    if (threadElement) {
        // Get content before thread marker
        const beforeThread = document.createElement('div');
        let node = temp.firstChild;

        // Clone nodes before thread element
        while (node && !node.contains(threadElement) && node !== threadElement) {
            beforeThread.appendChild(node.cloneNode(true));
            node = node.nextSibling;
        }

        // IMPORTANT: Capture the thread element AND all subsequent siblings
        // This ensures the body text following the "From:" block is included
        const fullThread = document.createElement('div');
        let threadNode = node; // Current node is the thread marker or contains it
        while (threadNode) {
            fullThread.appendChild(threadNode.cloneNode(true));
            threadNode = threadNode.nextSibling;
        }

        currentMessage = beforeThread.textContent || beforeThread.innerText || '';
        threadContent = fullThread.textContent || fullThread.innerText || '';
    } else {
        // No thread marker found - entire content is current message
        currentMessage = temp.textContent || temp.innerText || '';
        threadContent = '';
    }

    return {
        currentMessage: currentMessage.trim(),
        threadContent: threadContent.trim()
    };
}

/**
 * Strip signature from message text
 * @param {string} text - Message text
 * @param {string} template - Initial body captured at startup (optional)
 * @returns {string} - Text with signature removed
 */
function stripSignature(text, template) {
    let result = text.trim();

    // Strategy 1: Smart Comparison
    // If the message ends with our initial template (the signature), strip it
    if (template && template.length > 5) {
        if (result.endsWith(template)) {
            return result.substring(0, result.length - template.length).trim();
        }
    }

    // Strategy 2: Phrase-based signature detection
    const signaturePatterns = [
        /\n--\s*\n[\s\S]*/,
        /\nBest regards,[\s\S]*/i,
        /\nKind regards,[\s\S]*/i,
        /\nRegards,[\s\S]*/i,
        /\nSincerely,[\s\S]*/i,
        /\nThanks,[\s\S]*/i,
        /\nThank you,[\s\S]*/i,
        /\nCheers,[\s\S]*/i,
        /\nCordialement,[\s\S]*/i,
        /\nCdlt,[\s\S]*/i,
        /\nMit freundlichen Grüßen,[\s\S]*/i,
        /\nMfG,[\s\S]*/i,
        /\nSent from my iPhone[\s\S]*/i,
        /\nSent from my Android[\s\S]*/i,
        /\nGet Outlook for[\s\S]*/i,
        /\nEnvoyé de mon[\s\S]*/i,
        /\nVerzonden vanaf[\s\S]*/i,
    ];

    for (const pattern of signaturePatterns) {
        if (pattern.test(result)) {
            result = result.replace(pattern, '');
            // Optimization: if we hit a standard phrase, we likely found it
            return result.trim();
        }
    }

    // Strategy 3: Structural fallback (looking for contact blocks)
    // If we see phone numbers, websites, or multiple short lines with links near the end
    const lines = result.split('\n');
    if (lines.length > 3) {
        // Work backwards from the end
        for (let i = lines.length - 1; i >= Math.max(0, lines.length - 8); i--) {
            const line = lines[i].trim();
            // Patterns for common signature elements: phone, email, url
            const hasPhone = /\+?[\d\s\-\(\).]{8,20}/.test(line);
            const hasEmail = /[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i.test(line);
            const hasUrl = /www\.|http/i.test(line);

            if (hasPhone || hasEmail || hasUrl) {
                // If we found contact info, consider this and everything below as signature
                // (assuming it's near the very end)
                if (i > lines.length - 10) {
                    return lines.slice(0, i).join('\n').trim();
                }
            }
        }
    }

    return result;
}

/**
 * Clean thread content by removing email headers and metadata
 * Keeps only the actual message body text
 * @param {string} text - Raw thread text
 * @returns {string} - Cleaned thread content
 */
function cleanThreadContent(text) {
    if (!text) return '';

    let result = text;

    // 1. Range-based removal (User requested: remove content BETWEEN markers)
    // We use [\s\S]*? to lazily match characters including newlines
    const rangePatterns = [
        // From: [any] Sent:
        { pattern: /From:[\s\S]*?(?=Sent:)/gi, replacement: "From:\n" },
        // Sent: [any] To:
        { pattern: /Sent:[\s\S]*?(?=To:)/gi, replacement: "Sent:\n" },
        // To: [any] Subject:
        { pattern: /To:[\s\S]*?(?=Subject:)/gi, replacement: "To:\n" },

        // French variants
        { pattern: /De\s*:[\s\S]*?(?=Envoyé)/gi, replacement: "De:\n" },
        { pattern: /Envoyé\s*:[\s\S]*?(?=À)/gi, replacement: "Envoyé:\n" },
        { pattern: /À\s*:[\s\S]*?(?=Objet)/gi, replacement: "À:\n" },

        // German variants
        { pattern: /Von:[\s\S]*?(?=Gesendet)/gi, replacement: "Von:\n" },
        { pattern: /Gesendet:[\s\S]*?(?=An)/gi, replacement: "Gesendet:\n" },
        { pattern: /An:[\s\S]*?(?=Betreff)/gi, replacement: "An:\n" }
    ];

    for (const { pattern, replacement } of rangePatterns) {
        result = result.replace(pattern, replacement);
    }

    // 2. Remove other individual email header lines
    const headerPatterns = [
        /^Cc:.*$/gm,
        /^Bcc:.*$/gm,
        /^Date:.*$/gm,
        /^Importance:.*$/gm,
        // Email addresses in angle brackets (global cleanup)
        /<[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}>/g,
        // "On [date], [person] wrote:" lines
        /^On\s+.*\s+wrote:\s*$/gim,
        /^Le\s+.*\s+a écrit\s*:\s*$/gim,
        /^Am\s+.*\s+schrieb\s*:\s*$/gim,
        // Separator lines
        /^[-_]{3,}.*$/gm,
        /^Original Message.*$/gim,
        /^Message d'origine.*$/gim,
        /^Forwarded message.*$/gim,
    ];

    for (const pattern of headerPatterns) {
        result = result.replace(pattern, '');
    }

    // 3. Clean up multiple blank lines
    result = result.replace(/\n{3,}/g, '\n\n');

    // 4. Also strip signatures from thread content
    result = stripSignature(result);

    return result.trim();
}

/**
 * Update status indicator
 * @param {boolean} hasContent 
 */
function updateStatusIndicator(hasContent) {
    const dot = elements.statusIndicator.querySelector('.status-dot');

    if (hasContent) {
        dot.classList.add('active');

        // Calculate total chars including thread if toggle is on
        const bodyChars = currentEmailBody.length;
        const threadChars = elements.includeThreadToggle.checked ? currentThreadContent.length : 0;
        const totalChars = bodyChars + threadChars;

        let statusMsg = '';
        if (threadChars > 0) {
            statusMsg = `Monitoring (${bodyChars} + ${threadChars} thread = ${totalChars} chars)`;
        } else {
            statusMsg = `Monitoring email (${bodyChars} chars)`;
        }
        elements.statusText.textContent = statusMsg;

        // Update preview content
        let previewText = currentEmailBody;
        if (elements.includeThreadToggle.checked && currentThreadContent) {
            previewText += '\n\n--- PREVIOUS THREAD ---\n\n' + currentThreadContent;
        }
        elements.contextPreviewContent.textContent = previewText || '(No content captured yet)';

    } else {
        dot.classList.remove('active');
        elements.statusText.textContent = 'Start typing your email...';
        elements.contextPreviewContent.textContent = '(Waiting for content...)';
    }
}

/**
 * Toggle context preview visibility
 */
function toggleContextPreview() {
    const container = document.querySelector('.status-container');
    container.classList.toggle('expanded');
}

/**
 * Load saved prompts into the dropdown
 */
function loadSavedPrompts() {
    const prompts = promptManager.getAll();

    // Clear existing options except the first one
    while (elements.savedPrompts.options.length > 1) {
        elements.savedPrompts.remove(1);
    }

    // Add prompts
    prompts.forEach(prompt => {
        const option = document.createElement('option');
        option.value = prompt.id;
        option.textContent = prompt.name;
        elements.savedPrompts.appendChild(option);
    });
}

/**
 * Handle saved prompt selection
 */
function onSavedPromptChange() {
    const selectedId = elements.savedPrompts.value;

    if (selectedId) {
        const prompt = promptManager.getById(selectedId);
        if (prompt) {
            elements.instruction.value = prompt.instruction;
            updateProcessButtonState();
        }
    }
}

/**
 * Update the process button state
 */
function updateProcessButtonState() {
    const hasBody = currentEmailBody.trim().length > 0;
    const hasInstruction = elements.instruction.value.trim().length > 0;
    elements.processBtn.disabled = !(hasBody && hasInstruction);
}

/**
 * Handle the process button click
 */
async function handleProcess() {
    hideError();
    setLoading(true);

    try {
        const instruction = elements.instruction.value.trim();

        // Build context
        let context = currentEmailBody;
        if (elements.includeThreadToggle.checked && currentThreadContent) {
            context += '\n\n--- Previous Thread ---\n' + currentThreadContent;
        }

        // Call LLM with structured output request
        const result = await processText(context, instruction);

        // Parse the result (expecting JSON with subject and body)
        try {
            currentResult = JSON.parse(result);
        } catch (e) {
            // If not JSON, treat entire result as body
            currentResult = {
                subject: '',
                body: result
            };
        }

        // Display results
        displayResults();

    } catch (error) {
        showError(error.message);
    } finally {
        setLoading(false);
    }
}

/**
 * Display the AI results
 */
function displayResults() {
    // Subject
    if (currentResult.subject) {
        elements.subjectText.textContent = currentResult.subject;
        document.getElementById('subjectCard').style.display = 'block';
    } else {
        document.getElementById('subjectCard').style.display = 'none';
    }

    // Body - render as HTML if it contains tags, otherwise as text
    if (currentResult.body) {
        if (currentResult.body.includes('<')) {
            elements.bodyPreview.innerHTML = currentResult.body;
        } else {
            elements.bodyPreview.textContent = currentResult.body;
        }
        document.getElementById('bodyCard').style.display = 'block';
    } else {
        document.getElementById('bodyCard').style.display = 'none';
    }

    elements.resultSection.style.display = 'block';
    elements.resultSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * Insert the suggested subject
 */
function handleInsertSubject() {
    if (!currentResult.subject) return;

    Office.context.mailbox.item.subject.setAsync(
        currentResult.subject,
        (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showSuccess('Subject inserted!');
            } else {
                showError('Failed to insert subject: ' + result.error.message);
            }
        }
    );
}

/**
 * Copy body to clipboard with rich text formatting
 * Uses legacy execCommand fallback for Office Add-in iframe compatibility
 */
async function handleCopyToClipboard() {
    if (!currentResult.body) return;

    const bodyHtml = formatBodyForInsert(currentResult.body);
    const plainText = currentResult.body.replace(/<[^>]*>/g, '');

    // Try modern Clipboard API first
    try {
        const htmlBlob = new Blob([bodyHtml], { type: 'text/html' });
        const textBlob = new Blob([plainText], { type: 'text/plain' });

        await navigator.clipboard.write([
            new ClipboardItem({
                'text/html': htmlBlob,
                'text/plain': textBlob
            })
        ]);

        showSuccess('Copied to clipboard!');
        return;
    } catch (err) {
        // Modern API blocked, move to legacy fallback
    }

    // Legacy fallback using execCommand (works in Office Add-in iframes)
    // This method supports rich text by copying from a visible/hidden DOM element
    try {
        // Create a temporary container for rich text
        const container = document.createElement('div');
        container.innerHTML = bodyHtml;

        // Ensure it's not visible but is in the DOM
        container.style.position = 'fixed';
        container.style.left = '-9999px';
        container.style.top = '0';
        container.style.whiteSpace = 'pre-wrap'; // Preserve some formatting

        document.body.appendChild(container);

        // Select the content
        const range = document.createRange();
        range.selectNodeContents(container);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);

        // Execute copy
        const success = document.execCommand('copy');

        // Cleanup
        selection.removeAllRanges();
        document.body.removeChild(container);

        if (success) {
            showSuccess('Copied to clipboard!');
        } else {
            // Last resort: simple plain text copy with textarea
            copyPlainTextFallback(plainText);
        }
    } catch (legacyErr) {
        console.error('Legacy copy failed:', legacyErr);
        copyPlainTextFallback(plainText);
    }
}

/**
 * Last resort plain text copy
 */
function copyPlainTextFallback(text) {
    try {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-9999px';
        document.body.appendChild(textArea);
        textArea.select();
        const success = document.execCommand('copy');
        document.body.removeChild(textArea);

        if (success) {
            showSuccess('Copied as plain text');
        } else {
            showError('Copy failed. Please select text manually.');
        }
    } catch (err) {
        showError('Copy not supported.');
    }
}

/**
 * Replace only the draft text above the signature, preserving signature and thread
 * Uses smart HTML parsing to identify and preserve email structure
 */
function handleReplaceBody() {
    if (!currentResult.body) return;

    const newBodyContent = formatBodyForInsert(currentResult.body);

    // Get the current full HTML to preserve signature and thread
    Office.context.mailbox.item.body.getAsync(
        Office.CoercionType.Html,
        (htmlResult) => {
            if (htmlResult.status !== Office.AsyncResultStatus.Succeeded) {
                showError('Failed to read current email body');
                return;
            }

            const fullHtml = htmlResult.value;

            // Extract signature and thread portions to preserve them
            const preserved = extractPreservedContent(fullHtml);

            // Reconstruct the email: new AI content + signature + thread
            let reconstructedBody = newBodyContent;

            if (preserved.signatureHtml) {
                reconstructedBody += preserved.signatureHtml;
            }

            if (preserved.threadHtml) {
                reconstructedBody += preserved.threadHtml;
            }

            // Set the reconstructed body
            Office.context.mailbox.item.body.setAsync(
                reconstructedBody,
                { coercionType: Office.CoercionType.Html },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        showSuccess('Body replaced!');
                        hideResults();
                    } else {
                        showError('Failed to replace: ' + result.error.message);
                    }
                }
            );
        }
    );
}

/**
 * Extract signature and thread HTML from the full email body
 * @param {string} html - Full email HTML
 * @returns {{signatureHtml: string, threadHtml: string}}
 */
function extractPreservedContent(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;

    let signatureHtml = '';
    let threadHtml = '';

    // === STEP 1: Find and extract thread content ===
    const threadMarkers = [
        '#divRplyFwdMsg',           // Outlook's reply/forward marker
        '#x_divRplyFwdMsg',         // Sometimes prefixed with x_
        '.gmail_quote',             // Gmail's quote marker
        '#appendonsend',            // Outlook appendonsend marker
        'hr[tabindex="-1"]',        // Common Outlook separator
        '.ms-MessageBody',          // Outlook Web separator
        'blockquote',               // Standard blockquote
    ];

    let threadElement = null;
    for (const selector of threadMarkers) {
        threadElement = temp.querySelector(selector);
        if (threadElement) break;
    }

    // Also check for "From: ... Sent:" pattern
    if (!threadElement) {
        const allElements = temp.querySelectorAll('div, p, font');
        for (const el of allElements) {
            const text = el.textContent || '';
            if (/From:.*Sent:/i.test(text) ||
                /De\s*:.*Envoyé/i.test(text) ||
                /Von:.*Gesendet/i.test(text) ||
                /---------- Original Message ----------/i.test(text)) {
                threadElement = el;
                break;
            }
        }
    }

    // If we found a thread element, capture it and all siblings after it
    if (threadElement) {
        const threadContainer = document.createElement('div');

        // Find the actual node in DOM tree (could be the element or its parent)
        let startNode = threadElement;

        // Walk up if this is a deep text match to get the containing block
        while (startNode.parentElement && startNode.parentElement !== temp) {
            // Check if parent is a direct child of temp
            if (startNode.parentElement.parentElement === temp) {
                startNode = startNode.parentElement;
                break;
            }
            startNode = startNode.parentElement;
        }

        // Clone this node and all following siblings
        let node = startNode;
        while (node) {
            threadContainer.appendChild(node.cloneNode(true));
            node = node.nextSibling;
        }

        threadHtml = threadContainer.innerHTML;

        // Remove thread content from temp so we can find signature in remaining content
        node = startNode;
        while (node) {
            const next = node.nextSibling;
            node.remove();
            node = next;
        }
    }

    // === STEP 2: Find and extract signature ===
    // Look for common signature markers in the remaining content
    const signatureMarkers = [
        '#Signature',               // Outlook signature ID
        '.Signature',               // Outlook signature class
        '#ms-outlook-mobile-signature',
        '[data-signature]',
        // Look for elements with typical signature styling
    ];

    let signatureElement = null;
    for (const selector of signatureMarkers) {
        signatureElement = temp.querySelector(selector);
        if (signatureElement) break;
    }

    // If no explicit signature found, try to detect signature patterns
    // Use the initial template if we captured it at startup
    if (!signatureElement && initialBodyTemplate && initialBodyTemplate.length > 5) {
        // The signature is likely the content that matches our initial template
        // Look for elements at the end that contain signature-like content
        const allElements = Array.from(temp.querySelectorAll('div, p, table'));

        // Check from the end of the document backwards
        for (let i = allElements.length - 1; i >= 0; i--) {
            const el = allElements[i];
            const elText = (el.textContent || '').trim();

            // Check if this element's text is part of the initial template (signature)
            if (elText.length > 5 && initialBodyTemplate.includes(elText)) {
                signatureElement = el;
                break;
            }
        }
    }

    // Also check for signature phrases as fallback
    if (!signatureElement) {
        const signaturePhrases = [
            /^Best regards/im,
            /^Kind regards/im,
            /^Regards/im,
            /^Sincerely/im,
            /^Thanks/im,
            /^Thank you/im,
            /^Cheers/im,
            /^Cordialement/im,
            /^Cdlt/im,
            /^Mit freundlichen Grüßen/im,
            /^MfG/im,
            /^--\s*$/m,
        ];

        const allElements = temp.querySelectorAll('div, p');
        for (const el of allElements) {
            const text = el.textContent || '';
            for (const phrase of signaturePhrases) {
                if (phrase.test(text)) {
                    signatureElement = el;
                    break;
                }
            }
            if (signatureElement) break;
        }
    }

    // If we found a signature, extract it and everything after
    if (signatureElement) {
        const sigContainer = document.createElement('div');

        // Walk up to get the block-level ancestor that's a direct child of temp
        let startNode = signatureElement;
        while (startNode.parentElement && startNode.parentElement !== temp) {
            if (startNode.parentElement.parentElement === temp) {
                startNode = startNode.parentElement;
                break;
            }
            startNode = startNode.parentElement;
        }

        // Clone from signature start to end
        let node = startNode;
        while (node) {
            sigContainer.appendChild(node.cloneNode(true));
            node = node.nextSibling;
        }

        signatureHtml = sigContainer.innerHTML;
    }

    return {
        signatureHtml: signatureHtml,
        threadHtml: threadHtml
    };
}

/**
 * Add inline styles for Outlook
 * Outlook often ignores external CSS and defaults, so we enforce spacing
 * @param {string} html 
 * @returns {string}
 */
function addOutlookStyles(html) {
    if (!html) return html;

    // Inject margin styles into paragraph tags
    let styledHtml = html.replace(/<p>/gi, '<p style="margin-top: 0; margin-bottom: 15px;">');

    // Inject margin styles into lists
    styledHtml = styledHtml.replace(/<ul>/gi, '<ul style="margin-bottom: 15px;">');
    styledHtml = styledHtml.replace(/<ol>/gi, '<ol style="margin-bottom: 15px;">');

    return styledHtml;
}

/**
 * Format body content for insertion
 * Handles both HTML and markdown/plain text conversion
 * @param {string} body 
 * @returns {string}
 */
function formatBodyForInsert(body) {
    // If already contains HTML tags, return with styles injected
    if (/<[a-z][\s\S]*>/i.test(body)) {
        return addOutlookStyles(body);
    }

    // Convert markdown/plain text to HTML
    let html = body;

    // Convert **bold** to <strong>
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');

    // Convert *italic* to <em> (but not if it's a list marker at start of line)
    html = html.replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '<em>$1</em>');

    // Convert __bold__ to <strong>
    html = html.replace(/__(.+?)__/g, '<strong>$1</strong>');

    // Convert _italic_ to <em>
    html = html.replace(/(?<!_)_(?!_)(.+?)(?<!_)_(?!_)/g, '<em>$1</em>');

    // Process lines for lists and paragraphs
    const lines = html.split('\n');
    const processedLines = [];
    let inUnorderedList = false;
    let inOrderedList = false;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // Check for unordered list items (- or *)
        const unorderedMatch = line.match(/^[-*]\s+(.+)$/);
        // Check for ordered list items (1. 2. etc.)
        const orderedMatch = line.match(/^\d+\.\s+(.+)$/);

        if (unorderedMatch) {
            if (inOrderedList) {
                processedLines.push('</ol>');
                inOrderedList = false;
            }
            if (!inUnorderedList) {
                processedLines.push('<ul>');
                inUnorderedList = true;
            }
            processedLines.push(`<li>${unorderedMatch[1]}</li>`);
        } else if (orderedMatch) {
            if (inUnorderedList) {
                processedLines.push('</ul>');
                inUnorderedList = false;
            }
            if (!inOrderedList) {
                processedLines.push('<ol>');
                inOrderedList = true;
            }
            processedLines.push(`<li>${orderedMatch[1]}</li>`);
        } else {
            // Close any open lists
            if (inUnorderedList) {
                processedLines.push('</ul>');
                inUnorderedList = false;
            }
            if (inOrderedList) {
                processedLines.push('</ol>');
                inOrderedList = false;
            }

            // Handle empty lines as paragraph breaks
            if (line === '') {
                if (processedLines.length > 0 && !processedLines[processedLines.length - 1].endsWith('</p>')) {
                    processedLines.push('</p><p>');
                }
            } else {
                // Regular text line
                if (processedLines.length === 0 || processedLines[processedLines.length - 1].endsWith('</p>') ||
                    processedLines[processedLines.length - 1].endsWith('</ul>') ||
                    processedLines[processedLines.length - 1].endsWith('</ol>')) {
                    processedLines.push(`<p>${line}`);
                } else if (processedLines[processedLines.length - 1] === '</p><p>') {
                    processedLines[processedLines.length - 1] = `</p><p>${line}`;
                } else {
                    // Regular text line - treat as new paragraph for proper spacing
                    // Close previous open paragraph
                    if (processedLines.length > 0 && !processedLines[processedLines.length - 1].endsWith('</p>') &&
                        !processedLines[processedLines.length - 1].endsWith('</ul>') &&
                        !processedLines[processedLines.length - 1].endsWith('</ol>')) {
                        processedLines[processedLines.length - 1] += '</p>';
                    }
                    processedLines.push(`<p>${line}`);
                }
            }
        }
    }

    // Close any remaining open tags
    if (inUnorderedList) processedLines.push('</ul>');
    if (inOrderedList) processedLines.push('</ol>');

    // Ensure we close the last paragraph
    const result = processedLines.join('');
    if (result.includes('<p>') && !result.endsWith('</p>') && !result.endsWith('</ul>') && !result.endsWith('</ol>')) {
        return result + '</p>';
    }

    // Handle case where there's no paragraph structure
    if (!result.includes('<p>') && !result.includes('<ul>') && !result.includes('ol>')) {
        return addOutlookStyles(`<p>${result}</p>`);
    }

    return addOutlookStyles(result);
}

/**
 * Hide results section
 */
function hideResults() {
    elements.resultSection.style.display = 'none';
    elements.instruction.value = '';
    elements.savedPrompts.value = '';
    currentResult = { subject: '', body: '' };
    updateProcessButtonState();
}

/**
 * Open settings page
 */
function openSettings() {
    window.location.href = '../settings/settings.html';
}

/**
 * Update the provider indicator in the footer
 */
function updateProviderIndicator() {
    const settings = storage.getProviderSettings();
    const provider = settings.activeProvider;
    const providerNames = {
        openai: 'OpenAI',
        claude: 'Claude',
        gemini: 'Gemini',
        custom: 'Custom API'
    };

    const hasApiKey = settings.providers[provider]?.apiKey;
    const displayName = hasApiKey ? providerNames[provider] : 'Not configured';

    elements.providerIndicator.innerHTML = `Provider: <strong>${displayName}</strong>`;
}

/**
 * Show loading state
 * @param {boolean} loading 
 */
function setLoading(loading) {
    const btnText = elements.processBtn.querySelector('.btn-text');
    const btnLoading = elements.processBtn.querySelector('.btn-loading');

    if (loading) {
        btnText.style.display = 'none';
        btnLoading.style.display = 'inline-flex';
        elements.processBtn.disabled = true;
    } else {
        btnText.style.display = 'inline';
        btnLoading.style.display = 'none';
        updateProcessButtonState();
    }
}

/**
 * Show error message
 * @param {string} message 
 */
function showError(message) {
    elements.errorMessage.textContent = message;
    elements.errorMessage.style.display = 'block';
}

/**
 * Hide error message
 */
function hideError() {
    elements.errorMessage.style.display = 'none';
}

/**
 * Show success feedback
 * @param {string} message 
 */
function showSuccess(message) {
    // Brief visual feedback
    const el = document.createElement('div');
    el.className = 'success-toast';
    el.textContent = message;
    document.body.appendChild(el);

    setTimeout(() => {
        el.remove();
    }, 2000);
}
