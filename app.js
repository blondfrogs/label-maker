// Label templates configuration - All measurements in inches
// Format: { width, height, topMargin, sideMargin, hPitch (horizontal spacing), vPitch (vertical spacing), cols, rows }

// US Templates (Letter paper: 8.5" x 11")
const US_TEMPLATES = {
    '8160': { width: 2.625, height: 1.0, topMargin: 0.5, sideMargin: 0.19, hPitch: 2.75, vPitch: 1.0, cols: 3, rows: 10, name: 'Avery 8160/5160 - Standard Address', desc: '1" × 2⅝", 30 per sheet', popular: true },
    '5160': { width: 2.625, height: 1.0, topMargin: 0.5, sideMargin: 0.19, hPitch: 2.75, vPitch: 1.0, cols: 3, rows: 10, name: 'Avery 5160 - Standard Address', desc: '1" × 2⅝", 30 per sheet' },
    '5162': { width: 4.0, height: 1.33, topMargin: 0.5, sideMargin: 0.16, hPitch: 4.19, vPitch: 1.33, cols: 2, rows: 7, name: 'Avery 5162/8162 - Wide Address', desc: '1⅓" × 4", 14 per sheet' },
    '5167': { width: 1.75, height: 0.5, topMargin: 0.5, sideMargin: 0.28, hPitch: 2.06, vPitch: 0.5, cols: 4, rows: 20, name: 'Avery 5167/8167 - Return Address', desc: '½" × 1¾", 80 per sheet' },
    '5161': { width: 4.0, height: 1.0, topMargin: 0.5, sideMargin: 0.16, hPitch: 4.19, vPitch: 1.0, cols: 2, rows: 10, name: 'Avery 5161/8161 - Large Address', desc: '1" × 4", 20 per sheet' },
    '5163': { width: 4.0, height: 2.0, topMargin: 0.5, sideMargin: 0.16, hPitch: 4.19, vPitch: 2.0, cols: 2, rows: 5, name: 'Avery 5163/8163 - Shipping Labels', desc: '2" × 4", 10 per sheet' },
    '5164': { width: 4.0, height: 3.33, topMargin: 0.5, sideMargin: 0.16, hPitch: 4.19, vPitch: 3.33, cols: 2, rows: 3, name: 'Avery 5164/8164 - Large Shipping', desc: '3⅓" × 4", 6 per sheet' }
};

// UK/EU Templates (A4 paper: 210mm x 297mm = 8.27" x 11.69")
// Measurements converted from mm to inches (mm / 25.4)
const UK_TEMPLATES = {
    'J8160': { width: 2.5, height: 1.5, topMargin: 0.59, sideMargin: 0.3, hPitch: 2.6, vPitch: 1.5, cols: 3, rows: 7, name: 'Avery J8160/L7160 - Standard Address', desc: '63.5 × 38.1mm, 21 per sheet', popular: true },
    'J8159': { width: 2.5, height: 1.335, topMargin: 0.52, sideMargin: 0.3, hPitch: 2.6, vPitch: 1.335, cols: 3, rows: 8, name: 'Avery J8159/L7159 - Address Labels', desc: '63.5 × 33.9mm, 24 per sheet' },
    'J8161': { width: 2.5, height: 1.81, topMargin: 0.55, sideMargin: 0.3, hPitch: 2.6, vPitch: 1.81, cols: 3, rows: 6, name: 'Avery J8161/L7161 - Large Address', desc: '63.5 × 46mm, 18 per sheet' },
    'J8162': { width: 3.94, height: 1.335, topMargin: 0.52, sideMargin: 0.3, hPitch: 4.04, vPitch: 1.335, cols: 2, rows: 8, name: 'Avery J8162/L7162 - Wide Address', desc: '100 × 33.9mm, 16 per sheet' },
    'J8163': { width: 3.94, height: 1.5, topMargin: 0.59, sideMargin: 0.3, hPitch: 4.04, vPitch: 1.5, cols: 2, rows: 7, name: 'Avery J8163/L7163 - Parcel Labels', desc: '100 × 38.1mm, 14 per sheet' },
    'J8165': { width: 3.94, height: 2.68, topMargin: 0.52, sideMargin: 0.3, hPitch: 4.04, vPitch: 2.68, cols: 2, rows: 4, name: 'Avery J8165/L7165 - Parcel Labels', desc: '100 × 68mm, 8 per sheet' },
    'J8166': { width: 3.94, height: 3.62, topMargin: 0.52, sideMargin: 0.3, hPitch: 4.04, vPitch: 3.62, cols: 2, rows: 3, name: 'Avery J8166/L7166 - Shipping Labels', desc: '100 × 92mm, 6 per sheet' },
    'J8167': { width: 7.87, height: 5.51, topMargin: 0.52, sideMargin: 0.2, hPitch: 7.87, vPitch: 5.51, cols: 1, rows: 2, name: 'Avery J8167/L7167 - Large Shipping', desc: '200 × 140mm, 2 per sheet' }
};

// Combined templates object (will be updated based on region)
let TEMPLATES = { ...US_TEMPLATES };
let currentRegion = 'US';
let currentTemplate = '8160';
let manualAddresses = [];
let previewAddresses = [];

// Editing state
let isEditing = false;
let editingIndex = -1;
let editingSource = null; // 'manual' or 'upload'

const STORAGE_KEY = 'labelMakerAddresses';
const UPLOAD_STORAGE_KEY = 'labelMakerUploadedAddresses';
const RETURN_STORAGE_KEY = 'labelMakerReturnAddress';
const REGION_STORAGE_KEY = 'labelMakerRegion';

// Style options state
let selectedIcon = 'none';
let selectedFont = 'helvetica';
let customIconData = null; // Stores custom SVG as data URL
let iconColor = '#000000'; // Default black
let currentEmoji = ''; // Stores selected emoji

// Icon position state (as percentage of label dimensions)
let iconPosX = 0.03; // 3% from left edge (default left position)
let iconPosY = 0.5;  // 50% from top (centered vertically)
let useCustomIconPosition = false; // Whether user has dragged the icon

// Snap and padding settings
let snapEnabled = true; // Whether snap-to-center is enabled
let iconTextPadding = 0; // No padding between icon and text

// Preview scale factor (pixels per inch for preview display)
const PREVIEW_SCALE = 96;

// Preset icon SVG paths (24x24 viewBox)
const PRESET_ICONS = {
    gift: 'M20 7h-1.1A5.5 5.5 0 0 0 12 2.6 5.5 5.5 0 0 0 5.1 7H4a2 2 0 0 0-2 2v2a1 1 0 0 0 1 1h1v7a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-7h1a1 1 0 0 0 1-1V9a2 2 0 0 0-2-2zm-8-2.7a3.5 3.5 0 0 1 3.5 3.2l.1.5H12V4.3zM8.5 7.5A3.5 3.5 0 0 1 12 4.3V8H8.4l.1-.5zM11 19H6v-7h5v7zm0-9H4V9h7v1zm2 9v-7h5v7h-5zm7-9h-7V9h7v1z',
    heart: 'M12 21.35l-1.45-1.32C5.4 15.36 2 12.28 2 8.5 2 5.42 4.42 3 7.5 3c1.74 0 3.41.81 4.5 2.09C13.09 3.81 14.76 3 16.5 3 19.58 3 22 5.42 22 8.5c0 3.78-3.4 6.86-8.55 11.54L12 21.35z',
    star: 'M12 17.27L18.18 21l-1.64-7.03L22 9.24l-7.19-.61L12 2 9.19 8.63 2 9.24l5.46 4.73L5.82 21z',
    snowflake: 'M11 1v4.07A8 8 0 0 0 5.07 11H1v2h4.07A8 8 0 0 0 11 18.93V23h2v-4.07A8 8 0 0 0 18.93 13H23v-2h-4.07A8 8 0 0 0 13 5.07V1h-2zm1 6a5 5 0 1 1 0 10 5 5 0 0 1 0-10z',
    tree: 'M12 2L4 12h3v4H4l8 6 8-6h-3v-4h3L12 2zm0 3.5L16.5 11H14v4h-4v-4H7.5L12 5.5z',
    home: 'M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z',
    mail: 'M20 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z',
    flower: 'M12 16a4 4 0 1 0 0-8 4 4 0 0 0 0 8zm0-14C9.5 2 8 4.5 8 6c0 1.2.5 2.3 1.3 3A6 6 0 0 0 6 12c0 1.2.3 2.3 1 3.3-.8.7-1.3 1.8-1.3 3 0 2.5 1.5 5 4.3 5 1.5 0 2.7-1 3.3-1.7.6.7 1.8 1.7 3.3 1.7 2.8 0 4.3-2.5 4.3-5 0-1.2-.5-2.3-1.3-3 .7-1 1-2.1 1-3.3a6 6 0 0 0-3.3-5.3c.8-.7 1.3-1.8 1.3-3 0-2.5-1.5-5-4.3-5-1.5 0-2.7 1-3.3 1.7-.6-.7-1.8-1.7-3.3-1.7z'
};

// Toggle style options section
function toggleStyleOptions(section = 'return') {
    const prefix = section === 'global' ? 'global' : '';
    const contentId = prefix ? 'globalStyleOptionsContent' : 'styleOptionsContent';
    const toggleId = prefix ? 'globalStyleOptionsToggle' : 'styleOptionsToggle';
    const selectedId = prefix ? 'globalStyleOptionsSelected' : 'styleOptionsSelected';

    const content = document.getElementById(contentId);
    const toggle = document.getElementById(toggleId);
    const selected = document.getElementById(selectedId);

    if (!content || !toggle) return;

    content.classList.toggle('active');
    const isActive = content.classList.contains('active');
    toggle.textContent = isActive ? '−' : '+';
    // Hide selected font text when expanded, show when collapsed
    if (selected) {
        selected.style.display = isActive ? 'none' : 'inline';
    }
}

// Update font from global selector and sync to return section
function updateFontFromGlobal() {
    const globalSelect = document.getElementById('globalFontSelect');
    const returnSelect = document.getElementById('fontSelect');
    if (globalSelect && returnSelect) {
        returnSelect.value = globalSelect.value;
    }
    updateFontPreview();
    updateGlobalFontPreview();
}

// Update font from return section selector and sync to global
function updateFontFromReturn() {
    const globalSelect = document.getElementById('globalFontSelect');
    const returnSelect = document.getElementById('fontSelect');
    if (globalSelect && returnSelect) {
        globalSelect.value = returnSelect.value;
    }
    updateFontPreview();
    updateGlobalFontPreview();
}

// Update global font preview text
function updateGlobalFontPreview() {
    const select = document.getElementById('globalFontSelect');
    const previewText = document.getElementById('globalFontPreviewText');
    const selectedText = document.getElementById('globalStyleOptionsSelected');
    if (!select) return;

    const fontMap = {
        'helvetica': 'Helvetica, Arial, sans-serif',
        'times': '"Times New Roman", Times, serif',
        'courier': '"Courier New", Courier, monospace'
    };

    const fontNames = {
        'helvetica': 'Helvetica',
        'times': 'Times New Roman',
        'courier': 'Courier'
    };

    if (previewText) {
        previewText.style.fontFamily = fontMap[select.value] || fontMap['helvetica'];
    }

    if (selectedText) {
        selectedText.textContent = '(' + (fontNames[select.value] || 'Helvetica') + ')';
    }
}

// Select icon
function selectIcon(iconName) {
    selectedIcon = iconName;

    // Update UI
    document.querySelectorAll('.icon-option').forEach(opt => {
        opt.classList.toggle('selected', opt.dataset.icon === iconName);
    });

    // If selecting a preset icon, clear custom icon and emoji
    if (iconName !== 'custom') {
        customIconData = null;
        document.getElementById('customIconName').textContent = '';
        document.getElementById('clearCustomIcon').style.display = 'none';
    }
    if (iconName !== 'emoji') {
        currentEmoji = '';
        // Reset emoji picker UI
        const preview = document.getElementById('selectedEmojiPreview');
        if (preview) {
            preview.textContent = 'Click to select emoji';
            preview.classList.remove('has-emoji');
        }
        const actions = document.getElementById('emojiPickerActions');
        if (actions) actions.style.display = 'none';
    }

    // Always reset to default centered position when selecting an icon
    // User must explicitly drag to use custom position
    useCustomIconPosition = false;
    iconPosX = 0.5;
    iconPosY = 0.5;

    // Update preview and controls visibility
    updatePreviewControlsVisibility();
    updateLabelPreview();
}

// Get selected icon position
function getIconPosition() {
    const select = document.getElementById('iconPosition');
    return select ? select.value : 'left';
}

// Set icon position and update preview
function setIconPosition(position) {
    // Reset to default position based on selection
    useCustomIconPosition = false;
    if (position === 'left') {
        iconPosX = 0.03;
    } else {
        iconPosX = 0.97;
    }
    iconPosY = 0.5;
    updateLabelPreview();
}

// Get selected font (check both selectors, they should be in sync)
function getSelectedFont() {
    const select = document.getElementById('fontSelect');
    const globalSelect = document.getElementById('globalFontSelect');
    return (select && select.value) || (globalSelect && globalSelect.value) || 'helvetica';
}

// Update font preview text to show selected font
function updateFontPreview() {
    const select = document.getElementById('fontSelect');
    const previewText = document.getElementById('fontPreviewText');
    const selectedText = document.getElementById('styleOptionsSelected');
    if (!select) return;

    const fontMap = {
        'helvetica': 'Helvetica, Arial, sans-serif',
        'times': '"Times New Roman", Times, serif',
        'courier': '"Courier New", Courier, monospace'
    };

    const fontNames = {
        'helvetica': 'Helvetica',
        'times': 'Times New Roman',
        'courier': 'Courier'
    };

    if (previewText) {
        previewText.style.fontFamily = fontMap[select.value] || fontMap['helvetica'];
    }

    // Update the header text to show selected font
    if (selectedText) {
        selectedText.textContent = '(' + (fontNames[select.value] || 'Helvetica') + ')';
    }

    updateLabelPreview();
}

// Handle custom SVG upload
function setupCustomIconUpload() {
    const input = document.getElementById('customIconInput');
    if (input) {
        input.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file) return;

            try {
                const svgText = await file.text();

                // Convert SVG to data URL for use in PDF
                const blob = new Blob([svgText], { type: 'image/svg+xml' });
                const url = URL.createObjectURL(blob);

                // Create an image to convert SVG to PNG (jsPDF works better with raster images)
                const img = new Image();
                img.onload = () => {
                    const canvas = document.createElement('canvas');
                    canvas.width = 100;
                    canvas.height = 100;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, 100, 100);
                    customIconData = canvas.toDataURL('image/png');
                    URL.revokeObjectURL(url);

                    // Update UI
                    selectedIcon = 'custom';
                    document.querySelectorAll('.icon-option').forEach(opt => {
                        opt.classList.remove('selected');
                    });
                    document.getElementById('customIconName').textContent = file.name;
                    document.getElementById('clearCustomIcon').style.display = '';
                };
                img.onerror = () => {
                    alert('Error loading SVG. Please make sure it\'s a valid SVG file.');
                    URL.revokeObjectURL(url);
                };
                img.src = url;

            } catch (error) {
                console.error('Error reading SVG file:', error);
                alert('Error reading file. Please try again.');
            }

            // Reset input
            input.value = '';
        });
    }
}

// Clear custom icon
function clearCustomIcon() {
    customIconData = null;
    selectedIcon = 'none';
    document.getElementById('customIconName').textContent = '';
    document.getElementById('clearCustomIcon').style.display = 'none';

    // Select "none" option
    document.querySelectorAll('.icon-option').forEach(opt => {
        opt.classList.toggle('selected', opt.dataset.icon === 'none');
    });

    // Update preview and controls visibility
    updatePreviewControlsVisibility();
    updateLabelPreview();
}

// Convert preset icon SVG path to data URL for PDF (with color support)
function getPresetIconDataUrl(iconName, color = null) {
    const path = PRESET_ICONS[iconName];
    if (!path) return null;

    const fillColor = color || iconColor || '#000000';
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="100" height="100"><path fill="${fillColor}" d="${path}"/></svg>`;
    const blob = new Blob([svg], { type: 'image/svg+xml' });
    const url = URL.createObjectURL(blob);

    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            const canvas = document.createElement('canvas');
            canvas.width = 100;
            canvas.height = 100;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, 100, 100);
            const dataUrl = canvas.toDataURL('image/png');
            URL.revokeObjectURL(url);
            resolve(dataUrl);
        };
        img.onerror = () => {
            URL.revokeObjectURL(url);
            resolve(null);
        };
        img.src = url;
    });
}

// Convert emoji to data URL for PDF
function getEmojiDataUrl(emoji) {
    return new Promise((resolve) => {
        const canvas = document.createElement('canvas');
        canvas.width = 100;
        canvas.height = 100;
        const ctx = canvas.getContext('2d');

        // Draw emoji centered on canvas
        ctx.font = '80px Arial, "Segoe UI Emoji", "Apple Color Emoji", sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(emoji, 50, 55);

        resolve(canvas.toDataURL('image/png'));
    });
}

// Handle icon color change
function setIconColor(color) {
    iconColor = color;
    // Regenerate cached icons with new color
    iconCache = {};
    preloadIcons();
    updateLabelPreview();
}

// Toggle emoji picker dropdown
function toggleEmojiPicker() {
    const dropdown = document.getElementById('emojiPickerDropdown');
    const toggle = document.getElementById('emojiPickerToggle');
    const isVisible = dropdown.style.display !== 'none';

    dropdown.style.display = isVisible ? 'none' : 'block';
    toggle.classList.toggle('active', !isVisible);
}

// Handle emoji selection from picker
function selectEmoji(emoji) {
    if (!emoji) {
        clearEmoji();
        return;
    }
    currentEmoji = emoji;
    selectedIcon = 'emoji';

    // Reset to default centered position
    useCustomIconPosition = false;
    iconPosX = 0.5;
    iconPosY = 0.5;

    // Update UI - deselect all icons
    document.querySelectorAll('.icon-option').forEach(opt => {
        opt.classList.remove('selected');
    });

    // Update preview text in toggle button
    const preview = document.getElementById('selectedEmojiPreview');
    if (preview) {
        preview.textContent = emoji;
        preview.classList.add('has-emoji');
    }

    // Show actions and update display
    const actions = document.getElementById('emojiPickerActions');
    const display = document.getElementById('selectedEmojiDisplay');
    if (actions) actions.style.display = 'flex';
    if (display) {
        display.innerHTML = `Selected: <span class="emoji-large">${emoji}</span>`;
    }

    // Close the dropdown
    const dropdown = document.getElementById('emojiPickerDropdown');
    const toggle = document.getElementById('emojiPickerToggle');
    if (dropdown) dropdown.style.display = 'none';
    if (toggle) toggle.classList.remove('active');

    // Update preview and controls visibility
    updatePreviewControlsVisibility();
    updateLabelPreview();
}

// Clear emoji
function clearEmoji() {
    currentEmoji = '';
    if (selectedIcon === 'emoji') {
        selectedIcon = 'none';
        document.querySelectorAll('.icon-option').forEach(opt => {
            opt.classList.toggle('selected', opt.dataset.icon === 'none');
        });
    }

    // Reset preview text in toggle button
    const preview = document.getElementById('selectedEmojiPreview');
    if (preview) {
        preview.textContent = 'Click to select emoji';
        preview.classList.remove('has-emoji');
    }

    // Hide actions
    const actions = document.getElementById('emojiPickerActions');
    if (actions) actions.style.display = 'none';

    // Update preview and controls visibility
    updatePreviewControlsVisibility();
    updateLabelPreview();
}

// Setup emoji picker event listener
function setupEmojiPicker() {
    const picker = document.getElementById('emojiPicker');
    if (picker) {
        picker.addEventListener('emoji-click', event => {
            selectEmoji(event.detail.unicode);
        });
    }

    // Close picker when clicking outside
    document.addEventListener('click', (e) => {
        const wrapper = document.querySelector('.emoji-picker-wrapper');
        const dropdown = document.getElementById('emojiPickerDropdown');
        const toggle = document.getElementById('emojiPickerToggle');

        if (wrapper && !wrapper.contains(e.target)) {
            if (dropdown) dropdown.style.display = 'none';
            if (toggle) toggle.classList.remove('active');
        }
    });
}

// Cache for icon data URLs
let iconCache = {};

// Update the label preview
function updateLabelPreview() {
    const preview = document.getElementById('labelPreview');
    const previewText = document.getElementById('previewText');
    const previewIcon = document.getElementById('previewIcon');
    const previewIconContent = document.getElementById('previewIconContent');
    const iconPositionInfo = document.getElementById('iconPositionInfo');

    if (!preview) return;

    // Update preview size based on template
    const template = TEMPLATES[currentTemplate];
    if (template) {
        preview.style.width = (template.width * PREVIEW_SCALE) + 'px';
        preview.style.height = (template.height * PREVIEW_SCALE) + 'px';
    }

    // Update text content from form
    const title = document.getElementById('returnTitle')?.value || '';
    const firstName = document.getElementById('returnFirst')?.value || '';
    const lastName = document.getElementById('returnLast')?.value || '';
    const suffix = document.getElementById('returnSuffix')?.value || '';
    const address1 = document.getElementById('returnAddress1')?.value || '123 Main Street';
    const address2 = document.getElementById('returnAddress2')?.value || '';
    const city = document.getElementById('returnCity')?.value || 'City';
    const state = document.getElementById('returnState')?.value || 'State';
    const postalCode = document.getElementById('returnPostalCode')?.value || '12345';

    // Build name line
    const nameParts = [title, firstName, lastName, suffix].filter(p => p.trim());
    const nameText = nameParts.length > 0 ? nameParts.join(' ') : 'Your Name';

    // Build address line
    const addressText = address2 ? `${address1}, ${address2}` : (address1 || '123 Main Street');

    // Build city line
    const cityText = `${city || 'City'}, ${state || 'ST'} ${postalCode || '12345'}`;

    // Update text in preview
    previewText.innerHTML = `
        <div class="preview-name">${escapeHtml(nameText)}</div>
        <div class="preview-address">${escapeHtml(addressText)}</div>
        <div class="preview-city">${escapeHtml(cityText)}</div>
    `;

    // Update font family
    const fontFamily = getSelectedFont();
    const fontMap = {
        'helvetica': 'Helvetica, Arial, sans-serif',
        'times': '"Times New Roman", Times, serif',
        'courier': 'Courier, monospace'
    };
    previewText.style.fontFamily = fontMap[fontFamily] || fontMap.helvetica;

    // Calculate common dimensions
    const labelWidth = template.width * PREVIEW_SCALE;
    const labelHeight = template.height * PREVIEW_SCALE;
    const edgePadding = Math.round(0.06 * PREVIEW_SCALE);

    // Update icon
    const hasIcon = selectedIcon !== 'none';
    if (hasIcon) {
        previewIcon.style.display = 'flex';

        // Calculate icon size (similar to PDF logic)
        const iconSize = Math.min(template.height * 0.6, template.width * 0.2, 0.5) * PREVIEW_SCALE;
        previewIcon.style.width = iconSize + 'px';
        previewIcon.style.height = iconSize + 'px';

        // Center icon vertically - this must match PDF calculation exactly
        const verticalCenter = (labelHeight / 2) - (iconSize / 2);

        if (useCustomIconPosition) {
            // Use custom dragged position
            let iconLeft = iconPosX * labelWidth - iconSize / 2;
            let iconTop = iconPosY * labelHeight - iconSize / 2;
            // Constrain to label bounds with padding (matching PDF behavior)
            iconLeft = Math.max(edgePadding, Math.min(labelWidth - edgePadding - iconSize, iconLeft));
            iconTop = Math.max(edgePadding, Math.min(labelHeight - edgePadding - iconSize, iconTop));
            previewIcon.style.left = iconLeft + 'px';
            previewIcon.style.top = iconTop + 'px';
        } else {
            // Use default position based on iconPosition setting
            const pos = getIconPosition();
            if (pos === 'left') {
                previewIcon.style.left = edgePadding + 'px';
                previewIcon.style.top = verticalCenter + 'px';
            } else {
                previewIcon.style.left = (labelWidth - iconSize - edgePadding) + 'px';
                previewIcon.style.top = verticalCenter + 'px';
            }
        }

        // Update icon content
        if (selectedIcon === 'emoji' && currentEmoji) {
            previewIconContent.innerHTML = currentEmoji;
            previewIconContent.style.fontSize = (iconSize * 0.8) + 'px';
        } else if (selectedIcon === 'custom' && customIconData) {
            previewIconContent.innerHTML = `<img src="${customIconData}" style="width:100%;height:100%">`;
        } else if (PRESET_ICONS[selectedIcon]) {
            const path = PRESET_ICONS[selectedIcon];
            previewIconContent.innerHTML = `<svg viewBox="0 0 24 24" style="width:100%;height:100%"><path fill="${iconColor}" d="${path}"/></svg>`;
        }

        // Update position info
        if (useCustomIconPosition) {
            iconPositionInfo.textContent = `Icon: Custom position (${Math.round(iconPosX * 100)}%, ${Math.round(iconPosY * 100)}%)`;
        } else {
            iconPositionInfo.textContent = `Icon: ${getIconPosition() === 'left' ? 'Left' : 'Right'} side`;
        }

        // Update text position based on icon location
        const iconLeft = parseFloat(previewIcon.style.left) || 0;
        const iconTop = parseFloat(previewIcon.style.top) || 0;
        updatePreviewTextPosition(iconLeft, iconTop, iconSize, labelWidth, labelHeight);
    } else {
        previewIcon.style.display = 'none';
        iconPositionInfo.textContent = 'Icon: Not selected';

        // Reset text position when no icon - center text vertically
        previewText.style.position = 'absolute';
        previewText.style.left = edgePadding + 'px';
        previewText.style.width = (labelWidth - edgePadding * 2) + 'px';
        previewText.style.top = '0px';
        previewText.style.marginLeft = '0';
        previewText.style.marginRight = '0';
        previewText.style.marginTop = '0';

        // Center text vertically after layout
        centerPreviewText(previewText, labelHeight);
    }
}

// Helper to center text vertically in preview
function centerPreviewText(textElement, labelHeight) {
    setTimeout(() => {
        const textHeight = textElement.offsetHeight;
        const centeredTop = (labelHeight - textHeight) / 2;
        textElement.style.top = Math.max(0, centeredTop) + 'px';
    }, 10);
}

// Get start padding based on label height (matches PDF logic)
function getStartPadding(labelHeight) {
    if (labelHeight <= 0.5) return 0.06;
    if (labelHeight <= 1.0) return 0.15;
    if (labelHeight <= 1.5) return 0.2;
    return 0.25;
}

// Reset icon position to default
function resetIconPosition() {
    useCustomIconPosition = false;
    iconPosX = 0.03;
    iconPosY = 0.5;
    updateLabelPreview();
}

// Toggle snap-to-center functionality
function toggleSnap(enabled) {
    snapEnabled = enabled;
}

// Toggle icon options section
function toggleIconOptions() {
    const content = document.getElementById('iconOptionsContent');
    const toggle = document.getElementById('iconOptionsToggle');
    if (content && toggle) {
        const isHidden = content.style.display === 'none';
        content.style.display = isHidden ? 'block' : 'none';
        toggle.textContent = isHidden ? '−' : '+';
    }
}

// Show/hide preview controls based on whether icon is selected
function updatePreviewControlsVisibility() {
    const hasIcon = selectedIcon !== 'none';
    const positionInfo = document.getElementById('previewPositionInfo');
    const controls = document.getElementById('previewControls');

    if (positionInfo) positionInfo.style.display = hasIcon ? 'flex' : 'none';
    if (controls) controls.style.display = hasIcon ? 'flex' : 'none';
}

// Snap threshold in pixels
const SNAP_THRESHOLD = 8;

// Setup drag and drop for icon positioning
function setupIconDragDrop() {
    const previewIcon = document.getElementById('previewIcon');
    const labelPreview = document.getElementById('labelPreview');
    const snapLineH = document.getElementById('snapLineH');
    const snapLineV = document.getElementById('snapLineV');

    if (!previewIcon || !labelPreview) return;

    let isDragging = false;
    let startX, startY;
    let iconStartX, iconStartY;

    function handleDragMove(clientX, clientY) {
        const rect = labelPreview.getBoundingClientRect();
        const template = TEMPLATES[currentTemplate];
        const iconSize = Math.min(template.height * 0.6, template.width * 0.2, 0.5) * PREVIEW_SCALE;

        // Calculate new position
        let newX = iconStartX + (clientX - startX);
        let newY = iconStartY + (clientY - startY);

        // Use same edge padding as PDF (0.06 inches * 96 DPI)
        const edgePadding = Math.round(0.06 * PREVIEW_SCALE);

        // Constrain to label bounds with padding (matching PDF behavior)
        newX = Math.max(edgePadding, Math.min(rect.width - edgePadding - iconSize, newX));
        newY = Math.max(edgePadding, Math.min(rect.height - edgePadding - iconSize, newY));

        // Minimum text width to prevent wrapping (40% of label width)
        const minTextWidth = rect.width * 0.4;
        const currentIconPadding = iconTextPadding;

        // Calculate where the icon center will be
        const iconCenterX = newX + iconSize / 2;
        const labelCenterX = rect.width / 2;

        // If icon is on left side, limit how far right it can go
        if (iconCenterX < labelCenterX) {
            const maxIconRight = rect.width - edgePadding - minTextWidth - currentIconPadding;
            if (newX + iconSize > maxIconRight) {
                newX = maxIconRight - iconSize;
            }
        }
        // If icon is on right side, limit how far left it can go
        else if (iconCenterX > labelCenterX) {
            const minIconLeft = edgePadding + minTextWidth + currentIconPadding;
            if (newX < minIconLeft) {
                newX = minIconLeft;
            }
        }

        // Re-apply edge constraints after text width constraints
        newX = Math.max(edgePadding, Math.min(rect.width - edgePadding - iconSize, newX));

        // Calculate center of icon (update after constraints)
        const finalIconCenterX = newX + iconSize / 2;
        const iconCenterY = newY + iconSize / 2;
        const labelCenterY = rect.height / 2;

        // Snap to center horizontally (only if snap is enabled)
        let snappedH = false;
        if (snapEnabled && Math.abs(finalIconCenterX - labelCenterX) < SNAP_THRESHOLD) {
            newX = labelCenterX - iconSize / 2;
            snappedH = true;
        }

        // Snap to center vertically (only if snap is enabled)
        let snappedV = false;
        if (snapEnabled && Math.abs(iconCenterY - labelCenterY) < SNAP_THRESHOLD) {
            newY = labelCenterY - iconSize / 2;
            snappedV = true;
        }

        // Show/hide snap lines (only show when snap is enabled)
        if (snapLineV) snapLineV.classList.toggle('visible', snapEnabled && snappedH);
        if (snapLineH) snapLineH.classList.toggle('visible', snapEnabled && snappedV);

        // Update icon position
        previewIcon.style.left = newX + 'px';
        previewIcon.style.top = newY + 'px';

        // Store as percentage (center of icon)
        iconPosX = (newX + iconSize / 2) / rect.width;
        iconPosY = (newY + iconSize / 2) / rect.height;
        useCustomIconPosition = true;

        // Update position info
        const iconPositionInfo = document.getElementById('iconPositionInfo');
        if (iconPositionInfo) {
            let posText = `Icon: Custom (${Math.round(iconPosX * 100)}%, ${Math.round(iconPosY * 100)}%)`;
            if (snappedH && snappedV) posText = 'Icon: Centered';
            else if (snappedH) posText = 'Icon: Centered horizontally';
            else if (snappedV) posText = 'Icon: Centered vertically';
            iconPositionInfo.textContent = posText;
        }

        // Update text positioning based on icon location
        updatePreviewTextPosition(newX, newY, iconSize, rect.width, rect.height);
    }

    function handleDragEnd() {
        if (isDragging) {
            isDragging = false;
            previewIcon.classList.remove('dragging');
            // Hide snap lines
            if (snapLineV) snapLineV.classList.remove('visible');
            if (snapLineH) snapLineH.classList.remove('visible');
        }
    }

    previewIcon.addEventListener('mousedown', (e) => {
        isDragging = true;
        previewIcon.classList.add('dragging');
        startX = e.clientX;
        startY = e.clientY;
        iconStartX = parseFloat(previewIcon.style.left) || 0;
        iconStartY = parseFloat(previewIcon.style.top) || 0;
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        handleDragMove(e.clientX, e.clientY);
    });

    document.addEventListener('mouseup', handleDragEnd);

    // Touch support
    previewIcon.addEventListener('touchstart', (e) => {
        isDragging = true;
        previewIcon.classList.add('dragging');
        const touch = e.touches[0];
        startX = touch.clientX;
        startY = touch.clientY;
        iconStartX = parseFloat(previewIcon.style.left) || 0;
        iconStartY = parseFloat(previewIcon.style.top) || 0;
        e.preventDefault();
    }, { passive: false });

    document.addEventListener('touchmove', (e) => {
        if (!isDragging) return;
        const touch = e.touches[0];
        handleDragMove(touch.clientX, touch.clientY);
    }, { passive: false });

    document.addEventListener('touchend', handleDragEnd);
}

// Update text position in preview based on icon location
function updatePreviewTextPosition(iconX, iconY, iconSize, labelWidth, labelHeight) {
    const previewText = document.getElementById('previewText');
    if (!previewText) return;

    const template = TEMPLATES[currentTemplate];
    const iconRight = iconX + iconSize;
    const iconBottom = iconY + iconSize;
    const iconCenterX = iconX + iconSize / 2;
    const iconCenterY = iconY + iconSize / 2;

    // Match PDF padding exactly: 0.06 inches * 96 DPI = ~6px
    const edgePadding = Math.round(0.06 * PREVIEW_SCALE);
    const currentIconPadding = iconTextPadding;

    let textLeft, textWidth;

    // Check if icon is horizontally centered
    const isHorizontallyCentered = Math.abs(iconCenterX - labelWidth / 2) < SNAP_THRESHOLD;

    if (isHorizontallyCentered) {
        // Icon is horizontally centered - text goes full width below or above
        textLeft = edgePadding;
        textWidth = labelWidth - (edgePadding * 2);
    } else if (iconCenterX < labelWidth / 2) {
        // Icon is on left side - text to the right of icon
        textLeft = iconRight + currentIconPadding;
        textWidth = labelWidth - edgePadding - iconRight - currentIconPadding;
    } else {
        // Icon is on right side - text on left
        textLeft = edgePadding;
        textWidth = iconX - edgePadding - currentIconPadding;
    }

    // Apply horizontal positions first so we can measure text height
    previewText.style.position = 'absolute';
    previewText.style.left = textLeft + 'px';
    previewText.style.width = textWidth + 'px';
    previewText.style.top = '0px';
    previewText.style.marginLeft = '0';
    previewText.style.marginRight = '0';
    previewText.style.marginTop = '0';

    // Center text vertically after layout
    if (isHorizontallyCentered && iconCenterY < labelHeight / 2) {
        // Icon in top half, text goes below icon
        previewText.style.top = (iconBottom + currentIconPadding) + 'px';
    } else {
        // Center text vertically in label
        centerPreviewText(previewText, labelHeight);
    }
}

// Setup return address form listeners for live preview
function setupReturnAddressPreviewListeners() {
    const fields = [
        'returnTitle', 'returnFirst', 'returnLast', 'returnSuffix',
        'returnAddress1', 'returnAddress2', 'returnCity', 'returnState',
        'returnPostalCode', 'returnCountry'
    ];

    fields.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', updateLabelPreview);
        }
    });
}

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const status = document.getElementById('status');
const templateSelect = document.getElementById('templateSelect');

templateSelect.addEventListener('change', (e) => {
    currentTemplate = e.target.value;
    updateManualFullSheetHint();
    updateManualStats();
    updateUploadFullSheetHint();
    updateUploadPreviewStats();
    updateReturnTemplateInfo();
    updateSkipLabelsMax();
    updateLabelPreview();
});

// Get the number of labels to skip
function getSkipCount() {
    const skipInput = document.getElementById('skipLabels');
    const value = parseInt(skipInput.value, 10);
    return isNaN(value) || value < 0 ? 0 : value;
}

// Update skip labels max based on template
function updateSkipLabelsMax() {
    const template = TEMPLATES[currentTemplate];
    if (!template) return;
    const labelsPerPage = template.cols * template.rows;
    const skipInput = document.getElementById('skipLabels');
    if (skipInput) {
        skipInput.max = labelsPerPage - 1;
        // Reset if current value exceeds max
        if (parseInt(skipInput.value, 10) >= labelsPerPage) {
            skipInput.value = labelsPerPage - 1;
        }
    }
}

// Reset skip labels to 0
function resetSkipLabels() {
    const skipInput = document.getElementById('skipLabels');
    if (skipInput) {
        skipInput.value = 0;
        updateManualStats();
        updateUploadPreviewStats();
    }
}

// Region switching function
function setRegion(region) {
    currentRegion = region;

    // Update button styles
    document.getElementById('regionUS').classList.toggle('active', region === 'US');
    document.getElementById('regionUK').classList.toggle('active', region === 'UK');

    // Update region info text
    const regionInfo = document.getElementById('regionInfo');
    if (region === 'US') {
        regionInfo.textContent = 'Paper: 8.5" × 11" (Letter)';
        TEMPLATES = { ...US_TEMPLATES };
        currentTemplate = '8160'; // Default US template
    } else {
        regionInfo.textContent = 'Paper: 210 × 297mm (A4)';
        TEMPLATES = { ...UK_TEMPLATES };
        currentTemplate = 'J8160'; // Default UK template
    }

    // Update template dropdown
    updateTemplateDropdown();

    // Save preference
    saveRegionToStorage();

    // Update all displays
    updateManualFullSheetHint();
    updateManualStats();
    updateUploadFullSheetHint();
    updateUploadPreviewStats();
    updateReturnTemplateInfo();
    updatePrintingInstructions();
    updateFormLabels();
    updateUploadEditFormLabels();
    updateFormatExample();
}

// Update template dropdown based on current region
function updateTemplateDropdown() {
    const select = document.getElementById('templateSelect');
    select.innerHTML = '';

    Object.keys(TEMPLATES).forEach(key => {
        const t = TEMPLATES[key];
        const option = document.createElement('option');
        option.value = key;
        option.textContent = `${t.name} (${t.desc})${t.popular ? ' ⭐ Most Popular' : ''}`;
        select.appendChild(option);
    });

    select.value = currentTemplate;
}

// Update printing instructions based on region
function updatePrintingInstructions() {
    const instructions = document.querySelector('.instructions ul');
    if (instructions) {
        const paperSize = currentRegion === 'US' ? 'US Letter (8.5" × 11")' : 'A4 (210 × 297mm)';
        instructions.innerHTML = `
            <li><strong>Print Scale:</strong> Set to <strong>100%</strong> or "Actual Size" (NOT "Fit to Page")</li>
            <li><strong>Paper Size:</strong> ${paperSize}</li>
            <li><strong>Test First:</strong> Print one page on regular paper and hold it up to the light against your label sheet to check alignment</li>
            <li><strong>No Margins Override:</strong> Use the PDF's built-in margins</li>
        `;
    }
}

// Update form labels and placeholders based on region
function updateFormLabels() {
    const isUS = currentRegion === 'US';

    // Define region-specific labels and placeholders
    const config = isUS ? {
        stateLabel: 'State',
        statePlaceholder: 'NY',
        postalLabel: 'ZIP Code *',
        postalPlaceholder: '10001',
        cityPlaceholder: 'New York',
        countryPlaceholder: 'USA',
        address2Hint: 'Per USPS, some addresses require this'
    } : {
        stateLabel: 'County/Region',
        statePlaceholder: 'Greater London',
        postalLabel: 'Postcode *',
        postalPlaceholder: 'SW1A 1AA',
        cityPlaceholder: 'London',
        countryPlaceholder: 'United Kingdom',
        address2Hint: 'Some addresses may require this on one line'
    };

    // Update Manual Entry form
    const manualStateLabel = document.getElementById('manualStateLabel');
    const manualPostalLabel = document.getElementById('manualPostalLabel');
    const manualState = document.getElementById('manualState');
    const manualPostalCode = document.getElementById('manualPostalCode');
    const manualCity = document.getElementById('manualCity');
    const manualCountry = document.getElementById('manualCountry');
    const manualAddress2Hint = document.getElementById('manualAddress2Hint');

    if (manualStateLabel) manualStateLabel.textContent = config.stateLabel;
    if (manualPostalLabel) manualPostalLabel.textContent = config.postalLabel;
    if (manualState) manualState.placeholder = config.statePlaceholder;
    if (manualPostalCode) manualPostalCode.placeholder = config.postalPlaceholder;
    if (manualCity) manualCity.placeholder = config.cityPlaceholder;
    if (manualCountry) manualCountry.placeholder = config.countryPlaceholder;
    if (manualAddress2Hint) manualAddress2Hint.textContent = config.address2Hint;

    // Update Return Address form
    const returnStateLabel = document.getElementById('returnStateLabel');
    const returnPostalLabel = document.getElementById('returnPostalLabel');
    const returnState = document.getElementById('returnState');
    const returnPostalCode = document.getElementById('returnPostalCode');
    const returnCity = document.getElementById('returnCity');
    const returnCountry = document.getElementById('returnCountry');
    const returnAddress2Hint = document.getElementById('returnAddress2Hint');

    if (returnStateLabel) returnStateLabel.textContent = config.stateLabel;
    if (returnPostalLabel) returnPostalLabel.textContent = config.postalLabel;
    if (returnState) returnState.placeholder = config.statePlaceholder;
    if (returnPostalCode) returnPostalCode.placeholder = config.postalPlaceholder;
    if (returnCity) returnCity.placeholder = config.cityPlaceholder;
    if (returnCountry) returnCountry.placeholder = config.countryPlaceholder;
    if (returnAddress2Hint) returnAddress2Hint.textContent = config.address2Hint;
}

// Save region preference to localStorage
function saveRegionToStorage() {
    try {
        localStorage.setItem(REGION_STORAGE_KEY, currentRegion);
    } catch (e) {
        console.warn('Could not save region to localStorage:', e);
    }
}

// Load region preference from localStorage
function loadRegionFromStorage() {
    try {
        const saved = localStorage.getItem(REGION_STORAGE_KEY);
        if (saved && (saved === 'US' || saved === 'UK')) {
            setRegion(saved);
        }
    } catch (e) {
        console.warn('Could not load region from localStorage:', e);
    }
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    // Initialize templates dropdown, form labels, and format example (before loading region which may change them)
    updateTemplateDropdown();
    updateFormLabels();
    updateUploadEditFormLabels();
    updateFormatExample();
    loadRegionFromStorage();
    loadAddressesFromStorage();
    loadUploadedAddressesFromStorage();
    loadReturnAddressFromStorage();
    setupAddress2Listeners();
    updateReturnTemplateInfo();
    updateSkipLabelsMax();
    setupSkipLabelsListener();
    setupCustomIconUpload();
    preloadIcons();
    setupReturnAddressPreviewListeners();
    setupIconDragDrop();
    setupStyleOptionListeners();
    setupEmojiPicker();
    updateLabelPreview();

    // Re-center preview text after fonts and layout are fully loaded
    setTimeout(() => updateLabelPreview(), 100);
});

// Also update preview after all resources (fonts, images) are loaded
window.addEventListener('load', () => {
    updateLabelPreview();
});

// Setup listeners for style options to update preview
function setupStyleOptionListeners() {
    // Font select - initialize preview and add listener
    const fontSelect = document.getElementById('fontSelect');
    const globalFontSelect = document.getElementById('globalFontSelect');
    if (fontSelect) {
        updateFontPreview(); // Initialize font preview on load
    }
    if (globalFontSelect) {
        updateGlobalFontPreview(); // Initialize global font preview on load
    }

    // Icon position
    const iconPosition = document.getElementById('iconPosition');
    if (iconPosition) {
        iconPosition.addEventListener('change', () => {
            // When changing position via dropdown, reset custom position
            useCustomIconPosition = false;
            updateLabelPreview();
        });
    }
}

// Preload all preset icons to cache
async function preloadIcons() {
    for (const iconName of Object.keys(PRESET_ICONS)) {
        if (!iconCache[iconName]) {
            iconCache[iconName] = await getPresetIconDataUrl(iconName);
        }
    }
}

// Setup skip labels input listener
function setupSkipLabelsListener() {
    const skipInput = document.getElementById('skipLabels');
    if (skipInput) {
        skipInput.addEventListener('input', () => {
            updateManualStats();
            updateUploadPreviewStats();
        });
    }
}

// Setup address2 field listeners to show/hide same-line option
function setupAddress2Listeners() {
    const manualAddr2 = document.getElementById('manualAddress2');
    const returnAddr2 = document.getElementById('returnAddress2');
    const uploadEditAddr2 = document.getElementById('uploadEditAddress2');

    if (manualAddr2) {
        manualAddr2.addEventListener('input', () => {
            const option = document.getElementById('address2Option');
            option.style.display = manualAddr2.value.trim() ? '' : 'none';
        });
    }

    if (returnAddr2) {
        returnAddr2.addEventListener('input', () => {
            const option = document.getElementById('returnAddress2Option');
            option.style.display = returnAddr2.value.trim() ? '' : 'none';
        });
    }

    if (uploadEditAddr2) {
        uploadEditAddr2.addEventListener('input', () => {
            const option = document.getElementById('uploadEditAddress2Option');
            option.style.display = uploadEditAddr2.value.trim() ? '' : 'none';
        });
    }
}

// Update return template info display
function updateReturnTemplateInfo() {
    const template = TEMPLATES[currentTemplate];
    if (!template) return;
    const labelsPerPage = template.cols * template.rows;
    const info = document.getElementById('returnTemplateInfo');
    if (info) {
        // Use the name from the template object, or fall back to the template key
        const templateName = template.name ? template.name.split(' - ')[0] : currentTemplate;
        info.textContent = `${templateName} - ${labelsPerPage} labels per sheet`;
    }
}

// Save addresses to localStorage
function saveAddressesToStorage() {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(manualAddresses));
    } catch (e) {
        console.warn('Could not save to localStorage:', e);
    }
}

// Load addresses from localStorage
function loadAddressesFromStorage() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved) {
            manualAddresses = JSON.parse(saved);
            updateAddressesList();
            if (manualAddresses.length > 0) {
                showStatus(`Restored ${manualAddresses.length} saved address${manualAddresses.length !== 1 ? 'es' : ''} from your last session.`, 'success');
            }
        }
    } catch (e) {
        console.warn('Could not load from localStorage:', e);
    }
}

// Save uploaded addresses to localStorage
function saveUploadedAddressesToStorage() {
    try {
        localStorage.setItem(UPLOAD_STORAGE_KEY, JSON.stringify(previewAddresses));
    } catch (e) {
        console.warn('Could not save uploaded addresses to localStorage:', e);
    }
}

// Load uploaded addresses from localStorage
function loadUploadedAddressesFromStorage() {
    try {
        const saved = localStorage.getItem(UPLOAD_STORAGE_KEY);
        if (saved) {
            previewAddresses = JSON.parse(saved);
            if (previewAddresses.length > 0) {
                showPreview(previewAddresses);
                showStatus(`Restored ${previewAddresses.length} uploaded address${previewAddresses.length !== 1 ? 'es' : ''} from your last session.`, 'success');
            }
        }
    } catch (e) {
        console.warn('Could not load uploaded addresses from localStorage:', e);
    }
}

// Clear uploaded addresses from storage
function clearUploadedAddressesFromStorage() {
    try {
        localStorage.removeItem(UPLOAD_STORAGE_KEY);
    } catch (e) {
        console.warn('Could not clear uploaded addresses from localStorage:', e);
    }
}

// Save return address to localStorage
function saveReturnAddressToStorage() {
    const returnAddress = getReturnAddressFromForm();
    if (returnAddress) {
        try {
            localStorage.setItem(RETURN_STORAGE_KEY, JSON.stringify(returnAddress));
        } catch (e) {
            console.warn('Could not save return address to localStorage:', e);
        }
    }
}

// Load return address from localStorage
function loadReturnAddressFromStorage() {
    try {
        const saved = localStorage.getItem(RETURN_STORAGE_KEY);
        if (saved) {
            const addr = JSON.parse(saved);
            document.getElementById('returnTitle').value = addr.title || '';
            document.getElementById('returnFirst').value = addr.firstName || '';
            document.getElementById('returnLast').value = addr.lastName || '';
            document.getElementById('returnSuffix').value = addr.suffix || '';
            document.getElementById('returnAddress1').value = addr.address1 || '';
            document.getElementById('returnAddress2').value = addr.address2 || '';
            document.getElementById('returnAddress2SameLine').checked = addr.address2SameLine || false;
            document.getElementById('returnAddress2Option').style.display = addr.address2 ? '' : 'none';
            document.getElementById('returnCity').value = addr.city || '';
            document.getElementById('returnState').value = addr.state || '';
            document.getElementById('returnPostalCode').value = addr.postalCode || '';
            document.getElementById('returnCountry').value = addr.country || '';
        }
    } catch (e) {
        console.warn('Could not load return address from localStorage:', e);
    }
}

// Get return address from form
function getReturnAddressFromForm() {
    const firstName = document.getElementById('returnFirst').value.trim();
    const lastName = document.getElementById('returnLast').value.trim();
    const address1 = document.getElementById('returnAddress1').value.trim();
    const city = document.getElementById('returnCity').value.trim();
    const state = document.getElementById('returnState').value.trim();
    const postalCode = document.getElementById('returnPostalCode').value.trim();

    // For return address: First name is optional (allows "The Smith Family" style)
    // State/County is also optional (supports international addresses)
    if (!lastName || !address1 || !city || !postalCode) {
        return null;
    }

    return {
        title: document.getElementById('returnTitle').value.trim(),
        firstName,
        lastName,
        suffix: document.getElementById('returnSuffix').value.trim(),
        address1,
        address2: document.getElementById('returnAddress2').value.trim(),
        address2SameLine: document.getElementById('returnAddress2SameLine').checked,
        city,
        state,
        postalCode,
        country: document.getElementById('returnCountry').value.trim()
    };
}

// Generate return address labels (full sheet)
async function generateReturnLabels() {
    const returnAddress = getReturnAddressFromForm();
    if (!returnAddress) {
        alert('Please fill in all required fields (marked with *)');
        return;
    }

    // Save for next time
    saveReturnAddressToStorage();

    const template = TEMPLATES[currentTemplate];
    const labelsPerPage = template.cols * template.rows;
    const skipCount = getSkipCount();

    // Fill remaining positions after skip with the same address
    const fillCount = labelsPerPage - skipCount;
    let addressesToPrint = Array(fillCount).fill(returnAddress);

    // Prepend empty entries for skipped labels
    if (skipCount > 0) {
        addressesToPrint = [...Array(skipCount).fill(null), ...addressesToPrint];
    }

    showStatus(`Generating ${fillCount} return address labels (skipping ${skipCount})...`, 'processing');

    try {
        const pdf = await generateLabelsPDF(addressesToPrint);
        pdf.save(`${currentTemplate}-return-labels.pdf`);

        let successMsg = `✓ Success! Generated ${fillCount} return address labels (1 page).`;
        if (skipCount > 0) {
            successMsg += ` Skipped first ${skipCount} position${skipCount > 1 ? 's' : ''}.`;
        }
        successMsg += ' PDF download started.';
        showStatus(successMsg, 'success');
    } catch (error) {
        console.error(error);
        showStatus('Error generating labels: ' + error.message, 'error');
    }
}

// Clear return address form
function clearReturnForm() {
    document.getElementById('returnTitle').value = '';
    document.getElementById('returnFirst').value = '';
    document.getElementById('returnLast').value = '';
    document.getElementById('returnSuffix').value = '';
    document.getElementById('returnAddress1').value = '';
    document.getElementById('returnAddress2').value = '';
    document.getElementById('returnAddress2SameLine').checked = false;
    document.getElementById('returnAddress2Option').style.display = 'none';
    document.getElementById('returnCity').value = '';
    document.getElementById('returnState').value = '';
    document.getElementById('returnPostalCode').value = '';
    document.getElementById('returnCountry').value = '';
    try {
        localStorage.removeItem(RETURN_STORAGE_KEY);
    } catch (e) {}
}

// Clear all saved addresses
function clearAllAddresses() {
    if (manualAddresses.length === 0) return;
    if (confirm(`Are you sure you want to delete all ${manualAddresses.length} address${manualAddresses.length !== 1 ? 'es' : ''}? This cannot be undone.`)) {
        manualAddresses = [];
        saveAddressesToStorage();
        updateAddressesList();
        showStatus('All addresses have been cleared.', 'success');
    }
}

function switchTab(tab) {
    const buttons = document.querySelectorAll('.tab-button');
    const sections = document.querySelectorAll('.input-section');

    buttons.forEach(btn => btn.classList.remove('active'));
    sections.forEach(section => section.classList.remove('active'));

    if (tab === 'upload') {
        buttons[0].classList.add('active');
        document.getElementById('uploadSection').classList.add('active');
    } else if (tab === 'manual') {
        buttons[1].classList.add('active');
        document.getElementById('manualSection').classList.add('active');
    } else if (tab === 'return') {
        buttons[2].classList.add('active');
        document.getElementById('returnSection').classList.add('active');
    }
}

// File upload handlers
dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

function showStatus(message, type) {
    status.textContent = message;
    status.className = `status ${type}`;
    status.style.display = 'block';
}

function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    result.push(current.trim());
    return result;
}

function parseData(content) {
    const lines = content.trim().split('\n');
    if (lines.length < 2) return [];

    // Parse header to detect format
    let headerParts;
    if (lines[0].includes('\t')) {
        headerParts = lines[0].split('\t').map(p => p.trim().toLowerCase());
    } else {
        headerParts = parseCSVLine(lines[0]).map(p => p.toLowerCase());
    }

    // Detect region from headers and auto-switch if needed
    const hasUKHeaders = headerParts.some(h =>
        h.includes('county') || h.includes('region') || h === 'postcode'
    );
    const hasUSHeaders = headerParts.some(h =>
        h === 'state' || h.includes('zip')
    );

    // Auto-switch region if we detect a clear match
    if (hasUKHeaders && !hasUSHeaders && currentRegion === 'US') {
        setRegion('UK');
        showStatus('Detected UK/EU format - automatically switched to A4 templates.', 'success');
    } else if (hasUSHeaders && !hasUKHeaders && currentRegion === 'UK') {
        setRegion('US');
        showStatus('Detected US format - automatically switched to Letter templates.', 'success');
    }

    // Detect if new format (has "first name" or similar columns)
    const isNewFormat = headerParts.some(h =>
        h.includes('first') || h.includes('last') || h.includes('postal') || h.includes('address 1')
    );

    const dataLines = lines.slice(1);

    return dataLines.map(line => {
        let parts;
        if (line.includes('\t')) {
            parts = line.split('\t').map(p => p.trim());
        } else {
            parts = parseCSVLine(line);
        }

        if (isNewFormat && parts.length >= 7) {
            // New expanded format: Title, First, Last, Suffix, Address1, Address2, City, State, PostalCode, Country
            return {
                title: parts[0] || '',
                firstName: parts[1] || '',
                lastName: parts[2] || '',
                suffix: parts[3] || '',
                address1: parts[4] || '',
                address2: parts[5] || '',
                address2SameLine: false,
                city: parts[6] || '',
                state: parts[7] || '',
                postalCode: parts[8] || '',
                country: parts[9] || ''
            };
        } else if (parts.length >= 3) {
            // Old 3-column format: Fullname, Address, CityStateZip
            const fullname = parts[0];
            const address = parts[1];
            let cityStateZip;

            if (parts.length === 3) {
                cityStateZip = parts[2];
            } else if (parts.length === 4) {
                cityStateZip = parts[2] + ' ' + parts[3];
            } else {
                cityStateZip = parts.slice(2).join(', ');
            }

            return { fullname, address, cityStateZip };
        }
        return null;
    }).filter(item => item !== null);
}

// Format address data for label - handles both old and new formats
function formatAddressForLabel(addressData) {
    // Check if this is the new format (has firstName field) or old format
    if (addressData.firstName !== undefined) {
        // New format - build name line from components
        const nameParts = [];
        if (addressData.title) nameParts.push(addressData.title);
        if (addressData.firstName) nameParts.push(addressData.firstName);
        if (addressData.lastName) nameParts.push(addressData.lastName);
        if (addressData.suffix) nameParts.push(addressData.suffix);
        const fullname = nameParts.join(' ');

        // Build address lines
        let addressLine1 = addressData.address1 || '';
        let addressLine2 = addressData.address2 || '';
        const address2SameLine = addressData.address2SameLine || false;

        // Store individual components for region-specific formatting
        const city = addressData.city || '';
        const state = addressData.state || '';
        const postalCode = addressData.postalCode || '';
        const country = addressData.country || '';

        // Build US-style location line (City, State ZIP)
        const locationParts = [];
        if (city) locationParts.push(city);
        if (state) {
            if (locationParts.length > 0) {
                locationParts[locationParts.length - 1] += ',';
            }
            locationParts.push(state);
        }
        if (postalCode) locationParts.push(postalCode);
        if (country && country.toUpperCase() !== 'USA' && country.toUpperCase() !== 'US') {
            locationParts.push(country);
        }
        const location = locationParts.join(' ');

        return { fullname, addressLine1, addressLine2, address2SameLine, location, city, state, postalCode, country };
    } else {
        // Old format - use as-is for backward compatibility
        return {
            fullname: addressData.fullname || '',
            addressLine1: addressData.address || '',
            addressLine2: '',
            address2SameLine: false,
            location: addressData.cityStateZip || '',
            city: '',
            state: '',
            postalCode: '',
            country: ''
        };
    }
}

function drawLabel(pdf, addressData, x, y, template, iconDataUrl) {
    if (!addressData) return;

    const formatted = formatAddressForLabel(addressData);
    const fontFamily = getSelectedFont();
    const iconPosition = getIconPosition();
    const hasIcon = iconDataUrl && selectedIcon !== 'none';

    // Build all address lines
    const addressLines = [];
    addressLines.push({ text: formatted.fullname, isBold: true });

    if (formatted.address2SameLine && formatted.addressLine2) {
        addressLines.push({ text: formatted.addressLine1 + ' ' + formatted.addressLine2, isBold: false });
    } else {
        addressLines.push({ text: formatted.addressLine1, isBold: false });
        if (formatted.addressLine2) {
            addressLines.push({ text: formatted.addressLine2, isBold: false });
        }
    }

    // For UK/international: City and Postcode on separate lines per Royal Mail guidelines
    // For US: City, State ZIP on one line per USPS guidelines
    if (currentRegion === 'UK') {
        // UK format per Royal Mail:
        // - POST TOWN (uppercase) - county is optional but can be included
        // - POSTCODE (uppercase, on its own line)
        // - COUNTRY (uppercase, only if international)
        let cityLine = '';
        if (formatted.city) {
            cityLine = formatted.city.toUpperCase();
        }
        // Add county/region if provided (Royal Mail says optional but some prefer it)
        if (formatted.state) {
            cityLine += (cityLine ? ', ' : '') + formatted.state;
        }
        if (cityLine) {
            addressLines.push({ text: cityLine, isBold: false });
        }
        if (formatted.postalCode) {
            addressLines.push({ text: formatted.postalCode.toUpperCase(), isBold: false });
        }
        // Only show country if it's NOT United Kingdom (for international mail TO UK)
        if (formatted.country &&
            formatted.country.toUpperCase() !== 'UNITED KINGDOM' &&
            formatted.country.toUpperCase() !== 'UK' &&
            formatted.country.toUpperCase() !== 'GB' &&
            formatted.country.toUpperCase() !== 'GREAT BRITAIN') {
            addressLines.push({ text: formatted.country.toUpperCase(), isBold: false });
        }
    } else {
        // US format per USPS: CITY, STATE ZIP on one line
        addressLines.push({ text: formatted.location, isBold: false });
    }

    // Filter out empty lines
    const filteredLines = addressLines.filter(line => line.text && line.text.trim());

    // Base font sizes based on label size
    let baseFontSize, minFontSize, lineSpacingRatio, startPadding;

    if (template.height <= 0.5) {
        baseFontSize = 7;
        minFontSize = 5;
        lineSpacingRatio = 1.15;
        startPadding = 0.06;
    } else if (template.height <= 1.0) {
        baseFontSize = 10;
        minFontSize = 7;
        lineSpacingRatio = 1.3;
        startPadding = 0.15;
    } else if (template.height <= 1.5) {
        baseFontSize = 11;
        minFontSize = 8;
        lineSpacingRatio = 1.3;
        startPadding = 0.2;
    } else {
        baseFontSize = 12;
        minFontSize = 9;
        lineSpacingRatio = 1.35;
        startPadding = 0.25;
    }

    const padding = 0.06;

    // Calculate icon size based on label height (icon is square)
    const iconSize = hasIcon ? Math.min(template.height * 0.6, template.width * 0.2, 0.5) : 0;
    // Convert pixel padding to inches (96 pixels per inch in preview)
    const iconPadding = hasIcon ? (iconTextPadding / PREVIEW_SCALE) : 0;

    // Adjust text area for icon
    let textStartX, textMaxWidth, iconX, iconY;

    // Track if we need to adjust vertical start for centered icon
    let textStartY = null;

    if (hasIcon && useCustomIconPosition) {
        // Custom position - icon placed at user-specified location
        iconX = x + (iconPosX * template.width) - (iconSize / 2);
        iconY = y + (iconPosY * template.height) - (iconSize / 2);

        // Constrain to label bounds
        iconX = Math.max(x + padding, Math.min(x + template.width - padding - iconSize, iconX));
        iconY = Math.max(y + padding, Math.min(y + template.height - padding - iconSize, iconY));

        // Determine if icon is on left or right half to adjust text
        const iconCenterX = iconX + iconSize / 2;
        const iconCenterY = iconY + iconSize / 2;
        const labelCenterX = x + template.width / 2;
        const labelCenterY = y + template.height / 2;

        // Check if icon is horizontally centered (within snap threshold converted to inches)
        const snapThresholdInches = 0.05; // ~5 pixels at 96dpi
        const isHorizontallyCentered = Math.abs(iconCenterX - labelCenterX) < snapThresholdInches;

        if (isHorizontallyCentered) {
            // Icon is centered - text goes below if icon is in top half, above if in bottom
            textStartX = x + padding;
            textMaxWidth = template.width - (padding * 2);
            if (iconCenterY < labelCenterY) {
                // Icon in top half, text starts below icon
                textStartY = iconY + iconSize + iconPadding;
            }
            // If icon in bottom half, text starts at normal position (top)
        } else if (iconCenterX < labelCenterX) {
            // Icon is on left side
            textStartX = iconX + iconSize + iconPadding;
            textMaxWidth = template.width - padding - (iconX - x) - iconSize - iconPadding;
        } else {
            // Icon is on right side
            textStartX = x + padding;
            textMaxWidth = iconX - x - padding - iconPadding;
        }
    } else if (hasIcon && iconPosition === 'left') {
        iconX = x + padding;
        // Center icon vertically within the label
        iconY = y + (template.height / 2) - iconSize;
        textStartX = x + padding + iconSize + iconPadding;
        textMaxWidth = template.width - (padding * 2) - iconSize - iconPadding;
    } else if (hasIcon && iconPosition === 'right') {
        textStartX = x + padding;
        textMaxWidth = template.width - (padding * 2) - iconSize - iconPadding;
        iconX = x + template.width - padding - iconSize;
        // Center icon vertically within the label
        iconY = y + (template.height / 2) - iconSize;
    } else {
        textStartX = x + padding;
        textMaxWidth = template.width - (padding * 2);
    }

    const maxHeight = template.height - (padding * 2) - startPadding;

    // Auto-fit: Find optimal font size that fits all content
    let fontSize = baseFontSize;
    let totalHeight;

    do {
        pdf.setFontSize(fontSize);
        const lineSpacing = (fontSize / 72) * lineSpacingRatio; // Convert pt to inches
        totalHeight = 0;

        for (const line of filteredLines) {
            pdf.setFont(fontFamily, line.isBold ? "bold" : "normal");
            const wrappedLines = pdf.splitTextToSize(line.text, textMaxWidth);
            totalHeight += wrappedLines.length * lineSpacing;
        }

        if (totalHeight <= maxHeight) break;
        fontSize -= 0.5;
    } while (fontSize >= minFontSize);

    // Draw icon if present
    if (hasIcon && iconDataUrl && iconX !== undefined && iconY !== undefined) {
        try {
            pdf.addImage(iconDataUrl, 'PNG', iconX, iconY, iconSize, iconSize);
        } catch (e) {
            console.warn('Could not add icon to PDF:', e);
        }
    }

    // Now draw the text - vertically centered in label
    const lineSpacing = (fontSize / 72) * lineSpacingRatio;
    // Calculate vertical center position for text block
    // totalHeight already calculated in the font-fitting loop above
    const textBlockTop = y + (template.height - totalHeight) / 2;
    let currentY = textStartY !== null ? textStartY : textBlockTop + lineSpacing * 0.85;

    for (const line of filteredLines) {
        pdf.setFont(fontFamily, line.isBold ? "bold" : "normal");
        pdf.setFontSize(line.isBold ? fontSize + 1 : fontSize);

        const wrappedLines = pdf.splitTextToSize(line.text, textMaxWidth);
        for (const wrappedLine of wrappedLines) {
            if (currentY > y + template.height - padding) break;
            pdf.text(wrappedLine, textStartX, currentY);
            currentY += lineSpacing;
        }
    }
}

async function generateLabelsPDF(addresses) {
    const template = TEMPLATES[currentTemplate];
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'in',
        format: currentRegion === 'US' ? 'letter' : 'a4'
    });

    // Get icon data URL if an icon is selected
    let iconDataUrl = null;
    if (selectedIcon !== 'none') {
        if (selectedIcon === 'custom' && customIconData) {
            iconDataUrl = customIconData;
        } else if (selectedIcon === 'emoji' && currentEmoji) {
            iconDataUrl = await getEmojiDataUrl(currentEmoji);
        } else if (iconCache[selectedIcon]) {
            iconDataUrl = iconCache[selectedIcon];
        } else {
            iconDataUrl = await getPresetIconDataUrl(selectedIcon);
        }
    }

    const labelsPerPage = template.cols * template.rows;
    const totalPages = Math.ceil(addresses.length / labelsPerPage);

    for (let page = 0; page < totalPages; page++) {
        if (page > 0) pdf.addPage();

        const pageStartIndex = page * labelsPerPage;

        for (let row = 0; row < template.rows; row++) {
            for (let col = 0; col < template.cols; col++) {
                const labelIndex = pageStartIndex + (row * template.cols) + col;
                if (labelIndex >= addresses.length) break;

                const addressData = addresses[labelIndex];
                const x = template.sideMargin + (col * template.hPitch);
                const y = template.topMargin + (row * template.vPitch);

                drawLabel(pdf, addressData, x, y, template, iconDataUrl);
            }
        }
    }

    return pdf;
}

async function handleFile(file) {
    showStatus('Processing file...', 'processing');

    try {
        const content = await file.text();
        const addresses = parseData(content);

        if (addresses.length === 0) {
            showStatus('No valid addresses found in file. Please check the format.', 'error');
            return;
        }

        // Store addresses for preview and save to storage
        previewAddresses = addresses;
        saveUploadedAddressesToStorage();
        showPreview(addresses);
        showStatus(`Found ${addresses.length} addresses. Review below and click "Generate PDF" to create labels.`, 'success');

    } catch (error) {
        console.error(error);
        showStatus('Error processing file: ' + error.message, 'error');
    }
}

function showPreview(addresses) {
    const previewSection = document.getElementById('previewSection');
    const previewList = document.getElementById('previewList');
    const fullSheetOption = document.getElementById('uploadFullSheetOption');
    const fullSheetCheckbox = document.getElementById('uploadFullSheetCheckbox');

    // Show/hide full sheet option based on address count
    if (addresses.length === 1) {
        fullSheetOption.style.display = '';
        updateUploadFullSheetHint();
    } else {
        fullSheetOption.style.display = 'none';
        fullSheetCheckbox.checked = false;
    }

    // Update stats
    updateUploadPreviewStats();

    // Build preview list with edit/delete buttons
    renderUploadPreviewList();

    // Show preview section
    previewSection.classList.add('active');
}

function renderUploadPreviewList() {
    const previewList = document.getElementById('previewList');
    previewList.innerHTML = previewAddresses.map((addr, index) => {
        const display = getAddressDisplayText(addr);
        return `
        <div class="preview-item" id="upload-preview-${index}">
            <div class="preview-item-content">
                <div class="preview-item-number">Address #${index + 1}</div>
                <div class="preview-item-name">${escapeHtml(display.name)}</div>
                <div class="preview-item-details">${display.address}<br>${escapeHtml(display.location)}</div>
            </div>
            <div class="preview-item-actions">
                <button class="btn-edit" onclick="editUploadAddress(${index})">Edit</button>
                <button class="btn-delete" onclick="deleteUploadAddress(${index})">Delete</button>
            </div>
        </div>
    `}).join('');

    // Update full sheet visibility
    const fullSheetOption = document.getElementById('uploadFullSheetOption');
    if (previewAddresses.length === 1) {
        fullSheetOption.style.display = '';
    } else {
        fullSheetOption.style.display = 'none';
        document.getElementById('uploadFullSheetCheckbox').checked = false;
    }

    updateUploadPreviewStats();
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function editUploadAddress(index) {
    const addr = previewAddresses[index];

    // Set editing state
    isEditing = true;
    editingIndex = index;
    editingSource = 'upload';

    // Show inline edit form
    const editForm = document.getElementById('uploadEditForm');
    editForm.style.display = '';

    // Load address into upload edit form
    if (addr.firstName !== undefined) {
        // New format
        document.getElementById('uploadEditTitle').value = addr.title || '';
        document.getElementById('uploadEditFirst').value = addr.firstName || '';
        document.getElementById('uploadEditLast').value = addr.lastName || '';
        document.getElementById('uploadEditSuffix').value = addr.suffix || '';
        document.getElementById('uploadEditAddress1').value = addr.address1 || '';
        document.getElementById('uploadEditAddress2').value = addr.address2 || '';
        document.getElementById('uploadEditAddress2SameLine').checked = addr.address2SameLine || false;
        document.getElementById('uploadEditAddress2Option').style.display = addr.address2 ? '' : 'none';
        document.getElementById('uploadEditCity').value = addr.city || '';
        document.getElementById('uploadEditState').value = addr.state || '';
        document.getElementById('uploadEditPostalCode').value = addr.postalCode || '';
        document.getElementById('uploadEditCountry').value = addr.country || '';
    } else {
        // Old format - parse as best we can
        const nameParts = (addr.fullname || '').split(' ');
        document.getElementById('uploadEditTitle').value = '';
        document.getElementById('uploadEditFirst').value = nameParts[0] || '';
        document.getElementById('uploadEditLast').value = nameParts.slice(1).join(' ') || '';
        document.getElementById('uploadEditSuffix').value = '';
        document.getElementById('uploadEditAddress1').value = addr.address || '';
        document.getElementById('uploadEditAddress2').value = '';
        document.getElementById('uploadEditAddress2SameLine').checked = false;
        document.getElementById('uploadEditAddress2Option').style.display = 'none';
        document.getElementById('uploadEditCity').value = addr.cityStateZip || '';
        document.getElementById('uploadEditState').value = '';
        document.getElementById('uploadEditPostalCode').value = '';
        document.getElementById('uploadEditCountry').value = '';
    }

    // Update form labels for region
    updateUploadEditFormLabels();

    // Scroll to form and focus
    editForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
    document.getElementById('uploadEditFirst').focus();

    showStatus('Editing address. Click "Save Address" when done.', 'success');
}

function updateUploadEditFormLabels() {
    const isUS = currentRegion === 'US';
    const config = isUS ? {
        stateLabel: 'State',
        statePlaceholder: 'NY',
        postalLabel: 'ZIP Code *',
        postalPlaceholder: '10001',
        cityPlaceholder: 'New York',
        countryPlaceholder: 'USA',
        address2Hint: 'Per USPS, some addresses require this'
    } : {
        stateLabel: 'County/Region',
        statePlaceholder: 'Greater London',
        postalLabel: 'Postcode *',
        postalPlaceholder: 'SW1A 1AA',
        cityPlaceholder: 'London',
        countryPlaceholder: 'United Kingdom',
        address2Hint: 'Some addresses may require this on one line'
    };

    const stateLabel = document.getElementById('uploadEditStateLabel');
    const postalLabel = document.getElementById('uploadEditPostalLabel');
    const state = document.getElementById('uploadEditState');
    const postalCode = document.getElementById('uploadEditPostalCode');
    const city = document.getElementById('uploadEditCity');
    const country = document.getElementById('uploadEditCountry');
    const hint = document.getElementById('uploadEditAddress2Hint');

    if (stateLabel) stateLabel.textContent = config.stateLabel;
    if (postalLabel) postalLabel.textContent = config.postalLabel;
    if (state) state.placeholder = config.statePlaceholder;
    if (postalCode) postalCode.placeholder = config.postalPlaceholder;
    if (city) city.placeholder = config.cityPlaceholder;
    if (country) country.placeholder = config.countryPlaceholder;
    if (hint) hint.textContent = config.address2Hint;
}

function deleteUploadAddress(index) {
    if (previewAddresses.length === 1) {
        if (confirm('This will remove the last address and cancel the preview. Continue?')) {
            cancelPreview();
        }
        return;
    }
    previewAddresses.splice(index, 1);
    saveUploadedAddressesToStorage();
    renderUploadPreviewList();
}

function updateUploadFullSheetHint() {
    const template = TEMPLATES[currentTemplate];
    if (!template) return;
    const labelsPerPage = template.cols * template.rows;
    const hint = document.getElementById('uploadFullSheetHint');
    if (hint) {
        hint.textContent = `Creates ${labelsPerPage} labels of the same address (based on current template)`;
    }
}

function updateUploadPreviewStats() {
    if (previewAddresses.length === 0) return;

    const template = TEMPLATES[currentTemplate];
    if (!template) return;
    const labelsPerPage = template.cols * template.rows;
    const fullSheetCheckbox = document.getElementById('uploadFullSheetCheckbox');
    const skipCount = getSkipCount();

    let labelCount = previewAddresses.length;
    if (fullSheetCheckbox && fullSheetCheckbox.checked && previewAddresses.length === 1) {
        labelCount = labelsPerPage - skipCount; // Full sheet minus skipped
    }

    // Total positions needed = skipped + actual labels
    const totalPositions = skipCount + labelCount;
    const totalPages = Math.ceil(totalPositions / labelsPerPage);

    let labelText = `${labelCount} label${labelCount !== 1 ? 's' : ''}`;
    if (skipCount > 0) {
        labelText += ` (skip ${skipCount})`;
    }

    document.getElementById('previewCount').textContent = labelText;
    document.getElementById('previewPages').textContent = `${totalPages} page${totalPages !== 1 ? 's' : ''}`;
}

function cancelPreview() {
    previewAddresses = [];
    clearUploadedAddressesFromStorage();
    document.getElementById('previewSection').classList.remove('active');
    document.getElementById('previewList').innerHTML = '';
    document.getElementById('uploadFullSheetOption').style.display = 'none';
    document.getElementById('uploadFullSheetCheckbox').checked = false;
    fileInput.value = '';
    status.style.display = 'none';
}

async function generateFromPreview() {
    if (previewAddresses.length === 0) return;

    const template = TEMPLATES[currentTemplate];
    const labelsPerPage = template.cols * template.rows;
    const fullSheetCheckbox = document.getElementById('uploadFullSheetCheckbox');
    const skipCount = getSkipCount();

    let addressesToPrint = previewAddresses;

    // If full sheet is checked and there's exactly one address, replicate it
    if (fullSheetCheckbox.checked && previewAddresses.length === 1) {
        // Fill remaining positions after skip
        const fillCount = labelsPerPage - skipCount;
        addressesToPrint = Array(fillCount).fill(previewAddresses[0]);
    }

    // Prepend empty entries for skipped labels
    if (skipCount > 0) {
        addressesToPrint = [...Array(skipCount).fill(null), ...addressesToPrint];
    }

    const actualLabelCount = addressesToPrint.filter(a => a !== null).length;
    showStatus(`Generating PDF labels for ${actualLabelCount} labels (skipping ${skipCount})...`, 'processing');

    try {
        const pdf = await generateLabelsPDF(addressesToPrint);
        pdf.save(`${currentTemplate}-labels.pdf`);

        const totalPages = Math.ceil(addressesToPrint.length / labelsPerPage);
        let successMsg = `✓ Success! Generated ${actualLabelCount} labels (${totalPages} page${totalPages > 1 ? 's' : ''}).`;
        if (skipCount > 0) {
            successMsg += ` Skipped first ${skipCount} position${skipCount > 1 ? 's' : ''}.`;
        }
        successMsg += ' PDF download started.';
        showStatus(successMsg, 'success');

        // Clear preview after generating
        cancelPreview();
    } catch (error) {
        console.error(error);
        showStatus('Error generating labels: ' + error.message, 'error');
    }
}

// Manual entry functions
function saveAddress() {
    // Determine which form we're using based on editing source
    const prefix = (isEditing && editingSource === 'upload') ? 'uploadEdit' : 'manual';

    const title = document.getElementById(prefix + 'Title').value.trim();
    const firstName = document.getElementById(prefix + 'First').value.trim();
    const lastName = document.getElementById(prefix + 'Last').value.trim();
    const suffix = document.getElementById(prefix + 'Suffix').value.trim();
    const address1 = document.getElementById(prefix + 'Address1').value.trim();
    const address2 = document.getElementById(prefix + 'Address2').value.trim();
    const address2SameLine = document.getElementById(prefix + 'Address2SameLine').checked;
    const city = document.getElementById(prefix + 'City').value.trim();
    const state = document.getElementById(prefix + 'State').value.trim();
    const postalCode = document.getElementById(prefix + 'PostalCode').value.trim();
    const country = document.getElementById(prefix + 'Country').value.trim();

    // Validate required fields - State/County is optional (supports international addresses)
    if (!firstName || !lastName || !address1 || !city || !postalCode) {
        alert('Please fill in all required fields (marked with *)');
        return;
    }

    const addressData = {
        title, firstName, lastName, suffix,
        address1, address2, address2SameLine,
        city, state, postalCode, country
    };

    if (isEditing) {
        if (editingSource === 'upload') {
            // Update the uploaded address at the original index
            previewAddresses[editingIndex] = addressData;
            saveUploadedAddressesToStorage();
            renderUploadPreviewList();
            showStatus('Address updated successfully.', 'success');
        } else if (editingSource === 'manual') {
            // Update the manual address at the original index
            manualAddresses[editingIndex] = addressData;
            saveAddressesToStorage();
            updateAddressesList();
            showStatus('Address updated successfully.', 'success');
        }
    } else {
        // Adding new address to manual list
        manualAddresses.push(addressData);
        saveAddressesToStorage();
        updateAddressesList();
        showStatus('Address added successfully.', 'success');
    }

    cancelEdit();
}

function cancelEdit() {
    // Reset editing state
    isEditing = false;
    editingIndex = -1;
    editingSource = null;

    // Reset manual form UI
    document.getElementById('manualFormTitle').textContent = 'Add Address';
    document.getElementById('manualFormButton').textContent = 'Add Address';
    document.getElementById('manualClearButton').textContent = 'Clear';

    // Clear manual form
    document.getElementById('manualTitle').value = '';
    document.getElementById('manualFirst').value = '';
    document.getElementById('manualLast').value = '';
    document.getElementById('manualSuffix').value = '';
    document.getElementById('manualAddress1').value = '';
    document.getElementById('manualAddress2').value = '';
    document.getElementById('manualAddress2SameLine').checked = false;
    document.getElementById('address2Option').style.display = 'none';
    document.getElementById('manualCity').value = '';
    document.getElementById('manualState').value = '';
    document.getElementById('manualPostalCode').value = '';
    document.getElementById('manualCountry').value = '';

    // Hide and clear upload edit form
    const uploadEditForm = document.getElementById('uploadEditForm');
    if (uploadEditForm) {
        uploadEditForm.style.display = 'none';
        document.getElementById('uploadEditTitle').value = '';
        document.getElementById('uploadEditFirst').value = '';
        document.getElementById('uploadEditLast').value = '';
        document.getElementById('uploadEditSuffix').value = '';
        document.getElementById('uploadEditAddress1').value = '';
        document.getElementById('uploadEditAddress2').value = '';
        document.getElementById('uploadEditAddress2SameLine').checked = false;
        document.getElementById('uploadEditAddress2Option').style.display = 'none';
        document.getElementById('uploadEditCity').value = '';
        document.getElementById('uploadEditState').value = '';
        document.getElementById('uploadEditPostalCode').value = '';
        document.getElementById('uploadEditCountry').value = '';
    }
}

// Legacy function name for backwards compatibility
function clearForm() {
    cancelEdit();
}

// Helper to get display text for an address (handles both old and new formats)
function getAddressDisplayText(addr) {
    const formatted = formatAddressForLabel(addr);
    let addressDisplay = formatted.addressLine1;
    if (formatted.addressLine2) {
        if (formatted.address2SameLine) {
            addressDisplay += ' ' + formatted.addressLine2;
        } else {
            addressDisplay += '<br>' + formatted.addressLine2;
        }
    }
    return {
        name: formatted.fullname,
        address: addressDisplay,
        location: formatted.location
    };
}

function updateAddressesList() {
    const section = document.getElementById('manualAddressesSection');
    const list = document.getElementById('manualAddressesList');
    const fullSheetOption = document.getElementById('manualFullSheetOption');

    if (manualAddresses.length === 0) {
        section.classList.remove('active');
        return;
    }

    // Show the section
    section.classList.add('active');

    // Update stats
    updateManualStats();

    // Only show full sheet option when exactly 1 address
    if (manualAddresses.length === 1) {
        fullSheetOption.style.display = '';
        updateManualFullSheetHint();
    } else {
        fullSheetOption.style.display = 'none';
        document.getElementById('manualFullSheetCheckbox').checked = false;
    }

    // Render addresses like file upload preview
    list.innerHTML = manualAddresses.map((addr, index) => {
        const display = getAddressDisplayText(addr);
        return `
        <div class="preview-item" id="manual-item-${index}">
            <div class="preview-item-content">
                <div class="preview-item-number">Address #${index + 1}</div>
                <div class="preview-item-name">${escapeHtml(display.name)}</div>
                <div class="preview-item-details">${display.address}<br>${escapeHtml(display.location)}</div>
            </div>
            <div class="preview-item-actions">
                <button class="btn-edit" onclick="editManualAddress(${index})">Edit</button>
                <button class="btn-delete" onclick="deleteManualAddress(${index})">Delete</button>
            </div>
        </div>
    `}).join('');
}

function updateManualStats() {
    const template = TEMPLATES[currentTemplate];
    if (!template) return;
    const labelsPerPage = template.cols * template.rows;
    const fullSheetCheckbox = document.getElementById('manualFullSheetCheckbox');
    const skipCount = getSkipCount();

    let labelCount = manualAddresses.length;
    if (fullSheetCheckbox && fullSheetCheckbox.checked && manualAddresses.length === 1) {
        labelCount = labelsPerPage - skipCount; // Full sheet minus skipped
    }

    // Total positions needed = skipped + actual labels
    const totalPositions = skipCount + labelCount;
    const totalPages = Math.ceil(totalPositions / labelsPerPage);

    let labelText = `${labelCount} label${labelCount !== 1 ? 's' : ''}`;
    if (skipCount > 0) {
        labelText += ` (skip ${skipCount})`;
    }

    document.getElementById('manualAddressCount').textContent = labelText;
    document.getElementById('manualPageCount').textContent = `${totalPages} page${totalPages !== 1 ? 's' : ''}`;
}

function updateManualFullSheetHint() {
    const template = TEMPLATES[currentTemplate];
    if (!template) return;
    const labelsPerPage = template.cols * template.rows;
    const hint = document.getElementById('manualFullSheetHint');
    if (hint) {
        hint.textContent = `Creates ${labelsPerPage} labels of the same address (based on current template)`;
    }
}

// Edit loads address into the form for modification
function editManualAddress(index) {
    const addr = manualAddresses[index];

    // Set editing state
    isEditing = true;
    editingIndex = index;
    editingSource = 'manual';

    // Update form UI to show editing mode
    document.getElementById('manualFormTitle').textContent = 'Edit Address';
    document.getElementById('manualFormButton').textContent = 'Save Address';
    document.getElementById('manualClearButton').textContent = 'Cancel';

    // Handle both old and new format
    if (addr.firstName !== undefined) {
        // New format
        document.getElementById('manualTitle').value = addr.title || '';
        document.getElementById('manualFirst').value = addr.firstName || '';
        document.getElementById('manualLast').value = addr.lastName || '';
        document.getElementById('manualSuffix').value = addr.suffix || '';
        document.getElementById('manualAddress1').value = addr.address1 || '';
        document.getElementById('manualAddress2').value = addr.address2 || '';
        document.getElementById('manualAddress2SameLine').checked = addr.address2SameLine || false;
        document.getElementById('address2Option').style.display = addr.address2 ? '' : 'none';
        document.getElementById('manualCity').value = addr.city || '';
        document.getElementById('manualState').value = addr.state || '';
        document.getElementById('manualPostalCode').value = addr.postalCode || '';
        document.getElementById('manualCountry').value = addr.country || '';
    } else {
        // Old format - parse as best we can
        const nameParts = (addr.fullname || '').split(' ');
        document.getElementById('manualTitle').value = '';
        document.getElementById('manualFirst').value = nameParts[0] || '';
        document.getElementById('manualLast').value = nameParts.slice(1).join(' ') || '';
        document.getElementById('manualSuffix').value = '';
        document.getElementById('manualAddress1').value = addr.address || '';
        document.getElementById('manualAddress2').value = '';
        document.getElementById('manualAddress2SameLine').checked = false;
        document.getElementById('address2Option').style.display = 'none';
        // Parse cityStateZip - try to split
        const csz = addr.cityStateZip || '';
        document.getElementById('manualCity').value = csz;
        document.getElementById('manualState').value = '';
        document.getElementById('manualPostalCode').value = '';
        document.getElementById('manualCountry').value = '';
    }

    // Scroll to form
    document.getElementById('manualSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
    document.getElementById('manualFirst').focus();

    showStatus('Editing address. Click "Save Address" when done.', 'success');
}

function deleteManualAddress(index) {
    if (manualAddresses.length === 1) {
        if (confirm('This will remove the last address. Continue?')) {
            manualAddresses = [];
            saveAddressesToStorage();
            updateAddressesList();
        }
        return;
    }
    manualAddresses.splice(index, 1);
    saveAddressesToStorage();
    updateAddressesList();
}

async function generateFromManual() {
    if (manualAddresses.length === 0) return;

    const template = TEMPLATES[currentTemplate];
    const labelsPerPage = template.cols * template.rows;
    const fullSheetCheckbox = document.getElementById('manualFullSheetCheckbox');
    const skipCount = getSkipCount();

    let addressesToPrint = manualAddresses;

    // If full sheet is checked and there's exactly one address, replicate it
    if (fullSheetCheckbox && fullSheetCheckbox.checked && manualAddresses.length === 1) {
        // Fill remaining positions after skip
        const fillCount = labelsPerPage - skipCount;
        addressesToPrint = Array(fillCount).fill(manualAddresses[0]);
    }

    // Prepend empty entries for skipped labels
    if (skipCount > 0) {
        addressesToPrint = [...Array(skipCount).fill(null), ...addressesToPrint];
    }

    const actualLabelCount = addressesToPrint.filter(a => a !== null).length;
    showStatus(`Generating PDF labels for ${actualLabelCount} labels (skipping ${skipCount})...`, 'processing');

    try {
        const pdf = await generateLabelsPDF(addressesToPrint);
        pdf.save(`${currentTemplate}-labels.pdf`);

        const totalPages = Math.ceil(addressesToPrint.length / labelsPerPage);
        let successMsg = `✓ Success! Generated ${actualLabelCount} labels (${totalPages} page${totalPages > 1 ? 's' : ''}).`;
        if (skipCount > 0) {
            successMsg += ` Skipped first ${skipCount} position${skipCount > 1 ? 's' : ''}.`;
        }
        successMsg += ' PDF download started.';
        showStatus(successMsg, 'success');
    } catch (error) {
        console.error(error);
        showStatus('Error generating labels: ' + error.message, 'error');
    }
}

// Sample address data - expanded format (US)
const US_SAMPLE_ADDRESSES = [
    ['Title', 'First Name', 'Last Name', 'Suffix', 'Address 1', 'Address 2', 'City', 'State', 'Postal Code', 'Country'],
    ['', 'John', 'Smith', '', '123 Main St', '', 'New York', 'NY', '10001', 'USA'],
    ['Dr.', 'Jane', 'Doe', 'PhD', '456 Oak Ave', 'Suite 200', 'Boston', 'MA', '02101', 'USA'],
    ['', 'Robert', 'Johnson', '', '789 Pine Rd', '', 'Chicago', 'IL', '60601', 'USA'],
    ['', 'Maria', 'Garcia', '', '321 Maple Dr', 'Apt 4B', 'Los Angeles', 'CA', '90001', 'USA'],
    ['Hon.', 'David', 'Lee', 'Jr.', '654 Cedar Ln', '', 'Houston', 'TX', '77001', 'USA'],
    ['', 'Sarah', 'Wilson', '', '987 Birch St', '', 'Phoenix', 'AZ', '85001', 'USA'],
    ['', 'Michael', 'Brown', '', '147 Elm Ave', 'Unit 12', 'Philadelphia', 'PA', '19101', 'USA'],
    ['', 'Emily', 'Davis', '', '258 Walnut Blvd', '', 'San Antonio', 'TX', '78201', 'USA'],
    ['', 'Christopher', 'Martinez', '', '369 Ash Ct', '', 'San Diego', 'CA', '92101', 'USA'],
    ['', 'Jessica', 'Anderson', '', '741 Oak Circle', '', 'Dallas', 'TX', '75201', 'USA']
];

// Sample address data - expanded format (UK/EU)
const UK_SAMPLE_ADDRESSES = [
    ['Title', 'First Name', 'Last Name', 'Suffix', 'Address 1', 'Address 2', 'City', 'County/Region', 'Postcode', 'Country'],
    ['', 'James', 'Smith', '', '10 Downing Street', '', 'London', '', 'SW1A 2AA', 'United Kingdom'],
    ['Dr.', 'Emma', 'Williams', 'PhD', '221B Baker Street', 'Flat 2', 'London', '', 'NW1 6XE', 'United Kingdom'],
    ['', 'Oliver', 'Brown', '', '45 High Street', '', 'Edinburgh', 'Scotland', 'EH1 1SR', 'United Kingdom'],
    ['', 'Sophie', 'Jones', '', '78 Queen Street', 'Suite 3', 'Cardiff', 'Wales', 'CF10 2GP', 'United Kingdom'],
    ['Mr.', 'William', 'Taylor', '', '23 Victoria Road', '', 'Manchester', 'Greater Manchester', 'M1 2HN', 'United Kingdom'],
    ['', 'Charlotte', 'Davies', '', '156 Castle Street', '', 'Belfast', 'Northern Ireland', 'BT1 1HE', 'United Kingdom'],
    ['', 'Harry', 'Wilson', '', '89 Church Lane', 'Unit 4', 'Birmingham', 'West Midlands', 'B1 1AA', 'United Kingdom'],
    ['', 'Amelia', 'Evans', '', '34 Park Avenue', '', 'Glasgow', 'Scotland', 'G1 1XQ', 'United Kingdom'],
    ['', 'George', 'Thomas', '', '67 Market Place', '', 'Bristol', '', 'BS1 3AE', 'United Kingdom'],
    ['', 'Isla', 'Roberts', '', '12 Abbey Road', '', 'Liverpool', 'Merseyside', 'L1 9AA', 'United Kingdom']
];

// Get sample addresses based on current region
function getSampleAddresses() {
    return currentRegion === 'US' ? US_SAMPLE_ADDRESSES : UK_SAMPLE_ADDRESSES;
}

// Update format example based on region
function updateFormatExample() {
    const formatExample = document.getElementById('formatExample');
    if (!formatExample) return;

    if (currentRegion === 'US') {
        formatExample.innerHTML = `
            <strong>Expanded format (recommended):</strong><br>
            Title,First Name,Last Name,Suffix,Address 1,Address 2,City,State,Postal Code,Country<br>
            Dr.,Jane,Doe,PhD,456 Oak Ave,Suite 200,Boston,MA,02101,USA<br><br>
            <strong>Simple format (also works):</strong><br>
            Fullname,Address,City State Zip<br>
            John Smith,123 Main St,New York NY 10001
        `;
    } else {
        formatExample.innerHTML = `
            <strong>Expanded format (recommended):</strong><br>
            Title,First Name,Last Name,Suffix,Address 1,Address 2,City,County/Region,Postcode,Country<br>
            Dr.,Emma,Williams,PhD,221B Baker Street,Flat 2,London,,NW1 6XE,United Kingdom<br><br>
            <strong>Simple format (also works):</strong><br>
            Fullname,Address,City Postcode<br>
            James Smith,10 Downing Street,London SW1A 2AA
        `;
    }
}

// Helper function to download addresses as CSV
function downloadAddressesCSV(addresses, filename) {
    const header = ['Title', 'First Name', 'Last Name', 'Suffix', 'Address 1', 'Address 2', 'City', 'State', 'Postal Code', 'Country'];

    const escapeCsvField = (field) => {
        const str = field || '';
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
    };

    const rows = addresses.map(addr => {
        // Handle both old and new format
        if (addr.firstName !== undefined) {
            // New format
            return [
                escapeCsvField(addr.title),
                escapeCsvField(addr.firstName),
                escapeCsvField(addr.lastName),
                escapeCsvField(addr.suffix),
                escapeCsvField(addr.address1),
                escapeCsvField(addr.address2),
                escapeCsvField(addr.city),
                escapeCsvField(addr.state),
                escapeCsvField(addr.postalCode),
                escapeCsvField(addr.country)
            ].join(',');
        } else {
            // Old format - convert to new header format as best we can
            const nameParts = (addr.fullname || '').split(' ');
            return [
                '', // title
                escapeCsvField(nameParts[0] || ''),
                escapeCsvField(nameParts.slice(1).join(' ') || ''),
                '', // suffix
                escapeCsvField(addr.address),
                '', // address2
                escapeCsvField(addr.cityStateZip),
                '', // state
                '', // postal code
                ''  // country
            ].join(',');
        }
    });

    const csvContent = [header.join(','), ...rows].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);

    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Export manual addresses to CSV
function exportManualAddressesCSV() {
    if (manualAddresses.length === 0) return;
    downloadAddressesCSV(manualAddresses, 'my-addresses.csv');
    showStatus(`Exported ${manualAddresses.length} address${manualAddresses.length !== 1 ? 'es' : ''} to CSV.`, 'success');
}

// Export uploaded addresses to CSV
function exportUploadedAddressesCSV() {
    if (previewAddresses.length === 0) return;
    downloadAddressesCSV(previewAddresses, 'updated-addresses.csv');
    showStatus(`Exported ${previewAddresses.length} address${previewAddresses.length !== 1 ? 'es' : ''} to CSV.`, 'success');
}

// Download example CSV file
function downloadExampleCSV() {
    const csvContent = getSampleAddresses().map(row => row.join(',')).join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);

    link.setAttribute('href', url);
    link.setAttribute('download', 'example-addresses.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Download example Excel file
function downloadExampleExcel() {
    // Create a new workbook
    const wb = XLSX.utils.book_new();

    // Create worksheet from array data
    const ws = XLSX.utils.aoa_to_sheet(getSampleAddresses());

    // Set column widths for new 10-column format
    ws['!cols'] = [
        { wch: 8 },   // Title
        { wch: 15 },  // First Name
        { wch: 15 },  // Last Name
        { wch: 8 },   // Suffix
        { wch: 25 },  // Address 1
        { wch: 15 },  // Address 2
        { wch: 15 },  // City
        { wch: 8 },   // State
        { wch: 12 },  // Postal Code
        { wch: 10 }   // Country
    ];

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Addresses');

    // Generate Excel file and trigger download
    XLSX.writeFile(wb, 'example-addresses.xlsx');
}
