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
});

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
});

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

    // State/County is optional (supports international addresses)
    if (!firstName || !lastName || !address1 || !city || !postalCode) {
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
function generateReturnLabels() {
    const returnAddress = getReturnAddressFromForm();
    if (!returnAddress) {
        alert('Please fill in all required fields (marked with *)');
        return;
    }

    // Save for next time
    saveReturnAddressToStorage();

    const template = TEMPLATES[currentTemplate];
    const labelsPerPage = template.cols * template.rows;

    // Fill entire sheet with the same address
    const addressesToPrint = Array(labelsPerPage).fill(returnAddress);

    showStatus(`Generating ${labelsPerPage} return address labels...`, 'processing');

    try {
        const pdf = generateLabelsPDF(addressesToPrint);
        pdf.save(`${currentTemplate}-return-labels.pdf`);

        showStatus(`✓ Success! Generated ${labelsPerPage} return address labels (1 page). PDF download started.`, 'success');
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

function drawLabel(pdf, addressData, x, y, template) {
    if (!addressData) return;

    const formatted = formatAddressForLabel(addressData);

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
    const maxWidth = template.width - (padding * 2);
    const maxHeight = template.height - (padding * 2) - startPadding;

    // Auto-fit: Find optimal font size that fits all content
    let fontSize = baseFontSize;
    let totalHeight;

    do {
        pdf.setFontSize(fontSize);
        const lineSpacing = (fontSize / 72) * lineSpacingRatio; // Convert pt to inches
        totalHeight = 0;

        for (const line of filteredLines) {
            pdf.setFont("helvetica", line.isBold ? "bold" : "normal");
            const wrappedLines = pdf.splitTextToSize(line.text, maxWidth);
            totalHeight += wrappedLines.length * lineSpacing;
        }

        if (totalHeight <= maxHeight) break;
        fontSize -= 0.5;
    } while (fontSize >= minFontSize);

    // Now draw the text
    const startX = x + padding;
    let currentY = y + padding + startPadding;
    const lineSpacing = (fontSize / 72) * lineSpacingRatio;

    for (const line of filteredLines) {
        pdf.setFont("helvetica", line.isBold ? "bold" : "normal");
        pdf.setFontSize(line.isBold ? fontSize + 1 : fontSize);

        const wrappedLines = pdf.splitTextToSize(line.text, maxWidth);
        for (const wrappedLine of wrappedLines) {
            if (currentY > y + template.height - padding) break;
            pdf.text(wrappedLine, startX, currentY);
            currentY += lineSpacing;
        }
    }
}

function generateLabelsPDF(addresses) {
    const template = TEMPLATES[currentTemplate];
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'in',
        format: currentRegion === 'US' ? 'letter' : 'a4'
    });

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

                drawLabel(pdf, addressData, x, y, template);
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

    let labelCount = previewAddresses.length;
    if (fullSheetCheckbox && fullSheetCheckbox.checked && previewAddresses.length === 1) {
        labelCount = labelsPerPage;
    }

    const totalPages = Math.ceil(labelCount / labelsPerPage);

    document.getElementById('previewCount').textContent = `${labelCount} label${labelCount !== 1 ? 's' : ''}`;
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

function generateFromPreview() {
    if (previewAddresses.length === 0) return;

    const template = TEMPLATES[currentTemplate];
    const labelsPerPage = template.cols * template.rows;
    const fullSheetCheckbox = document.getElementById('uploadFullSheetCheckbox');

    let addressesToPrint = previewAddresses;

    // If full sheet is checked and there's exactly one address, replicate it
    if (fullSheetCheckbox.checked && previewAddresses.length === 1) {
        addressesToPrint = Array(labelsPerPage).fill(previewAddresses[0]);
    }

    showStatus(`Generating PDF labels for ${addressesToPrint.length} labels...`, 'processing');

    try {
        const pdf = generateLabelsPDF(addressesToPrint);
        pdf.save(`${currentTemplate}-labels.pdf`);

        const totalPages = Math.ceil(addressesToPrint.length / labelsPerPage);
        showStatus(`✓ Success! Generated ${addressesToPrint.length} labels (${totalPages} page${totalPages > 1 ? 's' : ''}). PDF download started.`, 'success');

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

    let labelCount = manualAddresses.length;
    if (fullSheetCheckbox && fullSheetCheckbox.checked && manualAddresses.length === 1) {
        labelCount = labelsPerPage;
    }

    const totalPages = Math.ceil(labelCount / labelsPerPage);

    document.getElementById('manualAddressCount').textContent = `${labelCount} label${labelCount !== 1 ? 's' : ''}`;
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

function generateFromManual() {
    if (manualAddresses.length === 0) return;

    const template = TEMPLATES[currentTemplate];
    const labelsPerPage = template.cols * template.rows;
    const fullSheetCheckbox = document.getElementById('manualFullSheetCheckbox');

    let addressesToPrint = manualAddresses;

    // If full sheet is checked and there's exactly one address, replicate it
    if (fullSheetCheckbox && fullSheetCheckbox.checked && manualAddresses.length === 1) {
        addressesToPrint = Array(labelsPerPage).fill(manualAddresses[0]);
    }

    showStatus(`Generating PDF labels for ${addressesToPrint.length} labels...`, 'processing');

    try {
        const pdf = generateLabelsPDF(addressesToPrint);
        pdf.save(`${currentTemplate}-labels.pdf`);

        const totalPages = Math.ceil(addressesToPrint.length / labelsPerPage);
        showStatus(`✓ Success! Generated ${addressesToPrint.length} labels (${totalPages} page${totalPages > 1 ? 's' : ''}). PDF download started.`, 'success');
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
