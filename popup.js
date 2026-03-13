const SHEET_ID = '1jjzb4CUl_9iJ9Hlgov7tqqifrRJPojTGkCItJ22PSTk';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`;
const PARTNER_SHEET_NAME = 'Doi_Tac_Giao_Hang';
const SUPPLIER_SHEET_NAME = 'Nha_Cung_Cap';
const PRODUCT_SHEET_NAME = 'San_Pham';
const PARTNER_SHEET_GID = '0';
const SUPPLIER_SHEET_GID = '1394006768';
const PRODUCT_SHEET_GID = '1330714216';
const PARTNER_SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${PARTNER_SHEET_GID}`;
const SUPPLIER_SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${SUPPLIER_SHEET_GID}`;
const PRODUCT_SHEET_CSV_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${PRODUCT_SHEET_GID}`;
const PARTNER_CSV_PATH = 'doitacgiaohang.csv';
const SUPPLIER_CSV_PATH = 'nhacungcap.csv';
const PRODUCT_CSV_PATH = 'sanpham.csv';
const PARTNER_CACHE_KEY = 'partnerCacheV3';
const SUPPLIER_CACHE_KEY = 'supplierCacheV3';
const PRODUCT_CACHE_KEY = 'productCacheV1';
const CASES_WITH_PARTNER_FIRST = new Set(['case1', 'case2']);
const PARTNER_SUGGESTION_LIMIT = 80;
const SUPPLIER_SUGGESTION_LIMIT = 80;
const PRODUCT_SUGGESTION_LIMIT = 80;
const ACTIVATION_OPTIONS = [
    'Còn Kích',
    'Hết Kích',
    'Còn Kích, còn Urbox',
    'Còn Kích, hết Urbox',
    'Hết Kích, còn Urbox',
    'Hết Kích, hết Urbox',
    'Còn Vip/Tích Lũy, còn Kích',
    'Hết Vip/Tích Lũy, còn Kích',
    'Hết Vip/Tích Lũy, hết Kích',
    'Không rõ tình trạng Kích'
];
const CASE_COPY_LABELS = {
    case1: 'Lấy NCC giao khách',
    case2: 'Lấy NCC về kho',
    case3: 'NCC giao về kho',
    case4: 'NCC giao khách'
};

let partnerRecords = [];
let partnerIndex = {
    byLabel: new Map(),
    byCode: new Map(),
    byName: new Map()
};

let supplierRecords = [];
let supplierIndex = {
    byLabel: new Map(),
    byCode: new Map(),
    byName: new Map()
};

let productRecords = [];
let productIndex = {
    byCode: new Map()
};

const statusState = {
    partner: 'Chưa tải dữ liệu đối tác',
    supplier: 'Chưa tải dữ liệu nhà cung cấp',
    product: 'Chưa tải dữ liệu mã sản phẩm'
};

document.addEventListener('DOMContentLoaded', () => {
    const elements = {
        caseSelect: document.getElementById('caseSelect'),
        partnerInput: document.getElementById('partnerInput'),
        partnerList: document.getElementById('partnerList'),
        supplierInput: document.getElementById('supplierInput'),
        supplierList: document.getElementById('supplierList'),
        productStatusList: document.getElementById('productStatusList'),
        productCodeList: document.getElementById('productCodeList'),
        shippingFee: document.getElementById('shippingFee'),
        note: document.getElementById('note'),
        copyButton: document.getElementById('copyButton'),
        resetButton: document.getElementById('resetButton'),
        restoreInput: document.getElementById('restoreInput'),
        restoreButton: document.getElementById('restoreButton'),
        restoreStatus: document.getElementById('restoreStatus'),
        addProductStatusButton: document.getElementById('addProductStatus'),
        refreshButton: document.getElementById('refreshPartners'),
        openSheetButton: document.getElementById('openSheet'),
        dataStatus: document.getElementById('dataStatus')
    };

    setStatus(elements, 'partner', statusState.partner);
    setStatus(elements, 'supplier', statusState.supplier);
    setStatus(elements, 'product', statusState.product);

    elements.partnerInput.addEventListener('input', () => handlePartnerInput(elements));
    elements.partnerInput.addEventListener('focus', () => renderPartnerOptions(elements.partnerList, partnerRecords, elements.partnerInput.value));
    elements.supplierInput.addEventListener('input', () => handleSupplierInput(elements));
    elements.supplierInput.addEventListener('focus', () => renderSupplierOptions(elements.supplierList, supplierRecords, elements.supplierInput.value));
    elements.copyButton.addEventListener('click', () => copyText(elements));
    elements.resetButton.addEventListener('click', () => resetForm(elements));
    elements.restoreButton.addEventListener('click', () => restoreFromText(elements));
    elements.addProductStatusButton.addEventListener('click', () => addProductStatusRow(elements, { focusCode: true }));
    elements.refreshButton.addEventListener('click', () => refreshAllData(elements));
    elements.openSheetButton.addEventListener('click', openSheet);
    elements.productStatusList.addEventListener('input', (event) => handleProductStatusInput(event, elements));
    elements.productStatusList.addEventListener('focusin', (event) => handleProductStatusFocus(event, elements));
    elements.productStatusList.addEventListener('click', (event) => handleProductStatusClick(event, elements));

    addProductStatusRow(elements);
    setupShippingFeeInput(elements.shippingFee);
    loadAllData(elements);
});

function handlePartnerInput(elements) {
    renderPartnerOptions(elements.partnerList, partnerRecords, elements.partnerInput.value);
}

function handleSupplierInput(elements) {
    renderSupplierOptions(elements.supplierList, supplierRecords, elements.supplierInput.value);
}

function handleProductStatusInput(event, elements) {
    if (!event.target.classList.contains('product-code-input')) {
        return;
    }

    renderProductOptions(elements.productCodeList, productRecords, event.target.value);
}

function handleProductStatusFocus(event, elements) {
    if (!event.target.classList.contains('product-code-input')) {
        return;
    }

    renderProductOptions(elements.productCodeList, productRecords, event.target.value);
}

function handleProductStatusClick(event, elements) {
    const removeButton = event.target.closest('.product-status-remove');
    if (!removeButton) {
        return;
    }

    const row = removeButton.closest('.product-status-row');
    if (!row) {
        return;
    }

    row.remove();
    ensureAtLeastOneProductStatusRow(elements);
    updateProductStatusRowActions(elements);
}

function addProductStatusRow(elements, options = {}) {
    const row = document.createElement('div');
    row.className = 'product-status-row';
    row.innerHTML = `
        <div class="input-stack">
            <input type="text" class="product-code-input" list="productCodeList" placeholder="Chọn mã sản phẩm">
        </div>
        <div class="input-stack">
            <select class="product-status-select">
                ${buildActivationOptionsMarkup()}
            </select>
        </div>
        <button type="button" class="mini-ghost-button product-status-remove">Xóa</button>
    `;

    elements.productStatusList.appendChild(row);
    updateProductStatusRowActions(elements);

    if (options.focusCode) {
        const codeInput = row.querySelector('.product-code-input');
        if (codeInput) {
            codeInput.focus();
            renderProductOptions(elements.productCodeList, productRecords, codeInput.value);
        }
    }

    return row;
}

function ensureAtLeastOneProductStatusRow(elements) {
    if (elements.productStatusList.querySelector('.product-status-row')) {
        return;
    }

    addProductStatusRow(elements);
}

function updateProductStatusRowActions(elements) {
    const rows = Array.from(elements.productStatusList.querySelectorAll('.product-status-row'));
    const shouldHideRemove = rows.length === 1;

    rows.forEach((row) => {
        const removeButton = row.querySelector('.product-status-remove');
        if (!removeButton) {
            return;
        }
        removeButton.hidden = shouldHideRemove;
    });
}

function resetForm(elements) {
    elements.caseSelect.value = '';
    elements.partnerInput.value = '';
    elements.supplierInput.value = '';
    elements.shippingFee.value = '';
    elements.note.value = '';
    elements.productStatusList.innerHTML = '';
    setRestoreStatus(elements, '');

    addProductStatusRow(elements);
    renderPartnerOptions(elements.partnerList, partnerRecords, '');
    renderSupplierOptions(elements.supplierList, supplierRecords, '');
    renderProductOptions(elements.productCodeList, productRecords, '');
    elements.caseSelect.focus();
}

function restoreFromText(elements) {
    const rawText = elements.restoreInput.value.trim();
    if (!rawText) {
        setRestoreStatus(elements, 'Vui lòng paste note cũ cần xử lý.');
        return;
    }

    const restored = parseRestoredText(rawText);
    if (!restored.caseValue) {
        setRestoreStatus(elements, 'Không nhận diện được Trường hợp từ note cũ.');
        return;
    }

    if (!restored.productEntries.length) {
        setRestoreStatus(elements, 'Không nhận diện được mã sản phẩm và trạng thái.');
        return;
    }

    resetForm(elements);

    elements.caseSelect.value = restored.caseValue;
    elements.partnerInput.value = restored.partnerValue;
    elements.supplierInput.value = restored.supplierValue;
    elements.shippingFee.value = formatShippingFeeForInput(restored.shippingFeeDigits);
    elements.note.value = restored.note;

    const existingRow = elements.productStatusList.querySelector('.product-status-row');
    restored.productEntries.forEach((entry, index) => {
        const row = index === 0
            ? existingRow
            : addProductStatusRow(elements);
        if (!row) {
            return;
        }

        const codeInput = row.querySelector('.product-code-input');
        const statusSelect = row.querySelector('.product-status-select');
        if (codeInput) {
            codeInput.value = entry.productCode;
        }
        if (statusSelect) {
            statusSelect.value = entry.status;
        }
    });

    updateProductStatusRowActions(elements);
    setRestoreStatus(elements, `Đã điền lại ${restored.productEntries.length} sản phẩm.`);
}

function setRestoreStatus(elements, message) {
    if (!elements.restoreStatus) {
        return;
    }

    elements.restoreStatus.textContent = message || '';
}

function parseRestoredText(rawText) {
    const lines = rawText
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);

    if (!lines.length) {
        return {
            caseValue: '',
            partnerValue: '',
            supplierValue: '',
            productEntries: [],
            shippingFeeDigits: '',
            note: ''
        };
    }

    const context = parseRestoredContextLine(lines[0]);
    const productEntries = [];
    let shippingFeeDigits = '';
    const noteParts = [];

    lines.slice(1).forEach((line) => {
        const productEntry = parseRestoredProductLine(line);
        if (productEntry) {
            productEntries.push(productEntry);
            return;
        }

        const footer = parseRestoredFooterLine(line);
        if (footer.shippingFeeDigits) {
            shippingFeeDigits = footer.shippingFeeDigits;
        }
        if (footer.note) {
            noteParts.push(footer.note);
        }
    });

    return {
        ...context,
        productEntries,
        shippingFeeDigits,
        note: noteParts.join('\n')
    };
}

function parseRestoredContextLine(line) {
    const tokens = line
        .split('|')
        .map((token) => token.trim())
        .filter(Boolean);

    let caseValue = '';
    let caseIndex = -1;

    tokens.some((token, index) => {
        const matchedCaseValue = resolveCaseValueFromLabel(token);
        if (!matchedCaseValue) {
            return false;
        }

        caseValue = matchedCaseValue;
        caseIndex = index;
        return true;
    });

    if (!caseValue) {
        return {
            caseValue: '',
            partnerValue: '',
            supplierValue: ''
        };
    }

    let contextTokens = [];
    if (caseIndex === 0) {
        contextTokens = tokens.slice(1);
    } else if (caseIndex === tokens.length - 1) {
        contextTokens = tokens.slice(0, -1);
    } else {
        contextTokens = tokens.filter((_, index) => index !== caseIndex);
    }

    const resolvedContext = resolveRestoredContextTokens(contextTokens);
    return {
        caseValue,
        ...resolvedContext
    };
}

function resolveCaseValueFromLabel(value) {
    const normalized = normalizePartnerValue(value);

    const matchedEntry = Object.entries(CASE_COPY_LABELS).find(([, label]) => normalizePartnerValue(label) === normalized);
    return matchedEntry ? matchedEntry[0] : '';
}

function resolveRestoredContextTokens(tokens) {
    const sanitizedTokens = tokens.map((token) => token.trim()).filter(Boolean);
    if (!sanitizedTokens.length) {
        return {
            partnerValue: '',
            supplierValue: ''
        };
    }

    if (sanitizedTokens.length >= 2) {
        return {
            partnerValue: getExactPartnerLabel(sanitizedTokens[0]) || sanitizedTokens[0],
            supplierValue: getExactSupplierLabel(sanitizedTokens[1]) || sanitizedTokens[1]
        };
    }

    const singleToken = sanitizedTokens[0];
    const exactPartner = getExactPartnerLabel(singleToken);
    const exactSupplier = getExactSupplierLabel(singleToken);

    if (exactPartner && !exactSupplier) {
        return {
            partnerValue: exactPartner,
            supplierValue: ''
        };
    }

    if (exactSupplier && !exactPartner) {
        return {
            partnerValue: '',
            supplierValue: exactSupplier
        };
    }

    if (singleToken.includes('+')) {
        return {
            partnerValue: exactPartner || singleToken,
            supplierValue: ''
        };
    }

    return {
        partnerValue: '',
        supplierValue: exactSupplier || singleToken
    };
}

function getExactPartnerLabel(inputValue) {
    const normalized = normalizePartnerValue(inputValue);
    if (!normalized) {
        return '';
    }

    if (partnerIndex.byLabel.has(normalized)) {
        return partnerIndex.byLabel.get(normalized);
    }

    if (partnerIndex.byCode.has(normalized)) {
        return partnerIndex.byCode.get(normalized);
    }

    if (partnerIndex.byName.has(normalized)) {
        return partnerIndex.byName.get(normalized);
    }

    return '';
}

function getExactSupplierLabel(inputValue) {
    const normalized = normalizeSupplierValue(inputValue);
    if (!normalized) {
        return '';
    }

    if (supplierIndex.byLabel.has(normalized)) {
        return supplierIndex.byLabel.get(normalized);
    }

    if (supplierIndex.byCode.has(normalized)) {
        return supplierIndex.byCode.get(normalized);
    }

    if (supplierIndex.byName.has(normalized)) {
        return supplierIndex.byName.get(normalized);
    }

    return '';
}

function parseRestoredProductLine(line) {
    const trimmedLine = line.trim();
    if (!trimmedLine) {
        return null;
    }

    const cleanedLine = trimmedLine.replace(/\|\s*$/, '').trim();
    const separatorIndex = cleanedLine.indexOf(' - ');
    if (separatorIndex === -1) {
        return null;
    }

    const productCodeRaw = cleanedLine.slice(0, separatorIndex).trim();
    const statusRaw = cleanedLine.slice(separatorIndex + 3).trim();
    const resolvedStatus = resolveActivationStatus(statusRaw);

    if (!productCodeRaw || !resolvedStatus) {
        return null;
    }

    return {
        productCode: resolveProductCode(productCodeRaw),
        status: resolvedStatus
    };
}

function resolveActivationStatus(inputValue) {
    const normalized = normalizePartnerValue(inputValue);
    const matchedStatus = ACTIVATION_OPTIONS.find((status) => normalizePartnerValue(status) === normalized);
    return matchedStatus || '';
}

function parseRestoredFooterLine(line) {
    const tokens = line
        .split('|')
        .map((token) => token.trim())
        .filter(Boolean);

    let shippingFeeDigits = '';
    const noteTokens = [];

    tokens.forEach((token) => {
        if (normalizePartnerValue(token).startsWith('cuoc:')) {
            shippingFeeDigits = token.replace(/[^\d]/g, '');
            return;
        }

        noteTokens.push(token);
    });

    return {
        shippingFeeDigits,
        note: noteTokens.join(' | ')
    };
}

function openSheet() {
    if (chrome?.tabs?.create) {
        chrome.tabs.create({ url: SHEET_URL });
        return;
    }

    window.open(SHEET_URL, '_blank', 'noopener');
}

function setStatus(elements, key, message) {
    statusState[key] = message;
    const timestamp = extractStatusTimestamp(statusState.partner)
        || extractStatusTimestamp(statusState.supplier)
        || extractStatusTimestamp(statusState.product);

    elements.dataStatus.textContent = timestamp
        ? `Dữ liệu: ${timestamp}`
        : 'Dữ liệu: chưa có';
}

function extractStatusTimestamp(message) {
    if (!message) {
        return '';
    }

    const match = message.match(/(\d{1,2}:\d{2}:\d{2}\s+\d{1,2}\/\d{1,2}\/\d{4})/);
    return match ? match[1] : '';
}

function loadAllData(elements) {
    loadPartnerData(elements);
    loadSupplierData(elements);
    loadProductData(elements);
}

function loadPartnerData(elements) {
    if (chrome?.storage?.local) {
        chrome.storage.local.get(PARTNER_CACHE_KEY, (result) => {
            const cache = result[PARTNER_CACHE_KEY];
            if (cache?.records?.length) {
                setPartnerRecords(cache.records, elements);
                setStatus(elements, 'partner', `Đã tải dữ liệu đối tác (cache ${formatTimestamp(cache.fetchedAt)})`);
                return;
            }

            fetchLocalPartnerCsv(elements);
        });
        return;
    }

    fetchLocalPartnerCsv(elements);
}

function loadSupplierData(elements) {
    if (chrome?.storage?.local) {
        chrome.storage.local.get(SUPPLIER_CACHE_KEY, (result) => {
            const cache = result[SUPPLIER_CACHE_KEY];
            if (cache?.records?.length) {
                setSupplierRecords(cache.records, elements);
                setStatus(elements, 'supplier', `Đã tải dữ liệu nhà cung cấp (cache ${formatTimestamp(cache.fetchedAt)})`);
                return;
            }

            fetchLocalSupplierCsv(elements);
        });
        return;
    }

    fetchLocalSupplierCsv(elements);
}

function loadProductData(elements) {
    if (chrome?.storage?.local) {
        chrome.storage.local.get(PRODUCT_CACHE_KEY, (result) => {
            const cache = result[PRODUCT_CACHE_KEY];
            if (cache?.records?.length) {
                setProductRecords(cache.records, elements);
                setStatus(elements, 'product', `Đã tải dữ liệu mã sản phẩm (cache ${formatTimestamp(cache.fetchedAt)})`);
                return;
            }

            fetchLocalProductCsv(elements);
        });
        return;
    }

    fetchLocalProductCsv(elements);
}

function refreshAllData(elements) {
    setStatus(elements, 'partner', 'Đang tải dữ liệu đối tác...');
    setStatus(elements, 'supplier', 'Đang tải dữ liệu nhà cung cấp...');
    setStatus(elements, 'product', 'Đang tải dữ liệu mã sản phẩm...');

    Promise.allSettled([
        refreshPartnerData(elements),
        refreshSupplierData(elements),
        refreshProductData(elements)
    ]).then(() => {
    });
}

function refreshPartnerData(elements) {
    return fetch(PARTNER_SHEET_CSV_URL, { cache: 'no-store' })
        .then((response) => {
            if (!response.ok) {
                throw new Error('Không thể tải dữ liệu đối tác');
            }
            return response.text();
        })
        .then((csvText) => {
            const records = parseCsvRecords(csvText, ['ma', 'madoitac', 'code'], normalizePartnerValue);
            if (!records.length) {
                throw new Error('Dữ liệu đối tác trống');
            }

            setPartnerRecords(records, elements);
            setStatus(elements, 'partner', `Đã cập nhật đối tác (${formatTimestamp(Date.now())})`);

            if (chrome?.storage?.local) {
                chrome.storage.local.set({
                    [PARTNER_CACHE_KEY]: {
                        fetchedAt: Date.now(),
                        records
                    }
                });
            }
        })
        .catch((error) => {
            console.error(error);
            setStatus(elements, 'partner', 'Tải dữ liệu đối tác thất bại');
        });
}

function refreshSupplierData(elements) {
    return fetch(SUPPLIER_SHEET_CSV_URL, { cache: 'no-store' })
        .then((response) => {
            if (!response.ok) {
                throw new Error('Không thể tải dữ liệu nhà cung cấp');
            }
            return response.text();
        })
        .then((csvText) => {
            const records = parseCsvRecords(
                csvText,
                ['ma', 'manhacungcap', 'code', 'ma nha cung cap', 'ma_nha_cung_cap'],
                normalizeSupplierValue
            );
            if (!records.length) {
                throw new Error('Dữ liệu nhà cung cấp trống');
            }

            setSupplierRecords(records, elements);
            setStatus(elements, 'supplier', `Đã cập nhật nhà cung cấp (${formatTimestamp(Date.now())})`);

            if (chrome?.storage?.local) {
                chrome.storage.local.set({
                    [SUPPLIER_CACHE_KEY]: {
                        fetchedAt: Date.now(),
                        records
                    }
                });
            }
        })
        .catch((error) => {
            console.error(error);
            setStatus(elements, 'supplier', 'Tải dữ liệu nhà cung cấp thất bại');
        });
}

function refreshProductData(elements) {
    return fetch(PRODUCT_SHEET_CSV_URL, { cache: 'no-store' })
        .then((response) => {
            if (!response.ok) {
                throw new Error('Không thể tải dữ liệu mã sản phẩm');
            }
            return response.text();
        })
        .then((csvText) => {
            const records = parseCsvValues(csvText, ['ma hang', 'ma san pham', 'product code'], normalizeProductValue);
            if (!records.length) {
                throw new Error('Dữ liệu mã sản phẩm trống');
            }

            setProductRecords(records, elements);
            setStatus(elements, 'product', `Đã cập nhật mã sản phẩm (${formatTimestamp(Date.now())})`);

            if (chrome?.storage?.local) {
                chrome.storage.local.set({
                    [PRODUCT_CACHE_KEY]: {
                        fetchedAt: Date.now(),
                        records
                    }
                });
            }
        })
        .catch((error) => {
            console.error(error);
            setStatus(elements, 'product', 'Tải dữ liệu mã sản phẩm thất bại');
        });
}

function fetchLocalPartnerCsv(elements) {
    fetch(PARTNER_CSV_PATH)
        .then((response) => {
            if (!response.ok) {
                throw new Error('Không tìm thấy file đối tác');
            }
            return response.text();
        })
        .then((csvText) => {
            const records = parseCsvRecords(csvText, ['ma', 'madoitac', 'code'], normalizePartnerValue);
            if (!records.length) {
                throw new Error('File đối tác trống');
            }

            setPartnerRecords(records, elements);
            setStatus(elements, 'partner', `Đã tải dữ liệu đối tác nội bộ (${formatTimestamp(Date.now())})`);
        })
        .catch((error) => {
            console.error(error);
            setStatus(elements, 'partner', 'Chưa có dữ liệu đối tác');
        });
}

function fetchLocalSupplierCsv(elements) {
    fetch(SUPPLIER_CSV_PATH)
        .then((response) => {
            if (!response.ok) {
                throw new Error('Không tìm thấy file nhà cung cấp');
            }
            return response.text();
        })
        .then((csvText) => {
            const records = parseCsvRecords(
                csvText,
                ['ma', 'manhacungcap', 'code', 'ma nha cung cap', 'ma_nha_cung_cap'],
                normalizeSupplierValue
            );
            if (!records.length) {
                throw new Error('File nhà cung cấp trống');
            }

            setSupplierRecords(records, elements);
            setStatus(elements, 'supplier', `Đã tải dữ liệu nhà cung cấp nội bộ (${formatTimestamp(Date.now())})`);
        })
        .catch((error) => {
            console.error(error);
            setStatus(elements, 'supplier', 'Chưa có dữ liệu nhà cung cấp');
        });
}

function fetchLocalProductCsv(elements) {
    fetch(PRODUCT_CSV_PATH)
        .then((response) => {
            if (!response.ok) {
                throw new Error('Không tìm thấy file mã sản phẩm');
            }
            return response.text();
        })
        .then((csvText) => {
            const records = parseCsvValues(csvText, ['ma hang', 'ma san pham', 'product code'], normalizeProductValue);
            if (!records.length) {
                throw new Error('File mã sản phẩm trống');
            }

            setProductRecords(records, elements);
            setStatus(elements, 'product', `Đã tải dữ liệu mã sản phẩm nội bộ (${formatTimestamp(Date.now())})`);
        })
        .catch((error) => {
            console.error(error);
            setStatus(elements, 'product', 'Chưa có dữ liệu mã sản phẩm');
        });
}

function setPartnerRecords(records, elements) {
    partnerRecords = sortRecordsByLabel(records, 'name').map((record) => ({
        ...record,
        normalizedCode: normalizePartnerValue(record.code),
        normalizedName: normalizePartnerValue(record.name),
        tokens: splitPartnerTokens(record.name),
        label: record.name,
        labelNormalized: normalizePartnerValue(record.name)
    }));

    buildPartnerIndex(partnerRecords);
    renderPartnerOptions(elements.partnerList, partnerRecords, elements.partnerInput.value);
}

function buildPartnerIndex(records) {
    partnerIndex = {
        byLabel: new Map(),
        byCode: new Map(),
        byName: new Map()
    };

    records.forEach((record) => {
        const labelKey = normalizePartnerValue(record.label);
        const codeKey = normalizePartnerValue(record.code);
        const nameKey = normalizePartnerValue(record.name);

        partnerIndex.byLabel.set(labelKey, record.label);
        partnerIndex.byCode.set(codeKey, record.label);
        if (!partnerIndex.byName.has(nameKey)) {
            partnerIndex.byName.set(nameKey, record.label);
        }
    });
}

function renderPartnerOptions(listElement, records, inputValue) {
    const normalizedInput = normalizePartnerValue(inputValue || '');
    const matches = getPartnerMatches(records, normalizedInput)
        .slice(0, PARTNER_SUGGESTION_LIMIT);

    renderDatalistOptions(listElement, matches.map((record) => record.label));
}

function setSupplierRecords(records, elements) {
    supplierRecords = sortRecordsByLabel(records, 'name').map((record) => ({
        ...record,
        normalizedCode: normalizeSupplierValue(record.code),
        normalizedName: normalizeSupplierValue(record.name),
        label: record.name,
        labelNormalized: normalizeSupplierValue(record.name)
    }));

    buildSupplierIndex(supplierRecords);
    renderSupplierOptions(elements.supplierList, supplierRecords, elements.supplierInput.value);
}

function buildSupplierIndex(records) {
    supplierIndex = {
        byLabel: new Map(),
        byCode: new Map(),
        byName: new Map()
    };

    records.forEach((record) => {
        const labelKey = normalizeSupplierValue(record.label);
        const codeKey = normalizeSupplierValue(record.code);
        const nameKey = normalizeSupplierValue(record.name);

        supplierIndex.byLabel.set(labelKey, record.label);
        supplierIndex.byCode.set(codeKey, record.label);
        if (!supplierIndex.byName.has(nameKey)) {
            supplierIndex.byName.set(nameKey, record.label);
        }
    });
}

function renderSupplierOptions(listElement, records, inputValue) {
    const normalizedInput = normalizeSupplierValue(inputValue || '');
    const matches = getSupplierMatches(records, normalizedInput)
        .slice(0, SUPPLIER_SUGGESTION_LIMIT);

    renderDatalistOptions(listElement, matches.map((record) => record.label));
}

function setProductRecords(records, elements) {
    const seen = new Set();

    productRecords = sortValuesByLengthThenAlpha(records.map((value) => sanitizeProductCode(value)))
        .filter((value) => {
            const normalized = normalizeProductValue(value);
            if (!normalized || seen.has(normalized)) {
                return false;
            }

            seen.add(normalized);
            return true;
        })
        .map((code) => ({
            code,
            normalizedCode: normalizeProductValue(code),
            label: code
        }));

    buildProductIndex(productRecords);
    renderProductOptions(elements.productCodeList, productRecords, getActiveProductInputValue(elements));
}

function buildProductIndex(records) {
    productIndex = {
        byCode: new Map()
    };

    records.forEach((record) => {
        productIndex.byCode.set(record.normalizedCode, record.code);
    });
}

function renderProductOptions(listElement, records, inputValue) {
    const normalizedInput = normalizeProductValue(inputValue || '');
    const matches = getProductMatches(records, normalizedInput)
        .slice(0, PRODUCT_SUGGESTION_LIMIT);

    renderDatalistOptions(listElement, matches.map((record) => record.label));
}

function renderDatalistOptions(listElement, values) {
    listElement.innerHTML = '';

    const fragment = document.createDocumentFragment();
    values.forEach((value) => {
        const option = document.createElement('option');
        option.value = value;
        fragment.appendChild(option);
    });

    listElement.appendChild(fragment);
}

function getActiveProductInputValue(elements) {
    const activeElement = document.activeElement;
    if (activeElement && activeElement.classList?.contains('product-code-input')) {
        return activeElement.value;
    }

    const firstInput = elements.productStatusList.querySelector('.product-code-input');
    return firstInput ? firstInput.value : '';
}

function getSupplierMatches(records, normalizedInput) {
    if (!normalizedInput) {
        return records;
    }

    return records.filter((record) => {
        if (record.normalizedCode.includes(normalizedInput)) {
            return true;
        }

        if (record.normalizedName.includes(normalizedInput)) {
            return true;
        }

        if (record.labelNormalized.includes(normalizedInput)) {
            return true;
        }

        return false;
    });
}

function resolveSupplierLabel(inputValue) {
    const trimmed = inputValue.trim();
    if (!trimmed) {
        return '';
    }

    const normalized = normalizeSupplierValue(trimmed);

    if (supplierIndex.byLabel.has(normalized)) {
        return supplierIndex.byLabel.get(normalized);
    }

    if (supplierIndex.byCode.has(normalized)) {
        return supplierIndex.byCode.get(normalized);
    }

    if (supplierIndex.byName.has(normalized)) {
        return supplierIndex.byName.get(normalized);
    }

    const matches = getSupplierMatches(supplierRecords, normalized);
    if (matches.length > 0) {
        return matches[0].label;
    }

    return trimmed;
}

function parseCsvRecords(csvText, codeHeaders, normalizeValue) {
    const lines = csvText.split(/\r?\n/).filter((line) => line.trim().length > 0);
    const records = [];
    const normalizedCodeHeaders = codeHeaders.map((header) => normalizeHeaderValue(header, normalizeValue));

    lines.forEach((line, index) => {
        const fields = parseCsvLine(line);
        if (fields.length < 2) {
            return;
        }

        const code = fields[0].trim();
        const name = fields[1].trim();
        if (!code || !name) {
            return;
        }

        const normalizedCode = normalizeHeaderValue(code, normalizeValue);
        const normalizedName = normalizeHeaderValue(name, normalizeValue);
        const isHeader = index === 0 && normalizedCodeHeaders.includes(normalizedCode) && normalizedName.includes('ten');

        if (isHeader) {
            return;
        }

        records.push({
            code,
            name
        });
    });

    return records;
}

function parseCsvValues(csvText, headers, normalizeValue) {
    const lines = csvText.split(/\r?\n/).filter((line) => line.trim().length > 0);
    const normalizedHeaders = headers.map((header) => normalizeHeaderValue(header, normalizeValue));
    const values = [];

    lines.forEach((line, index) => {
        const fields = parseCsvLine(line);
        if (!fields.length) {
            return;
        }

        const value = fields[0].trim();
        if (!value) {
            return;
        }

        const normalizedValue = normalizeHeaderValue(value, normalizeValue);
        const isHeader = index === 0 && normalizedHeaders.includes(normalizedValue);
        if (isHeader) {
            return;
        }

        values.push(value);
    });

    return values;
}

function parseCsvLine(line) {
    const fields = [];
    let current = '';
    let inQuotes = false;

    for (let index = 0; index < line.length; index += 1) {
        const char = line[index];

        if (char === '"') {
            if (inQuotes && line[index + 1] === '"') {
                current += '"';
                index += 1;
            } else {
                inQuotes = !inQuotes;
            }
            continue;
        }

        if (char === ',' && !inQuotes) {
            fields.push(current);
            current = '';
            continue;
        }

        current += char;
    }

    fields.push(current);
    return fields;
}

function removeDiacritics(value) {
    return value
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd')
        .replace(/Đ/g, 'D');
}

function normalizePartnerValue(value) {
    return removeDiacritics(value).toLowerCase().replace(/\s+/g, ' ').trim();
}

function normalizeSupplierValue(value) {
    return removeDiacritics(value).toLowerCase().replace(/\s+/g, ' ').trim();
}

function normalizeProductValue(value) {
    return removeDiacritics(sanitizeProductCode(value)).toLowerCase().replace(/\s+/g, ' ').trim();
}

function normalizeHeaderValue(value, normalizeValue) {
    return normalizeValue(value)
        .replace(/[^\w\s]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

function splitPartnerTokens(value) {
    return normalizePartnerValue(value)
        .split('+')
        .map((part) => part.trim())
        .filter(Boolean);
}

function getPartnerMatches(records, normalizedInput) {
    const inputTokens = splitPartnerTokens(normalizedInput);
    if (!normalizedInput) {
        return records;
    }

    return records.filter((record) => {
        if (record.normalizedCode.includes(normalizedInput)) {
            return true;
        }

        if (record.normalizedName.includes(normalizedInput)) {
            return true;
        }

        if (record.labelNormalized.includes(normalizedInput)) {
            return true;
        }

        return inputTokens.every((token) => record.tokens.some((recordToken) => recordToken.includes(token)));
    });
}

function resolvePartnerLabel(inputValue) {
    const trimmed = inputValue.trim();
    if (!trimmed) {
        return '';
    }

    const normalized = normalizePartnerValue(trimmed);

    if (partnerIndex.byLabel.has(normalized)) {
        return partnerIndex.byLabel.get(normalized);
    }

    if (partnerIndex.byCode.has(normalized)) {
        return partnerIndex.byCode.get(normalized);
    }

    if (partnerIndex.byName.has(normalized)) {
        return partnerIndex.byName.get(normalized);
    }

    const matches = getPartnerMatches(partnerRecords, normalized);
    if (matches.length > 0) {
        return matches[0].label;
    }

    return trimmed;
}

function getProductMatches(records, normalizedInput) {
    if (!normalizedInput) {
        return records;
    }

    return records.filter((record) => record.normalizedCode.includes(normalizedInput));
}

function resolveProductCode(inputValue) {
    const trimmed = sanitizeProductCode(inputValue);
    if (!trimmed) {
        return '';
    }

    const normalized = normalizeProductValue(trimmed);
    if (productIndex.byCode.has(normalized)) {
        return productIndex.byCode.get(normalized);
    }

    const matches = getProductMatches(productRecords, normalized);
    if (matches.length > 0) {
        return matches[0].code;
    }

    return trimmed;
}

function getProductStatusEntries(elements) {
    return Array.from(elements.productStatusList.querySelectorAll('.product-status-row')).map((row) => {
        const codeInput = row.querySelector('.product-code-input');
        const statusSelect = row.querySelector('.product-status-select');

        return {
            productCodeInput: codeInput ? codeInput.value.trim() : '',
            productCode: codeInput ? resolveProductCode(codeInput.value) : '',
            statusInput: statusSelect ? statusSelect.value.trim() : ''
        };
    });
}

function buildActivationOptionsMarkup() {
    const placeholderOption = '<option value="">Chọn trạng thái</option>';
    const options = ACTIVATION_OPTIONS.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`);
    return [placeholderOption, ...options].join('');
}

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function copyText(elements) {
    const caseValue = elements.caseSelect.value;
    const caseLabel = CASE_COPY_LABELS[caseValue] || '';
    const partnerValue = resolvePartnerLabel(elements.partnerInput.value);
    const supplierValue = resolveSupplierLabel(elements.supplierInput.value);
    const productStatusEntries = getProductStatusEntries(elements);
    const filledProductStatusEntries = productStatusEntries.filter((entry) => entry.productCodeInput || entry.statusInput);

    if (!caseValue) {
        alert('Vui lòng chọn trường hợp.');
        return;
    }

    if (!filledProductStatusEntries.length) {
        alert('Vui lòng chọn ít nhất 1 mã sản phẩm và trạng thái.');
        return;
    }

    const hasIncompleteRow = filledProductStatusEntries.some((entry) => !entry.productCodeInput || !entry.statusInput);
    if (hasIncompleteRow) {
        alert('Vui lòng chọn đủ mã sản phẩm và trạng thái cho từng dòng.');
        return;
    }

    const contextParts = [];

    if (CASES_WITH_PARTNER_FIRST.has(caseValue)) {
        if (partnerValue) {
            contextParts.push(partnerValue);
        }
        if (supplierValue) {
            contextParts.push(supplierValue);
        }
        if (caseLabel) {
            contextParts.push(caseLabel);
        }
    } else {
        if (caseLabel) {
            contextParts.push(caseLabel);
        }
        if (partnerValue) {
            contextParts.push(partnerValue);
        }
        if (supplierValue) {
            contextParts.push(supplierValue);
        }
    }

    const productLines = filledProductStatusEntries.map((entry) => `${entry.productCode} - ${entry.statusInput} | `);
    const footerParts = [];

    const formattedShippingFee = formatShippingFeeForCopy(elements.shippingFee.value.trim());
    if (formattedShippingFee) {
        footerParts.push(`Cước: ${formattedShippingFee}`);
    }

    const noteValue = elements.note.value.trim();
    if (noteValue) {
        footerParts.push(noteValue);
    }

    const lines = [];

    if (contextParts.length) {
        lines.push(contextParts.join(' | '));
    }

    lines.push(...productLines);

    if (footerParts.length) {
        lines.push(footerParts.join(' | '));
    }

    const textToCopy = lines.join('\n');

    navigator.clipboard.writeText(textToCopy).catch((error) => {
        console.error('Không thể sao chép', error);
    });
}

function compareValuesByLengthThenAlpha(firstValue, secondValue) {
    const first = (firstValue || '').trim();
    const second = (secondValue || '').trim();
    const lengthDifference = first.length - second.length;

    if (lengthDifference !== 0) {
        return lengthDifference;
    }

    return first.localeCompare(second, 'vi', { sensitivity: 'base', numeric: true });
}

function sortRecordsByLabel(records, key) {
    return [...records].sort((firstRecord, secondRecord) => {
        const primaryComparison = compareValuesByLengthThenAlpha(firstRecord[key], secondRecord[key]);
        if (primaryComparison !== 0) {
            return primaryComparison;
        }

        return compareValuesByLengthThenAlpha(firstRecord.code, secondRecord.code);
    });
}

function sortValuesByLengthThenAlpha(values) {
    return [...values].sort((firstValue, secondValue) => compareValuesByLengthThenAlpha(firstValue, secondValue));
}

function sanitizeProductCode(value) {
    let sanitized = String(value || '').replace(/\s+/g, ' ').trim();
    if (!sanitized) {
        return '';
    }

    const suffixPatterns = [
        /\s+\[IMEI\]$/i,
        /\s+\[(?:CŨ|CU)\]$/iu,
        /\s+MN$/i,
        /\s+ML$/i
    ];

    let changed = true;
    while (changed) {
        changed = false;
        suffixPatterns.forEach((pattern) => {
            if (pattern.test(sanitized)) {
                sanitized = sanitized.replace(pattern, '').replace(/\s+/g, ' ').trim();
                changed = true;
            }
        });
    }

    return sanitized;
}

function formatTimestamp(timestamp) {
    if (!timestamp) {
        return '';
    }

    const date = new Date(timestamp);
    return date.toLocaleString('vi-VN');
}

function setupShippingFeeInput(shippingFeeInput) {
    if (!shippingFeeInput) {
        return;
    }

    shippingFeeInput.addEventListener('focus', () => {
        shippingFeeInput.value = shippingFeeInput.value.replace(/[^\d]/g, '');
    });

    shippingFeeInput.addEventListener('input', () => {
        shippingFeeInput.value = shippingFeeInput.value.replace(/[^\d]/g, '');
    });

    shippingFeeInput.addEventListener('blur', () => {
        shippingFeeInput.value = formatShippingFeeForInput(shippingFeeInput.value);
    });

    shippingFeeInput.value = formatShippingFeeForInput(shippingFeeInput.value);
}

function formatShippingFeeForInput(rawValue) {
    const digitsOnly = rawValue.replace(/[^\d]/g, '');
    if (!digitsOnly) {
        return '';
    }

    const numberValue = parseInt(digitsOnly, 10);
    if (Number.isNaN(numberValue)) {
        return '';
    }

    return numberValue.toLocaleString('vi-VN');
}

function formatShippingFeeForCopy(rawValue) {
    const formatted = formatShippingFeeForInput(rawValue);
    return formatted ? `${formatted}đ` : '';
}
