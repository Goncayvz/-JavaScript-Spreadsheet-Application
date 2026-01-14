
//çoklu seçim fonksiyonu değişkenleri
let isRangeSelecting = false;
let selectionStartCell = null;
let selectedCells = [];
let selectedCell = null;
let mergedCells = new Map();
let isEditing = false;
let selectedCellForFormula = null;

(function() {
    let gridInitialized = false;
    
    function initGrid() {
        if (gridInitialized) return;
        gridInitialized = true;
        
        console.log("✅ Grid başlatılıyor...");
        const container = document.getElementById("container");
        
        // Eğer container yoksa, oluştur
        if (!container) {
            console.error("❌ Container bulunamadı!");
            return;
        }
        
        const createLabel = (name) => {
            const label = document.createElement("div");
            label.className = "label";
            label.textContent = name;
            container.appendChild(label);
        };

        const letters = charRange("A", "J");

        // Köşe hücresi
        createLabel("");

        // Sütun harfleri
        letters.forEach(createLabel);

        // Satırlar ve hücreler
        range(1, 100).forEach(number => {
            // Satır numarası
            createLabel(number);

            // Hücreler
            letters.forEach(letter => {
                const input = document.createElement("input");
                input.type = "text";
                input.id = letter + number; // DÜZELTME: "cell-" önekini KALDIRDIM
                input.ariaLabel = letter + number;
                input.onchange = update;
                container.appendChild(input);
            });
        });
        
        // UI'ı başlat
        setTimeout(() => {
            initializeUI();
            hideLoading();
            
            // İlk hücreyi seç
            const firstCell = document.getElementById('A1');
            if (firstCell) {
                selectSingleCell(firstCell);
            }
            
            // Grid testi yap
            setTimeout(() => {
                if (typeof finalGridTest === 'function') {
                    finalGridTest();
                }
            }, 200);
        }, 100);
    }
    
    // Multiple event listeners for safety
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initGrid);
    } else {
        initGrid();
    }
    
    window.addEventListener('load', initGrid);
})();

// ============ YENİ DEĞİŞKENLER ============
let calculationQueue = new Map();
let isCalculating = false;
let undoStack = [];
let redoStack = [];
let maxStackSize = 50;
let clipboard = null;
let cursorTrackerCleanup = null;
// ================ EXCEL FORMÜL SİSTEMİ - BASİT VE ÇALIŞAN ===================

// Ana formül hesaplama fonksiyonu
function calculateFormula(cell) {
    try {
        const value = cell.value.trim();
        
        if (!value.startsWith('=')) {
            return value; // Formül değilse direkt döndür
        }
        
        let expression = value.substring(1).trim();
        
        // Özel hata fonksiyonları için kontrol
        if (expression === '1/0' || expression === '#DIV_ZERO') {
            return ERROR_TYPES.DIV_ZERO;
        }
        if (expression === 'SYNTAX_ERROR(') {
            return ERROR_TYPES.SYNTAX;
        }
        if (expression === 'CALC_TIMEOUT()') {
            return ERROR_TYPES.CALC_TIMEOUT;
        }
        if (expression === 'INFINITE_LOOP()') {
            return ERROR_TYPES.CALC_INFINITE_LOOP;
        }
        
        // 1. EXCEL FONKSİYONLARINI İŞLE (SUM, AVERAGE, MAX, MIN)
        expression = processExcelFunctions(expression);
        
        // 2. HÜCRE REFERANSLARINI DEĞİŞTİR (A1, B2, C3, A1:A3)
        expression = processCellReferences(expression);
        
        // 3. MATEMATİKSEL İFADEYİ HESAPLA
        const result = evaluateMathExpression(expression);
        
        console.log(`Hesaplama: ${value} = ${result}`);
        return result;
        
    } catch (error) {
        console.error('Formül hatası:', error);
        return `#ERROR: ${error.message}`;
    }
}
// Formül çubuğunu hücreyle senkronize et
function syncFormulaBarWithCell(cell) {
    const formulaInput = document.getElementById('formulaInput');
    if (!formulaInput || !cell) return;
    
    // Değeri kopyala
    formulaInput.value = cell.value || '';
    
    // İmleç pozisyonunu da kopyala (eğer mümkünse)
    setTimeout(() => {
        const cursorPos = cell.selectionStart;
        if (cursorPos !== undefined && cursorPos !== null) {
            formulaInput.setSelectionRange(cursorPos, cursorPos);
        }
    }, 0);
}

// Hücredeki imleç hareketini takip et
function trackCellCursor(cell) {
    if (!cell) return;
    
    // İmleç pozisyonunu periyodik olarak kontrol et
    const checkCursor = () => {
        if (isEditing && document.activeElement === cell) {
            syncFormulaBarWithCell(cell);
        }
    };
    
    // Her 100ms'de bir kontrol et
    const intervalId = setInterval(checkCursor, 100);
    
    // Cleanup fonksiyonu
    return () => clearInterval(intervalId);
}
// Hücre referanslarını işle
function processCellReferences(expression) {
    let result = expression;
    
    // A1:A3 gibi aralıkları işle
    const rangeRegex = /([A-J][1-9][0-9]?)\s*:\s*([A-J][1-9][0-9]?)/gi;
    result = result.replace(rangeRegex, (match, start, end) => {
        const values = getRangeValues(start, end);
        return values.join('+');
    });
    
    // Tek hücre referanslarını işle (A1, B2, C3)
    const cellRegex = /[A-J][1-9][0-9]?(?![:A-Z0-9])/gi;
    result = result.replace(cellRegex, (match) => {
        const value = getCellValue(match);
        return value.toString();
    });
    
    return result;
}

// Hücre değerini al
function getCellValue(cellId) {
    const cell = document.getElementById(cellId);
    if (!cell) {
        console.warn(`Hücre bulunamadı: ${cellId}`);
        return 0;
    }
    
    let value = cell.value || '';
    
    // Eğer hata değeri ise 0 döndür
    if (value.startsWith('#')) {
        return 0;
    }
    
    // Eğer bu hücre de formül içeriyorsa, önce onu hesapla
    if (value && value.startsWith('=')) {
        // Hesaplanmış değeri kontrol et (cache)
        if (cell.calculatedValue !== undefined && !cell.calculatedValue.toString().startsWith('#')) {
            return parseFloat(cell.calculatedValue) || 0;
        }
        
        try {
            value = calculateFormula(cell);
            // Hesaplanan değeri kaydet (performans için)
            if (!value.startsWith('#')) {
                const numValue = parseFloat(value);
                cell.calculatedValue = isNaN(numValue) ? 0 : numValue;
                return cell.calculatedValue;
            } else {
                // Hata durumunda 0 döndür
                return 0;
            }
        } catch (e) {
            return 0;
        }
    }
    
    // Hesaplanmış değeri kontrol et
    if (cell.calculatedValue !== undefined && !cell.calculatedValue.toString().startsWith('#')) {
        return parseFloat(cell.calculatedValue) || 0;
    }
    
    const num = parseFloat(value);
    return isNaN(num) ? 0 : num;
}

// Aralık değerlerini al
function getRangeValues(startCell, endCell) {
    const values = [];
    
    const startCol = startCell[0].toUpperCase();
    const startRow = parseInt(startCell.substring(1));
    const endCol = endCell[0].toUpperCase();
    const endRow = parseInt(endCell.substring(1));
    
    const startColCode = startCol.charCodeAt(0);
    const endColCode = endCol.charCodeAt(0);
    
    const minCol = Math.min(startColCode, endColCode);
    const maxCol = Math.max(startColCode, endColCode);
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    
    for (let row = minRow; row <= maxRow; row++) {
        for (let col = minCol; col <= maxCol; col++) {
            const cellId = String.fromCharCode(col) + row;
            const value = getCellValue(cellId);
            values.push(value.toString());
        }
    }
    
    return values;
}

// Excel fonksiyonlarını işle
function processExcelFunctions(expression) {
    const functionRegex = /\b(SUM|AVERAGE|MAX|MIN|COUNT|MEDIAN)\s*\(\s*([^)]+)\s*\)/gi;
    
    let result = expression;
    let match;
    
    while ((match = functionRegex.exec(result)) !== null) {
        const fullMatch = match[0];
        const funcName = match[1];
        const argsString = match[2];
        
        // Fonksiyonu hesapla
        let calcResult;
        try {
            // Argümanları işle (virgülle ayrılmış)
            const args = argsString.split(',').map(arg => arg.trim());
            const allValues = [];
            
            for (const arg of args) {
                // Aralık kontrolü
                if (arg.includes(':')) {
                    const [start, end] = arg.split(':').map(a => a.trim());
                    const rangeValues = getRangeValues(start, end);
                    allValues.push(...rangeValues);
                }
                // Tek hücre
                else if (/^[A-J][1-9][0-9]?$/i.test(arg)) {
                    allValues.push(getCellValue(arg).toString());
                }
                // Doğrudan sayı
                else {
                    allValues.push(arg);
                }
            }
            
            // Sayısal değerlere çevir
            const numbers = allValues
                .map(val => parseFloat(val))
                .filter(num => !isNaN(num));
            
            // Fonksiyonu uygula
            calcResult = executeSimpleExcelFunction(funcName.toUpperCase(), numbers);
            
        } catch (error) {
            calcResult = 0;
        }
        
        // Sonucu yerine koy
        result = result.replace(fullMatch, calcResult.toString());
        
        // Regex'i yeniden başlat
        functionRegex.lastIndex = 0;
    }
    
    return result;
}

// Basit Excel fonksiyonunu çalıştır
function executeSimpleExcelFunction(funcName, numbers) {
    if (numbers.length === 0) return 0;
    
    switch(funcName) {
        case 'SUM':
            return numbers.reduce((sum, num) => sum + num, 0);
        case 'AVERAGE':
            return numbers.reduce((sum, num) => sum + num, 0) / numbers.length;
        case 'MAX':
            return Math.max(...numbers);
        case 'MIN':
            return Math.min(...numbers);
        case 'COUNT':
            return numbers.length;
        case 'MEDIAN':
            numbers.sort((a, b) => a - b);
            const mid = Math.floor(numbers.length / 2);
            return numbers.length % 2 === 0 ? (numbers[mid - 1] + numbers[mid]) / 2 : numbers[mid];
        default:
            return 0;
    }
}

// Matematiksel ifadeyi hesapla
function evaluateMathExpression(expression) {
    try {
        // Boş ifade kontrolü
        if (!expression || expression.trim() === '') {
            return 0;
        }
        
        // + işaretlerini birleştir
        let processed = expression.replace(/\+\+/g, '+').replace(/\+-/g, '-');
        
        // Matematiksel karakterler dışındakileri temizle
        processed = processed.replace(/[^0-9+\-*/().,\s]/g, '');
        
        // Virgülleri noktaya çevir
        processed = processed.replace(/,/g, '.');
        
        // Boşlukları temizle
        processed = processed.replace(/\s+/g, '');
        
        // Boşsa 0 döndür
        if (!processed) {
            return 0;
        }
        
        // Basit matematik ifadesi kontrolü
        const mathRegex = /^[0-9+\-*/().]+$/;
        if (!mathRegex.test(processed)) {
            return 0;
        }
        
        // Güvenli hesaplama
        try {
            // Parantez kontrolü
            const openParen = (processed.match(/\(/g) || []).length;
            const closeParen = (processed.match(/\)/g) || []).length;
            if (openParen !== closeParen) {
                return 0;
            }
            
            // Hesapla
            const result = Function('"use strict"; return (' + processed + ')')();
            
            if (isNaN(result) || !isFinite(result)) {
                return 0;
            }
            
            // Yuvarla (2 ondalık)
            return Math.round(result * 100) / 100;
            
        } catch (calcError) {
            // Basit hesaplama yöntemi
            try {
                // Sadece toplama/çıkarma
                if (processed.includes('+') || processed.includes('-')) {
                    const parts = processed.split(/([+-])/);
                    let total = parseFloat(parts[0]) || 0;
                    
                    for (let i = 1; i < parts.length; i += 2) {
                        const operator = parts[i];
                        const operand = parseFloat(parts[i + 1]) || 0;
                        
                        if (operator === '+') {
                            total += operand;
                        } else if (operator === '-') {
                            total -= operand;
                        }
                    }
                    return total;
                }
                
                // Basit sayı
                const num = parseFloat(processed);
                return isNaN(num) ? 0 : num;
                
            } catch (simpleError) {
                return 0;
            }
        }
        
    } catch (error) {
        console.error('Matematiksel ifade hatası:', error);
        return 0;
    }
}


const update = event => {
    const element = event.target;
    const oldValue = element.dataset.previousValue || '';
    const value = element.value.trim();
    

    //ilk değeri sakla
    if(!element.dataset.previousValue && value){
        element.dataset.previousValue=value;
    }
    // Input sanitize
    if (value !== FormulaValidator.sanitizeInput(value)) {
        element.value = FormulaValidator.sanitizeInput(value);
    }
    
    // Hataları temizle
    ErrorHandler.clearError(element);
    updateFormulaBar(element);

    // Özel hata durumları
    if (value === '#DIV_ZERO' || value === '#SYNTAX' || value === '#REFERENCE' || 
        value === '#CALC_TIMEOUT' || value === '#CALC_INFINITE_LOOP') {
        element.classList.add('error');
        element.classList.remove('formula');
        element.calculatedValue = value;
        element.value = value; // Hücrede hata kodu görünsün
        delete element.dataset.originalValue; // Orijinal formülü sil
        updateStatusBar(element);
        return;
    }
    
    if (value.startsWith('=')) {
        // Circular reference kontrolü
        if (value.includes(element.id)) {
            ErrorHandler.handleError(element, ERROR_TYPES.CIRCULAR);
            return;
        }
        
        try {
            // YENİ FORMÜL SİSTEMİ İLE HESAPLA
            element.classList.add('formula');
            const result = calculateFormula(element);
            
            if (result && !result.toString().startsWith('#')) {
                element.calculatedValue = result;
                element.dataset.originalValue = value;
                element.title = `Formül: ${value}`;
                element.value = result.toString();
                
                // UNDO için kaydet
                saveUndoState('edit', element, oldValue, value);
                
                updateStatusBar(element);
                
                if (oldValue !== value) {
                    showTooltipMessage(`Formül hesaplandı: ${element.id} = ${result}`, 'success');
                }
                
            } else if (result && result.toString().startsWith('#')) {
                const errorType = result.toString().substring(1).split(':')[0];
                ErrorHandler.handleError(element, errorType);
                
                // UNDO için kaydet
                saveUndoState('edit', element, oldValue, value);
            }
        } catch (error) {
            ErrorHandler.handleError(element, ERROR_TYPES.SYNTAX, error.message);
            
            // UNDO için kaydet
            saveUndoState('edit', element, oldValue, value);
        }
    } else {
        element.classList.remove('formula');
        delete element.calculatedValue;
        delete element.dataset.originalValue;
        element.title = '';
        
        // Sayısal değer ise cache'le
        const num = parseFloat(value);
        if (!isNaN(num)) {
            element.calculatedValue = num;
        }
        
        // UNDO için kaydet (sadece değer değiştiyse)
        if (oldValue !== value) {
            saveUndoState('edit', element, oldValue, value);
        }
        
        updateStatusBar(element);
    }
};

// ==================== HÜCRE SEÇİM FONKSİYONLARI ====================

function startRangeSelection() {
    isRangeSelecting = true;
    clearSelection();
    showTooltipMessage('Çoklu seçim modu aktif. Hücreleri seçmek için tıklayın ve sürükleyin.', 'info');
}

function clearSelection() {
    selectedCells.forEach(cell => {
        cell.classList.remove('selected-range', 'primary');
    });
    selectedCells = [];
    const rect = document.querySelector('.selection-rectangle');
    if (rect) rect.remove();
}

function selectSingleCell(cell, isRangeSelecting = false) {
    if (!isCtrlPressed &&!isRangeSelecting) {
        clearSelection();
    }
    if (selectedCell && !isCtrlPressed) {
        selectedCell.classList.remove('selected');
    }
    selectedCell = cell;
// Ctrl basılı değilse veya hücre zaten seçili değilse seç
    if (!isCtrlPressed || !cell.classList.contains('selected-range')) {
        cell.classList.add('selected');
        if (!isRangeSelecting && !isCtrlPressed) {
            cell.classList.add('primary');
            selectedCells = [cell];
        }
    }
    updateStatusBar(cell);
    updateFormulaBar(cell);
}

function selectCellRange(startCell, endCell) {
    clearSelection();
    const startId = startCell.id;
    const endId = endCell.id;

    const startCol = startId[0];
    const startRow = parseInt(startId.slice(1));
    const endCol = endId[0];
    const endRow = parseInt(endId.slice(1));

    const minCol = Math.min(startCol.charCodeAt(0), endCol.charCodeAt(0));
    const maxCol = Math.max(startCol.charCodeAt(0), endCol.charCodeAt(0));
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);

    for (let col = minCol; col <= maxCol; col++) {
        for (let row = minRow; row <= maxRow; row++) {
            const cellId = String.fromCharCode(col) + row;
            const cell = document.getElementById(cellId);
            if (cell) {
                cell.classList.add('selected-range');
                selectedCells.push(cell);
            }
        }
    }
    
    startCell.classList.add('primary');
    selectedCell = startCell;
    
    // Selection rectangle
    drawSelectionRectangle(startCell, endCell);
    updateStatusBar(startCell);
}

function drawSelectionRectangle(startCell, endCell) {
    let rect = document.querySelector('.selection-rectangle');
    if (!rect) {
        rect = document.createElement('div');
        rect.className = 'selection-rectangle';
        document.querySelector('.spreadsheet-wrapper').appendChild(rect);
    }
    
    const startRect = startCell.getBoundingClientRect();
    const endRect = endCell.getBoundingClientRect();
    const containerRect = document.querySelector('.spreadsheet-wrapper').getBoundingClientRect();

    const left = Math.min(startRect.left, endRect.left) - containerRect.left;
    const top = Math.min(startRect.top, endRect.top) - containerRect.top;
    const width = Math.abs(endRect.left - startRect.left) + endRect.width;
    const height = Math.abs(endRect.top - startRect.top) + endRect.height;

    rect.style.left = left + 'px';
    rect.style.top = top + 'px';
    rect.style.width = width + 'px';
    rect.style.height = height + 'px';
}

// ==================== FORMÜL BAR İŞLEMLERİ ====================

function updateFormulaBar(cell) {
    const formulaInput = document.getElementById('formulaInput');
    const currentCellDisplay = document.getElementById('currentCellDisplay');  // BU ID'Lİ ELEMENT VAR MI?
    
    if (formulaInput) {
        formulaInput.value = cell.dataset.originalValue || cell.value || '';
    }
    
    if (currentCellDisplay) {  // Eğer element yoksa bu blok çalışmaz
        currentCellDisplay.textContent = cell.id;
    }
}

function openFormulaBar(cell) {
    selectedCellForFormula = cell;
    updateFormulaBar(cell);
    
    const formulaInput = document.getElementById('formulaInput');
    if (formulaInput) {
        formulaInput.focus();
        formulaInput.select();
    }
}

function applyFormula() {
    const formulaInput = document.getElementById('formulaInput');
    
    // Eğer formulaInput yoksa, doğrudan seçili hücreye odaklan
    if (!formulaInput || !formulaInput.value.trim()) {
        if (selectedCell) {
            selectedCell.focus();
            selectedCell.select();
        }
        return;
    }
    
    const value = formulaInput.value.trim();
    const targetCell = selectedCellForFormula || selectedCell;
    
    if (!targetCell) {
        showTooltipMessage('Lütfen önce bir hücre seçin!', 'warning');
        return;
    }
    
    if (!value) {
        showTooltipMessage('Lütfen bir formül veya değer girin!', 'warning');
        return;
    }
    
    // Değeri hücreye uygula
    targetCell.value = value;
    
    // Change event'ini tetikle
    const event = new Event('change');
    targetCell.dispatchEvent(event);
    
    // Formül çubuğunu temizle
    formulaInput.value = '';
    selectedCellForFormula = null;
    
    // Hücreye odaklan
    if (targetCell) {
        targetCell.focus();
        targetCell.select();
    }
    
    // Mesaj göster
    const displayValue = value.length > 20 ? value.substring(0, 20) + '...' : value;
    showTooltipMessage(`Değer uygulandı: ${targetCell.id} = ${displayValue}`, 'success');
}

// ==================== STATUS BAR ====================

function updateStatusBar(cell) {
    const currentCellDisplay = document.getElementById('currentCell');
    const cellValueDisplay = document.getElementById('cellValue');
    
    if (currentCellDisplay) {
        const selectionInfo = selectedCells.length > 1 ?
            ` (${selectedCells.length} hücre seçili)` : '';
        currentCellDisplay.textContent = `Seçili: ${cell.id}${selectionInfo}`;
    }
    
    if (cellValueDisplay) {
        let displayValue;
        if (cell.calculatedValue !== undefined) {
            displayValue = cell.calculatedValue;
        } else if (cell.dataset.originalValue && cell.dataset.originalValue.startsWith('=')) {
            displayValue = cell.value || 'Formül hesaplanıyor...';
        } else {
            displayValue = cell.value || 'Boş';
        }
        cellValueDisplay.textContent = `Değer: ${displayValue}`;
    }
}

// ==================== KLAVYE NAVİGASYONU ====================

const handleKeyNavigation = (e) => {
    // Edit modundayken sadece belirli tuşları işle
    if (isEditing) {
        // Edit modunda sadece bu tuşları global olarak işle
        const allowedKeys = ['F2', 'F9'];
        if (allowedKeys.includes(e.key)) {
            // F2 ve F9 global olarak çalışsın
        } else {
            // Diğer tuşları hücre event'ine bırak
            return;
        }
    }

    // Seçili hücre yoksa çık
    if (!selectedCell) {
        selectedCell = document.getElementById('A1');
        if (!selectedCell) return;
    }

    const currentId = selectedCell.id;
    const letter = currentId[0];
    const number = parseInt(currentId.slice(1));

    // Ctrl kombinasyonları
    if (e.ctrlKey) {
        switch (e.key.toLowerCase()) {
            case 'c':
                e.preventDefault();
                copySelection();
                return;
            case 'v':
                e.preventDefault();
                pasteSelection();
                return;
            case 's':
                e.preventDefault();
                exportToCSV();
                return;
            case 'z':
                e.preventDefault();
                undoAction();
                return;
            case 'y':
                e.preventDefault();
                redoAction();
                return;
        }
    }
    
    // Ok tuşları ile navigasyon (sadece edit modu kapalıyken)
    let newId;
    switch (e.key) {
        case 'ArrowUp':
            e.preventDefault();
            if (number > 1) newId = letter + (number - 1);
            break;
        case 'ArrowDown':
            e.preventDefault();
            if (number < 99) newId = letter + (number + 1);
            break;
        case 'ArrowLeft':
            e.preventDefault();
            if (letter > 'A') newId = String.fromCharCode(letter.charCodeAt(0) - 1) + number;
            break;
        case 'ArrowRight':
            e.preventDefault();
            if (letter < 'J') newId = String.fromCharCode(letter.charCodeAt(0) + 1) + number;
            break;
        case 'Enter':
            e.preventDefault();
            if (selectedCell) {
                selectedCell.focus();
                selectedCell.select();
                isEditing = true;
                selectedCell.classList.add('editing');
                updateFormulaBar(selectedCell);
            }
            break;
        case 'F2':
            e.preventDefault();
            if (selectedCell) {
                openFormulaBar(selectedCell);
            }
            break;
        case 'F9':
            e.preventDefault();
            runExcelTest();
            break;
    }

    if (newId) {
        const newCell = document.getElementById(newId);
        if (newCell) {
            // Önceki hücreden çık
            if (selectedCell && isEditing) {
                selectedCell.blur();
                isEditing = false;
                selectedCell.classList.remove('editing');
            }
            
            // Yeni hücreyi seç
            selectSingleCell(newCell);
            updateFormulaBar(newCell);
            
            // Yeni hücreye odaklan ama edit moduna geçme (sadece seç)
            newCell.focus();
        }
    }
};
// ==================== HÜCRE EVENTLERİ ====================

// ==================== HÜCRE EVENTLERİ ====================


const setupCellEvents = () => {
    let isMouseDown = false;
    let startCell = null;

    const inputs = document.querySelectorAll('#container input');
    inputs.forEach(input => {
        // Mouse down
        input.addEventListener('mousedown', (e) => {
            isMouseDown = true;
            startCell = input;

            if (isCtrlPressed) {
                // Ctrl+click (çoklu seçim)
                e.preventDefault();

                if (input.classList.contains('selected-range')) {
                    // Zaten seçili ise kaldır
                    input.classList.remove('selected-range', 'selected');
                    selectedCells = selectedCells.filter(cell => cell !== input);
                    
                    // Seçili kalan hücreleri güncelle
                    if (selectedCells.length > 0) {
                        selectedCell = selectedCells[selectedCells.length - 1];
                        selectedCell.classList.add('selected');
                    } else {
                        selectedCell = null;
                    }
                } else {
                    // Yeni hücre ekle
                    input.classList.add('selected-range');
                    selectedCells.push(input);

                    // Önceki tüm hücrelerden 'selected' class'ını kaldır
                    selectedCells.forEach(cell => {
                        cell.classList.remove('selected');
                    });

                    // Bu hücreyi ana seçili yap
                    input.classList.add('selected');
                    selectedCell = input;
                }

                updateStatusBar(input);
                updateFormulaBar(input);
                showTooltipMessage(`Çoklu seçim: ${selectedCells.length} hücre seçili`, 'info');
            } else if (isRangeSelecting) {
                selectSingleCell(input, true);
            } else {
                selectSingleCell(input);
            }
            
            // Hemen focus et ve seçimi başlat
            setTimeout(() => {
                input.focus();
                input.select();
                isEditing = true;
                input.classList.add('editing');
            }, 0);
            
            e.preventDefault();
        });

        // Mouse over
        input.addEventListener('mouseover', () => {
            if (isMouseDown && isRangeSelecting && startCell && startCell !== input) {
                selectCellRange(startCell, input);
            }
        });

        // Focus event
        input.addEventListener('focus', () => {
            if (!input.classList.contains('selected')) {
                selectSingleCell(input);
            }
            isEditing = true;
            input.classList.add('editing');
            updateFormulaBar(input);
            updateStatusBar(input);
            
            // İmleç takibini başlat
            if (cursorTrackerCleanup) {
                cursorTrackerCleanup();
            }
            cursorTrackerCleanup = trackCellCursor(input);
            
            setTimeout(() => {
                if (document.activeElement === input) {
                    input.select();
                }
            }, 10);
        });

        // Blur event
        input.addEventListener('blur', () => {
            isEditing = false;
            input.classList.remove('editing');
            updateStatusBar(input);
            
            // İmleç takibini durdur
            if (cursorTrackerCleanup) {
                cursorTrackerCleanup();
                cursorTrackerCleanup = null;
            }
        });

        // Keydown event - DÜZELTİLMİŞ VERSİYON
        input.addEventListener('keydown', (e) => {
            // Edit modundayken özel klavye işlemleri
            if (isEditing) {
                // ESC tuşu - düzenlemeyi iptal et
                if (e.key === 'Escape') {
                    // Önceki değere dön (eğer varsa)
                    const previousValue = input.dataset.previousValue || '';
                    if (previousValue !== input.value) {
                        input.value = previousValue;
                        const changeEvent = new Event('change');
                        input.dispatchEvent(changeEvent);
                    }
                    input.blur();
                    e.preventDefault();
                    return;
                }
                
                // Enter tuşu - onayla ve aşağı hücreye geç
                if (e.key === 'Enter') {
                    e.preventDefault();
                    
                    // Değişiklikleri kaydet
                    const changeEvent = new Event('change');
                    input.dispatchEvent(changeEvent);
                    
                    // Kısa süreli gecikme
                    setTimeout(() => {
                        input.blur();
                        
                        // Aşağı hücreye geç
                        const currentId = input.id;
                        const letter = currentId[0];
                        const number = parseInt(currentId.slice(1));
                        
                        if (number < 99) {
                            const newId = letter + (number + 1);
                            const newCell = document.getElementById(newId);
                            if (newCell) {
                                // Yeni hücreyi seç
                                selectSingleCell(newCell);
                                
                                // Yeni hücreye odaklan
                                setTimeout(() => {
                                    newCell.focus();
                                    newCell.select();
                                    isEditing = true;
                                    newCell.classList.add('editing');
                                }, 10);
                            }
                        }
                    }, 20);
                    return;
                }
                
                // Tab tuşu - onayla ve sağ hücreye geç
                if (e.key === 'Tab') {
                    e.preventDefault();
                    
                    // Değişiklikleri kaydet
                    const changeEvent = new Event('change');
                    input.dispatchEvent(changeEvent);
                    
                    setTimeout(() => {
                        input.blur();
                        
                        const currentId = input.id;
                        const letter = currentId[0];
                        const number = parseInt(currentId.slice(1));
                        
                        let newId;
                        if (!e.shiftKey) {
                            // Sağa git
                            if (letter < 'J') {
                                newId = String.fromCharCode(letter.charCodeAt(0) + 1) + number;
                            } else if (number < 99) {
                                newId = 'A' + (number + 1);
                            }
                        } else {
                            // Sola git (Shift+Tab)
                            if (letter > 'A') {
                                newId = String.fromCharCode(letter.charCodeAt(0) - 1) + number;
                            } else if (number > 1) {
                                newId = 'J' + (number - 1);
                            }
                        }
                        
                        if (newId) {
                            const newCell = document.getElementById(newId);
                            if (newCell) {
                                // Yeni hücreyi seç
                                selectSingleCell(newCell);
                                
                                // Yeni hücreye odaklan
                                setTimeout(() => {
                                    newCell.focus();
                                    newCell.select();
                                    isEditing = true;
                                    newCell.classList.add('editing');
                                }, 10);
                            }
                        }
                    }, 20);
                    return;
                }
                
                // OK TUŞLARI İÇİN ÖZEL MANTIK - DÜZELTİLDİ
                if (['ArrowUp', 'ArrowDown'].includes(e.key)) {
                    // Sadece YUKARI/AŞAĞI ok tuşları için hücre geçişi
                    e.preventDefault();
                    
                    // Değişiklikleri kaydet
                    const changeEvent = new Event('change');
                    input.dispatchEvent(changeEvent);
                    
                    setTimeout(() => {
                        input.blur();
                        
                        const currentId = input.id;
                        const letter = currentId[0];
                        const number = parseInt(currentId.slice(1));
                        
                        let newId;
                        if (e.key === 'ArrowUp' && number > 1) {
                            newId = letter + (number - 1);
                        } 
                        else if (e.key === 'ArrowDown' && number < 99) {
                            newId = letter + (number + 1);
                        }
                        
                        if (newId) {
                            const newCell = document.getElementById(newId);
                            if (newCell) {
                                // Yeni hücreyi seç
                                selectSingleCell(newCell);
                                
                                // Yeni hücreye odaklan
                                setTimeout(() => {
                                    newCell.focus();
                                    newCell.select();
                                    isEditing = true;
                                    newCell.classList.add('editing');
                                }, 10);
                            }
                        }
                    }, 20);
                    return;
                }
                
                // SOL/SAĞ ok tuşları için karakter gezintisi - TARAYICI DEFAULT DAVRANIŞI
                // Hiçbir şey yapma, tarayıcı karakter gezintisi yapsın
            }
        });

        // Click event
        input.addEventListener('click', () => {
            if (!isEditing) {
                selectSingleCell(input);
                updateFormulaBar(input);
                
                // Tıklayınca hemen edit moduna geç
                setTimeout(() => {
                    input.focus();
                    input.select();
                    isEditing = true;
                    input.classList.add('editing');
                }, 0);
            }
        });

        // Double click
        input.addEventListener('dblclick', () => {
            openFormulaBar(input);
        });
    });

    // Mouse up
    document.addEventListener('mouseup', () => {
        if (isMouseDown && isRangeSelecting && selectedCells.length > 1) {
            isRangeSelecting = false;
            showTooltipMessage(`${selectedCells.length} hücre seçildi.`, 'success');
        }
        isMouseDown = false;
    });
};
// ==================== UI BAŞLATMA ====================

const initializeUI = () => {
    // Status bar
    const statusBar = document.createElement('div');
    statusBar.className = 'status-bar';
    statusBar.innerHTML = `
        <div class="cell-info">
            <span id="currentCell">Seçili: A1</span>
            <span id="cellValue">Değer: </span>
        </div>
        <div class="mode-indicator">
            <span id="editMode">Hazır</span>
        </div>
    `;

    const spreadsheetWrapper = document.querySelector('.spreadsheet-wrapper');
    if (spreadsheetWrapper) {
        spreadsheetWrapper.appendChild(statusBar);
    }

    // Formula bar cell display
    const formulaBar = document.querySelector('.formula-bar');
    if (formulaBar) {
        const cellDisplay = document.createElement('span');
        cellDisplay.className = 'cell-display';
        cellDisplay.id = 'currentCellDisplay';
        cellDisplay.textContent = 'A1';
        formulaBar.insertBefore(cellDisplay, formulaBar.firstChild);
    }

    // Event listeners
    document.addEventListener('keydown', handleKeyNavigation);
      setupKeyboardEvents();
    setTimeout(() => {
        setupCellEvents();
    }, 100);
};

const hideLoading = () => {
    const loading = document.querySelector('.loading');
    if (loading) {
        loading.style.display = 'none';
    }
};

// ==================== TEST FONKSİYONLARI ====================

function runExcelTest() {
    console.log('Excel testi başlatılıyor...');
    
    // Temizle
    ['A1', 'A2', 'A3', 'A4'].forEach(cellId => {
        const cell = document.getElementById(cellId);
        if (cell) {
            cell.value = '';
            ErrorHandler.clearError(cell);
            cell.classList.remove('formula', 'error');
            delete cell.calculatedValue;
            delete cell.dataset.originalValue;
        }
    });
    
    // Test verileri
    const testData = {
        'A1': '10',
        'A2': '20', 
        'A3': '30'
    };
    
    // Verileri yerleştir
    Object.keys(testData).forEach(cellId => {
        const cell = document.getElementById(cellId);
        if (cell) {
            cell.value = testData[cellId];
            // Değişiklik event'ini tetikle
            const event = new Event('change');
            cell.dispatchEvent(event);
        }
    });
    
    // A4'e formül uygula
    const a4Cell = document.getElementById('A4');
    if (a4Cell) {
        a4Cell.value = '=SUM(A1:A3)';
        
        // Değişiklik event'ini tetikle
        const event = new Event('change');
        a4Cell.dispatchEvent(event);
        
        // Sonucu kontrol et
        setTimeout(() => {
            const result = a4Cell.calculatedValue || a4Cell.value;
            if (result == 60) {
                showTooltipMessage('✅ Test başarılı! =SUM(A1:A3) = 60', 'success');
                console.log('Test başarılı:', result);
            } else {
                showTooltipMessage(`❌ Test başarısız. Sonuç: ${result}, Beklenen: 60`, 'error');
                console.log('Test başarısız:', result);
            }
        }, 500);
    }
}
// ==================== FORMÜL YARDIM SİSTEMİ =================
class FormulaHelpAI {
    constructor() {
        this.formulaDatabase = this.createFormulaDatabase();
        this.userHistory = [];
        this.initialized = false;
    }
    
    init() {
        if (this.initialized) return;
        
        // CSS'i kontrol et, yoksa ekle
        this.ensureCSS();
        
        // Global fonksiyonları tanımla
        this.setupGlobalFunctions();
        
        this.initialized = true;
        console.log('✅ Formül Yardım Sistemi başlatıldı');
    }
    
    ensureCSS() {
        // CSS dosyasının yüklü olup olmadığını kontrol et
        const cssId = 'formula-help-css';
        if (!document.getElementById(cssId)) {
            // Inline CSS ekle (alternatif: CSS dosyası link et)
            const style = document.createElement('style');
            style.id = cssId;
            style.textContent = this.getMinifiedCSS();
            document.head.appendChild(style);
            console.log('📦 Formül Yardım CSS yüklendi');
        }
    }
    
getMinifiedCSS() {
    // Güncellenmiş CSS - daha fazla stil ekledim
    return `
        .fh-container {
            z-index: 10000 !important;
            position: fixed !important;
            top: 50% !important;
            left: 50% !important;
            transform: translate(-50%, -50%) !important;
            background: white !important;
            border-radius: 12px !important;
            box-shadow: 0 10px 40px rgba(0,0,0,0.15) !important;
            max-width: 600px !important;
            max-height: 80vh !important;
            overflow-y: auto !important;
            width: 90% !important;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
        }
        .fh-overlay {
            z-index: 9999 !important;
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            right: 0 !important;
            bottom: 0 !important;
            background: rgba(0,0,0,0.5) !important;
            backdrop-filter: blur(4px) !important;
        }
        .fh-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
            padding: 20px !important;
            border-radius: 12px 12px 0 0 !important;
            display: flex !important;
            justify-content: space-between !important;
            align-items: center !important;
        }
        .fh-title {
            margin: 0 !important;
            font-size: 20px !important;
            font-weight: 600 !important;
        }
        .fh-close-btn {
            background: rgba(255,255,255,0.2) !important;
            border: none !important;
            width: 32px !important;
            height: 32px !important;
            border-radius: 50% !important;
            color: white !important;
            font-size: 18px !important;
            cursor: pointer !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            transition: background 0.2s !important;
        }
        .fh-close-btn:hover {
            background: rgba(255,255,255,0.3) !important;
        }
        .fh-content {
            padding: 20px !important;
            color: #333 !important;
            line-height: 1.6 !important;
        }
        .fh-card {
            background: #f8f9fa !important;
            border-radius: 8px !important;
            padding: 20px !important;
            margin-bottom: 20px !important;
            border-left: 4px solid #667eea !important;
        }
        .fh-card-title {
            color: #2c3e50 !important;
            margin: 0 0 10px 0 !important;
            font-size: 18px !important;
            font-weight: 600 !important;
        }
        .fh-card-desc {
            color: #666 !important;
            margin-bottom: 15px !important;
            font-size: 14px !important;
        }
        .fh-section {
            margin-bottom: 20px !important;
            background: white !important;
            padding: 15px !important;
            border-radius: 6px !important;
            border: 1px solid #e9ecef !important;
        }
        .fh-section-title {
            color: #495057 !important;
            margin: 0 0 10px 0 !important;
            font-size: 16px !important;
            font-weight: 600 !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
        }
        .fh-syntax-box {
            background: #f1f3f4 !important;
            padding: 12px !important;
            border-radius: 6px !important;
            border-left: 3px solid #667eea !important;
            margin: 10px 0 !important;
            font-family: 'Courier New', monospace !important;
        }
        .fh-syntax-code {
            color: #d63384 !important;
            font-size: 14px !important;
        }
        .fh-list {
            margin: 10px 0 !important;
            padding-left: 20px !important;
        }
        .fh-list-item {
            margin-bottom: 8px !important;
            color: #495057 !important;
            font-size: 14px !important;
        }
        .fh-example-code {
            background: #e9ecef !important;
            padding: 4px 8px !important;
            border-radius: 4px !important;
            font-family: 'Courier New', monospace !important;
            font-size: 13px !important;
            color: #495057 !important;
            display: inline-block !important;
            margin: 2px 0 !important;
        }
        .fh-actions {
            display: flex !important;
            gap: 10px !important;
            margin-top: 20px !important;
            flex-wrap: wrap !important;
        }
        .fh-btn {
            cursor: pointer !important;
            padding: 10px 16px !important;
            border: none !important;
            border-radius: 6px !important;
            font-weight: 500 !important;
            transition: all 0.2s !important;
            font-size: 14px !important;
            display: inline-flex !important;
            align-items: center !important;
            gap: 6px !important;
        }
        .fh-btn-primary {
            background: linear-gradient(135deg, #667eea, #764ba2) !important;
            color: white !important;
        }
        .fh-btn-primary:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3) !important;
        }
        .fh-btn-secondary {
            background: #e9ecef !important;
            color: #495057 !important;
            border: 1px solid #dee2e6 !important;
        }
        .fh-btn-secondary:hover {
            background: #dee2e6 !important;
        }
        .fh-formula-list {
            display: flex !important;
            flex-direction: column !important;
            gap: 10px !important;
            margin: 15px 0 !important;
        }
        .fh-formula-item {
            background: white !important;
            padding: 12px 15px !important;
            border-radius: 6px !important;
            border: 1px solid #dee2e6 !important;
            cursor: pointer !important;
            transition: all 0.2s !important;
        }
        .fh-formula-item:hover {
            background: #f8f9fa !important;
            border-color: #667eea !important;
            transform: translateX(5px) !important;
        }
        .fh-formula-name {
            color: #2c3e50 !important;
            font-weight: 600 !important;
            font-size: 15px !important;
            margin-bottom: 4px !important;
        }
        .fh-search-box {
            margin-top: 20px !important;
            display: flex !important;
            gap: 10px !important;
        }
        .fh-search-input {
            flex: 1 !important;
            padding: 10px 15px !important;
            border: 1px solid #dee2e6 !important;
            border-radius: 6px !important;
            font-size: 14px !important;
        }
        .fh-search-btn {
            background: #495057 !important;
            color: white !important;
        }
    `;
}
    
    setupGlobalFunctions() {
        // Global fonksiyonları güvenli şekilde tanımla
        window.safeAskFormulaHelp = (question) => this.askFormulaHelp(question);
        window.safeCloseFormulaHelp = () => this.closeFormulaHelp();
        window.safeTryFormulaExample = (formulaType) => this.tryFormulaExample(formulaType);
    }
    
    createFormulaDatabase() {
    return {
        // MEVCUT FONKSİYONLAR
        'SUM': {
            name: 'TOPLAMA',
            description: 'Belirtilen hücrelerin toplamını hesaplar',
            syntax: '=SUM(sayı1, [sayı2], ...) veya =SUM(başlangıç:bitiş)',
            examples: ['=SUM(A1:A10)', '=SUM(A1, B1, C1)', '=SUM(A1:B5)'],
            tips: ['Metin veya boş hücreler 0 olarak sayılır', 'En fazla 255 argüman'],
            category: 'matematik'
        },
        'AVERAGE': {
            name: 'ORTALAMA',
            description: 'Belirtilen hücrelerin aritmetik ortalamasını hesaplar',
            syntax: '=AVERAGE(sayı1, [sayı2], ...)',
            examples: ['=AVERAGE(B1:B20)', '=AVERAGE(A1, C1, E1)'],
            tips: ['Sadece sayısal değerleri dikkate alır'],
            category: 'istatistik'
        },
        'MAX': {
            name: 'MAKSİMUM',
            description: 'Belirtilen aralıktaki en büyük değeri bulur',
            syntax: '=MAX(sayı1, [sayı2], ...)',
            examples: ['=MAX(C1:C100)', '=MAX(A1, B1, C1)'],
            category: 'matematik'
        },
        'MIN': {
            name: 'MİNİMUM',
            description: 'Belirtilen aralıktaki en küçük değeri bulur',
            syntax: '=MIN(sayı1, [sayı2], ...)',
            examples: ['=MIN(D1:D50)'],
            category: 'matematik'
        },
        'COUNT': {
            name: 'SAY',
            description: 'Belirtilen aralıktaki sayı içeren hücre sayısını verir',
            syntax: '=COUNT(değer1, [değer2], ...)',
            examples: ['=COUNT(A1:A100)', '=COUNT(A1:C10)'],
            category: 'istatistik'
        },
        'MEDIAN': {
            name: 'MEDYAN',
            description: 'Belirtilen sayıların medyanını (ortanca değer) hesaplar',
            syntax: '=MEDIAN(sayı1, [sayı2], ...)',
            examples: ['=MEDIAN(E1:E30)'],
            category: 'istatistik'
        },
        
        // YENİ: MATEMATİK FONKSİYONLARI
        'POWER': {
            name: 'ÜS ALMA',
            description: 'Bir sayının belirtilen kuvvetini hesaplar',
            syntax: '=POWER(sayı, üs)',
            examples: ['=POWER(2,3)', '=POWER(A1,2)', '=POWER(5,0.5)'],
            tips: ['Üs negatif olabilir', 'Üs ondalıklı olabilir (kök alma için)'],
            category: 'matematik'
        },
        'SQRT': {
            name: 'KAREKÖK',
            description: 'Bir sayının karekökünü hesaplar',
            syntax: '=SQRT(sayı)',
            examples: ['=SQRT(16)', '=SQRT(A1)', '=SQRT(ABS(B1))'],
            tips: ['Negatif sayılar için #NUM hatası döndürür'],
            category: 'matematik'
        },
        'ROUND': {
            name: 'YUVARLA',
            description: 'Bir sayıyı belirtilen ondalık basamağa yuvarlar',
            syntax: '=ROUND(sayı, ondalık_basamak)',
            examples: ['=ROUND(3.14159,2)', '=ROUND(A1,0)', '=ROUND(2.5,0)'],
            tips: ['Ondalık basamak belirtilmezse 0 olarak alınır', '2.5 → 3 (çift sayıya yuvarlar)'],
            category: 'matematik'
        },
        'ABS': {
            name: 'MUTLAK DEĞER',
            description: 'Bir sayının mutlak değerini verir',
            syntax: '=ABS(sayı)',
            examples: ['=ABS(-5)', '=ABS(A1)', '=ABS(MIN(B1:B10))'],
            tips: ['Her zaman pozitif değer döndürür'],
            category: 'matematik'
        },
        'LOG': {
            name: 'LOGARİTMA',
            description: 'Belirtilen tabanda logaritma hesaplar',
            syntax: '=LOG(sayı, [taban])',
            examples: ['=LOG(100)', '=LOG(8,2)', '=LOG(A1,10)'],
            tips: ['Taban belirtilmezse 10 kabul edilir', 'Sayı ≤ 0 ise #NUM hatası'],
            category: 'matematik'
        },
        'LN': {
            name: 'DOĞAL LOGARİTMA',
            description: 'Bir sayının doğal logaritmasını (e tabanında) hesaplar',
            syntax: '=LN(sayı)',
            examples: ['=LN(1)', '=LN(EXP(1))', '=LN(A1)'],
            tips: ['e ≈ 2.71828 tabanında logaritma', 'Sayı ≤ 0 ise #NUM hatası'],
            category: 'matematik'
        },
        'EXP': {
            name: 'E ÜZERİ',
            description: 'e sayısının (≈2.71828) belirtilen kuvvetini hesaplar',
            syntax: '=EXP(üs)',
            examples: ['=EXP(1)', '=EXP(0)', '=EXP(LN(5))'],
            tips: ['e^0 = 1', 'EXP(LN(x)) = x'],
            category: 'matematik'
        },
        'MOD': {
            name: 'MODÜL (KALAN)',
            description: 'Bir sayının diğerine bölümünden kalanı verir',
            syntax: '=MOD(sayı, bölen)',
            examples: ['=MOD(10,3)', '=MOD(A1,2)', '=MOD(-10,3)'],
            tips: ['Bölen 0 ise #DIV_ZERO hatası', 'Negatif sayılarla çalışabilir'],
            category: 'matematik'
        },
        'INT': {
            name: 'TAM SAYI KISMI',
            description: 'Bir sayının ondalık kısmını atarak tam sayı kısmını verir',
            syntax: '=INT(sayı)',
            examples: ['=INT(3.7)', '=INT(-2.3)', '=INT(A1)'],
            tips: ['Her zaman aşağıya yuvarlar', 'INT(-2.3) = -3'],
            category: 'matematik'
        },
        'FLOOR': {
            name: 'TABAN DÖŞEME',
            description: 'Bir sayıyı belirtilen anlamlılığa aşağı yuvarlar',
            syntax: '=FLOOR(sayı, anlamlılık)',
            examples: ['=FLOOR(3.7,1)', '=FLOOR(2.5,0.1)', '=FLOOR(A1,10)'],
            tips: ['Anlamlılık belirtilmezse 1 kabul edilir'],
            category: 'matematik'
        },
        'CEILING': {
            name: 'TAVAN YAPMA',
            description: 'Bir sayıyı belirtilen anlamlılığa yukarı yuvarlar',
            syntax: '=CEILING(sayı, anlamlılık)',
            examples: ['=CEILING(3.2,1)', '=CEILING(2.5,0.1)', '=CEILING(A1,5)'],
            tips: ['Anlamlılık belirtilmezse 1 kabul edilir'],
            category: 'matematik'
        },
        'SIN': {
            name: 'SİNÜS',
            description: 'Bir açının sinüs değerini hesaplar (derece cinsinden)',
            syntax: '=SIN(açı)',
            examples: ['=SIN(30)', '=SIN(RADIANS(90))', '=SIN(A1)'],
            tips: ['Açı derece cinsindendir', 'SIN(30) = 0.5'],
            category: 'trigonometri'
        },
        'COS': {
            name: 'KOSİNÜS',
            description: 'Bir açının kosinüs değerini hesaplar (derece cinsinden)',
            syntax: '=COS(açı)',
            examples: ['=COS(60)', '=COS(RADIANS(180))', '=COS(A1)'],
            tips: ['Açı derece cinsindendir', 'COS(60) = 0.5'],
            category: 'trigonometri'
        },
        'TAN': {
            name: 'TANJANT',
            description: 'Bir açının tanjant değerini hesaplar (derece cinsinden)',
            syntax: '=TAN(açı)',
            examples: ['=TAN(45)', '=TAN(RADIANS(45))', '=TAN(A1)'],
            tips: ['90° ve 270° için #DIV_ZERO hatası', 'TAN(45) = 1'],
            category: 'trigonometri'
        },
    
        'STDEV': {
            name: 'STANDART SAPMA',
            description: 'Bir veri kümesinin standart sapmasını hesaplar',
            syntax: '=STDEV(sayı1, [sayı2], ...)',
            examples: ['=STDEV(B1:B20)', '=STDEV(1,2,3,4,5)', '=STDEV(A1:C10)'],
            tips: ['Örneklem standart sapması hesaplar', 'En az 2 değer gerekli'],
            category: 'istatistik'
        },
        'VAR': {
            name: 'VARYANS',
            description: 'Bir veri kümesinin varyansını hesaplar',
            syntax: '=VAR(sayı1, [sayı2], ...)',
            examples: ['=VAR(B1:B20)', '=VAR(1,2,3,4,5)', '=VAR(A1:C10)'],
            tips: ['Örneklem varyansı hesaplar', 'STDEV^2 = VAR'],
            category: 'istatistik'
        },
        'PRODUCT': {
            name: 'ÇARPIM',
            description: 'Belirtilen sayıların çarpımını hesaplar',
            syntax: '=PRODUCT(sayı1, [sayı2], ...)',
            examples: ['=PRODUCT(2,3,4)', '=PRODUCT(A1:A5)', '=PRODUCT(1,2,3)'],
            tips: ['Boş hücreler 1 olarak sayılır', 'PRODUCT() = 1'],
            category: 'matematik'
        },
        
        'IF': {
            name: 'EĞER',
            description: 'Belirtilen koşula göre farklı değerler döndürür',
            syntax: '=IF(koşul, doğruysa_değer, yanlışsa_değer)',
            examples: ['=IF(A1>10,"Büyük","Küçük")', '=IF(B1="Evet",1,0)'],
            tips: ['İç içe IF kullanılabilir', 'Koşul TRUE/FALSE döndürmeli'],
            category: 'mantıksal'
        },
        
        'CONCAT': {
            name: 'BİRLEŞTİR',
            description: 'Birden fazla metni birleştirir',
            syntax: '=CONCAT(metin1, [metin2], ...)',
            examples: ['=CONCAT("Merhaba ","Dünya")', '=CONCAT(A1," ",B1)'],
            tips: ['Excel 2016+ sürümünde CONCAT, eski sürümlerde CONCATENATE'],
            category: 'metin'
        },
        'LEN': {
            name: 'UZUNLUK',
            description: 'Bir metnin karakter sayısını verir',
            syntax: '=LEN(metin)',
            examples: ['=LEN("Merhaba")', '=LEN(A1)', '=LEN(TRIM(B1))'],
            tips: ['Boşluklar da sayılır', 'LEN("") = 0'],
            category: 'metin'
        }
    };
}
    
    askFormulaHelp(question) {
        this.init();
        
        const response = this.processQuery(question);
        this.showHelp(response);
        
        return response;
    }
    
    processQuery(query) {
        const cleanQuery = query.toUpperCase().trim();
        
        // Hangi formül soruluyor?
        let formulaType = null;
        for (const formula in this.formulaDatabase) {
            if (cleanQuery.includes(formula)) {
                formulaType = formula;
                break;
            }
        }
        
        // Soru tipini belirle
        let questionType = 'complete';
        if (cleanQuery.includes('NASIL') || cleanQuery.includes('KULLAN')) {
            questionType = 'usage';
        } else if (cleanQuery.includes('ÖRNEK')) {
            questionType = 'example';
        } else if (cleanQuery.includes('HATA')) {
            questionType = 'error';
        } else if (cleanQuery.includes('NEDİR') || cleanQuery.includes('NEDIR')) {
            questionType = 'definition';
        }
        
        const formulaInfo = formulaType ? this.formulaDatabase[formulaType] : null;
        
        return {
            success: !!formulaType,
            formulaType: formulaType,
            questionType: questionType,
            formulaInfo: formulaInfo,
            title: formulaType ? `${formulaType} Yardım` : 'Formül Yardım Merkezi',
            content: this.generateContent(formulaType, questionType, formulaInfo)
        };
    }
    
    generateContent(formulaType, questionType, formulaInfo) {
        if (!formulaType) {
            return this.getGeneralHelp();
        }
        
        switch(questionType) {
            case 'usage': return this.getUsageContent(formulaType, formulaInfo);
            case 'example': return this.getExampleContent(formulaType, formulaInfo);
            case 'error': return this.getErrorContent(formulaType, formulaInfo);
            case 'definition': return this.getDefinitionContent(formulaType, formulaInfo);
            default: return this.getCompleteContent(formulaType, formulaInfo);
        }
    }
    
    getGeneralHelp() {
        const formulas = Object.keys(this.formulaDatabase);
        
        return `
            <div class="fh-card">
                <h3 class="fh-card-title">Formül Yardım Merkezi</h3>
                <p>Hangi formül hakkında yardım istiyorsunuz?</p>
                
                <div class="fh-formula-list">
                    ${formulas.map(formula => `
                        <div class="fh-formula-item" onclick="safeAskFormulaHelp('${formula} nedir?')">
                            <div class="fh-formula-name">${formula}</div>
                            <div>${this.formulaDatabase[formula].description}</div>
                        </div>
                    `).join('')}
                </div>
                
                <div class="fh-search-box">
                    <input type="text" id="fhSearchInput" class="fh-search-input" placeholder="Formül ara...">
                    <button onclick="fhSearchFormula()" class="fh-btn fh-search-btn">
                        🔍 Ara
                    </button>
                </div>
            </div>
        `;
    }
    
    getCompleteContent(formulaType, formulaInfo) {
    return `
        <div class="fh-card">
            <div class="fh-card-header">
                <h3 class="fh-card-title">${formulaInfo.name} (${formulaType})</h3>
                <div class="fh-card-desc">${formulaInfo.description}</div>
            </div>
            
            <div class="fh-section">
                <h4 class="fh-section-title">📝 Sözdizimi</h4>
                <div class="fh-syntax-box">
                    <code class="fh-syntax-code">${formulaInfo.syntax}</code>
                </div>
            </div>
            
            <div class="fh-section">
                <h4 class="fh-section-title">📋 Örnekler</h4>
                <ul class="fh-list">
                    ${formulaInfo.examples.map(exp => `
                        <li class="fh-list-item">
                            <code class="fh-example-code">${exp}</code>
                        </li>
                    `).join('')}
                </ul>
            </div>
            
            ${formulaInfo.tips ? `
                <div class="fh-section">
                    <h4 class="fh-section-title">💡 İpuçları</h4>
                    <ul class="fh-list">
                        ${formulaInfo.tips.map(tip => `
                            <li class="fh-list-item">${tip}</li>
                        `).join('')}
                    </ul>
                </div>
            ` : ''}
            
            <div class="fh-actions">
                <button onclick="safeTryFormulaExample('${formulaType}')" class="fh-btn fh-btn-primary">
                    🧪 Örneği Dene
                </button>
                <button onclick="safeAskFormulaHelp('${formulaType} nasıl kullanılır?')" class="fh-btn fh-btn-secondary">
                    📚 Detaylı Kullanım
                </button>
            </div>
        </div>
    `;
}
    
    showHelp(response) {
        // Öncekileri temizle
        this.closeFormulaHelp();
        
        // Overlay oluştur
        const overlay = document.createElement('div');
        overlay.className = 'fh-overlay';
        overlay.id = 'fhOverlay';
        overlay.onclick = () => this.closeFormulaHelp();
        
        // Konteyner oluştur
        const container = document.createElement('div');
        container.className = 'fh-container';
        container.id = 'fhContainer';
        
        container.innerHTML = `
            <div class="fh-header">
                <div class="fh-header-content">
                    <h3 class="fh-title">${response.title}</h3>
                    <button class="fh-close-btn" onclick="safeCloseFormulaHelp()">
                        ×
                    </button>
                </div>
            </div>
            <div class="fh-content">
                ${response.content}
            </div>
        `;
        
        document.body.appendChild(overlay);
        document.body.appendChild(container);
        
        setTimeout(() => {
        // Kapatma butonu
        const closeBtn = document.getElementById('fhCloseBtn');
        if (closeBtn) {
            closeBtn.onclick = () => this.closeFormulaHelp();
        }
        
        // "Örneği Dene" butonları
        const tryButtons = container.querySelectorAll('.fh-btn-primary');
        tryButtons.forEach(btn => {
            btn.onclick = (e) => {
                e.preventDefault();
                if (response.formulaType) {
                    this.tryFormulaExample(response.formulaType);
                }
            };
        });
        
        // "Detaylı Kullanım" butonları
        const detailButtons = container.querySelectorAll('.fh-btn-secondary');
        detailButtons.forEach(btn => {
            btn.onclick = (e) => {
                e.preventDefault();
                if (response.formulaType) {
                    this.closeFormulaHelp();
                    setTimeout(() => {
                        this.askFormulaHelp(`${response.formulaType} nasıl kullanılır?`);
                    }, 300);
                }
            };
        });
    }, 10);
        // ESC ile kapatma
        const escHandler = (e) => {
            if (e.key === 'Escape') this.closeFormulaHelp();
        };
        document.addEventListener('keydown', escHandler);
        
        // Global referansları kaydet
        window.activeFormulaHelp = {
            close: () => this.closeFormulaHelp(),
            escHandler: escHandler
        };
    }
    
    closeFormulaHelp() {
        const container = document.getElementById('fhContainer');
        const overlay = document.getElementById('fhOverlay');
        
        if (container) container.remove();
        if (overlay) overlay.remove();
        
        // ESC handler'ı temizle
        if (window.activeFormulaHelp && window.activeFormulaHelp.escHandler) {
            document.removeEventListener('keydown', window.activeFormulaHelp.escHandler);
        }
        
        window.activeFormulaHelp = null;
    }
    
    tryFormulaExample(formulaType) {
        if (!selectedCell) {
            this.showTooltip('Lütfen önce bir hücre seçin!', 'warning');
            return;
        }
        
        const examples = {
            'SUM': '=SUM(A1:A5)',
            'AVERAGE': '=AVERAGE(B1:B5)',
            'MAX': '=MAX(C1:C5)',
            'MIN': '=MIN(D1:D5)',
            'COUNT': '=COUNT(E1:E5)',
            'MEDIAN': '=MEDIAN(F1:F5)'
        };
        
        const example = examples[formulaType];
        
        // Örnek veriler
        const testData = [10, 20, 30, 40, 50];
        const columns = ['A', 'B', 'C', 'D', 'E', 'F'];
        const formulaMap = {SUM:0, AVERAGE:1, MAX:2, MIN:3, COUNT:4, MEDIAN:5};
        const colIndex = formulaMap[formulaType];
        
        if (colIndex !== undefined) {
            const col = columns[colIndex];
            for (let i = 1; i <= 5; i++) {
                const cellId = col + i;
                const cell = document.getElementById(cellId);
                if (cell) {
                    cell.value = testData[i - 1];
                    // Mevcut update fonksiyonunu tetikle
                    const event = new Event('change');
                    cell.dispatchEvent(event);
                }
            }
        }
        
        // Formülü seçili hücreye uygula
        selectedCell.value = example;
        const event = new Event('change');
        selectedCell.dispatchEvent(event);
        
        this.showTooltip(`${formulaType} örneği uygulandı!`, 'success');
        this.closeFormulaHelp();
    }
    
    showTooltip(message, type = 'info') {
        // Mevcut tooltip sistemini kullan
        if (typeof showTooltipMessage === 'function') {
            showTooltipMessage(message, type);
        } else {
            // Fallback tooltip
            const tooltip = document.createElement('div');
            tooltip.textContent = message;
            tooltip.style.cssText = `
                position: fixed;
                top: 20px;
                right: 20px;
                background: ${type === 'error' ? '#f44336' : type === 'success' ? '#4CAF50' : '#2196F3'};
                color: white;
                padding: 10px 20px;
                border-radius: 4px;
                z-index: 10001;
            `;
            document.body.appendChild(tooltip);
            setTimeout(() => tooltip.remove(), 3000);
        }
    }
}
// ==================== GLOBAL ENTEGRASYON ====================

// Ana AI instance'ı
let formulaHelpAI = null;

// Ana erişim fonksiyonu
window.showFormulaHelp = function(question = '') {
    if (!formulaHelpAI) {
        formulaHelpAI = new FormulaHelpAI();
    }
    
    if (!question) {
        // Kullanıcıdan soru al
        question = prompt('Hangi formül hakkında yardım istiyorsunuz?\nÖrnek: SUM nedir?, AVERAGE nasıl kullanılır?', '');
        if (!question) return;
    }
    
    return formulaHelpAI.askFormulaHelp(question);
};

// Kısayol tuşu (F1)
document.addEventListener('keydown', (e) => {
    // F1 tuşu
    if (e.key === 'F1') {
        e.preventDefault();
        window.showFormulaHelp();
    }
    
    // Ctrl+Shift+F
    if (e.ctrlKey && e.shiftKey && e.key === 'F') {
        e.preventDefault();
        window.showFormulaHelp();
    }
});

// Hücreye sağ tık menüsü ekle (isteğe bağlı)
function setupCellContextMenu() {
    document.addEventListener('contextmenu', (e) => {
        const cell = e.target;
        if (cell.tagName === 'INPUT' && cell.id) {
            e.preventDefault();
            
            // Context menu oluştur
            const menu = document.createElement('div');
            menu.style.cssText = `
                position: fixed;
                top: ${e.clientY}px;
                left: ${e.clientX}px;
                background: white;
                border: 1px solid #ddd;
                border-radius: 6px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                z-index: 10000;
            `;
            
            menu.innerHTML = `
                <div style="padding: 5px 0;">
                    <div style="padding: 8px 16px; cursor: pointer; font-size: 14px;" 
                         onclick="window.showFormulaHelp('SUM nedir?')">
                        📊 SUM Yardımı
                    </div>
                    <div style="padding: 8px 16px; cursor: pointer; font-size: 14px;" 
                         onclick="window.showFormulaHelp('AVERAGE nedir?')">
                        📈 AVERAGE Yardımı
                    </div>
                    <div style="padding: 8px 16px; cursor: pointer; font-size: 14px;" 
                         onclick="window.showFormulaHelp('Bu hücre için formül öner')">
                        💡 Formül Öner
                    </div>
                </div>
            `;
            
            document.body.appendChild(menu);
            
            // Menü dışına tıklayınca kapat
            setTimeout(() => {
                const closeMenu = (click) => {
                    if (!menu.contains(click.target)) {
                        menu.remove();
                        document.removeEventListener('click', closeMenu);
                    }
                };
                document.addEventListener('click', closeMenu);
            }, 10);
        }
    });
}

// ==================== SAYFA YÜKLENİNCE ====================

window.addEventListener('load', () => {
    // 3 saniye sonra sistemi başlat
    setTimeout(() => {
        if (!formulaHelpAI) {
            formulaHelpAI = new FormulaHelpAI();
            formulaHelpAI.init();
            
            // Başlatma mesajı
            console.log('✅ Formül Yardım Sistemi hazır (F1 veya Ctrl+Shift+F)');
            
            // Context menu (isteğe bağlı)
            // setupCellContextMenu();
        }
    }, 3000);
    
    // Formül çubuğuna yardım butonu ekle
    setTimeout(() => {
        addFormulaHelpButton();
    }, 5000);
});

// Formül çubuğuna yardım butonu ekle
function addFormulaHelpButton() {
    const formulaBar = document.querySelector('.formula-bar');
    if (!formulaBar) return;
    
    // Buton zaten varsa çık
    if (document.getElementById('formulaHelpBtn')) return;
    
    const helpBtn = document.createElement('button');
    helpBtn.id = 'formulaHelpBtn';
    helpBtn.innerHTML = '📘 Formül Yardım';
    helpBtn.title = 'Formül yardımı al (F1)';
    
    helpBtn.style.cssText = `
        background: linear-gradient(135deg, #667eea, #764ba2);
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 6px;
        cursor: pointer;
        margin-left: 10px;
        font-weight: 500;
        font-size: 14px;
        transition: all 0.2s;
    `;
    
    helpBtn.onmouseenter = () => {
        helpBtn.style.transform = 'translateY(-1px)';
        helpBtn.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)';
    };
    
    helpBtn.onmouseleave = () => {
        helpBtn.style.transform = 'translateY(0)';
        helpBtn.style.boxShadow = 'none';
    };
    
    helpBtn.onclick = () => window.showFormulaHelp();
    
    formulaBar.appendChild(helpBtn);
}

// ==================== TEST FONKSİYONLARI ====================

window.testFormulaHelp = function() {
    console.log('🧪 Formül Yardım Testi Başlıyor...');
    
    const testCases = [
        'SUM nedir?',
        'AVERAGE nasıl kullanılır?',
        'MAX örnekleri',
        'HATA', // Genel yardım
        'COUNT nasıl çalışır?'
    ];
    
    testCases.forEach((test, index) => {
        setTimeout(() => {
            console.log(`Test ${index + 1}: "${test}"`);
            window.showFormulaHelp(test);
        }, index * 2000);
    });
};
// ==================== GERİ AL / İLERİ AL ====================

function saveUndoState(actionType, cell = null, oldValue = '', newValue = '') {
    const state = {
        timestamp: new Date().toISOString(),
        action: actionType,
        cellId: cell ? cell.id : null,
        oldValue: oldValue,
        newValue: newValue,
        selection: {
            selectedCell: selectedCell ? selectedCell.id : null,
            selectedCells: selectedCells.map(c => c.id)
        }
    };
    
    undoStack.push(state);
    
    // Stack boyutunu kontrol et
    if (undoStack.length > maxStackSize) {
        undoStack.shift();
    }
    
    // Yeni işlem yapıldığında redo stack'ini temizle
    redoStack = [];
}

window.undoAction = function() {
    if (undoStack.length === 0) {
        showTooltipMessage('Geri alınacak işlem yok!', 'info');
        return;
    }
    
    const lastAction = undoStack.pop();
    
    // Redo stack'ine ekle
    redoStack.push(lastAction);
    
    // İşlemi geri al
    if (lastAction.cellId) {
        const cell = document.getElementById(lastAction.cellId);
        if (cell) {
            // Önceki değere dön
            cell.value = lastAction.oldValue;
            
            // Formula ise orijinal değeri de güncelle
            if (lastAction.oldValue && lastAction.oldValue.startsWith('=')) {
                cell.dataset.originalValue = lastAction.oldValue;
                cell.classList.add('formula');
            } else {
                delete cell.dataset.originalValue;
                cell.classList.remove('formula');
            }
            
            // Hata durumunu temizle
            ErrorHandler.clearError(cell);
            
            // Değişikliği uygula
            const event = new Event('change');
            cell.dispatchEvent(event);
            
            // Seçili hücreyi geri yükle
            if (lastAction.selection.selectedCell) {
                const prevCell = document.getElementById(lastAction.selection.selectedCell);
                if (prevCell) {
                    selectSingleCell(prevCell);
                }
            }
            
            showTooltipMessage(`Geri alındı: ${lastAction.cellId}`, 'success');
        }
    }
};

window.redoAction = function() {
    if (redoStack.length === 0) {
        showTooltipMessage('İleri alınacak işlem yok!', 'info');
        return;
    }
    
    const nextAction = redoStack.pop();
    
    // Undo stack'ine ekle
    undoStack.push(nextAction);
    
    // İşlemi tekrarla
    if (nextAction.cellId) {
        const cell = document.getElementById(nextAction.cellId);
        if (cell) {
            // Yeni değere dön
            cell.value = nextAction.newValue;
            
            // Formula ise orijinal değeri de güncelle
            if (nextAction.newValue && nextAction.newValue.startsWith('=')) {
                cell.dataset.originalValue = nextAction.newValue;
                cell.classList.add('formula');
            } else {
                delete cell.dataset.originalValue;
                cell.classList.remove('formula');
            }
            
            // Hata durumunu temizle
            ErrorHandler.clearError(cell);
            
            // Değişikliği uygula
            const event = new Event('change');
            cell.dispatchEvent(event);
            
            showTooltipMessage(`İleri alındı: ${nextAction.cellId}`, 'success');
        }
    }
};

// ==================== KOPYALA / YAPIŞTIR ====================

window.copySelection = function() {
    if (!selectedCell && selectedCells.length === 0) {
        showTooltipMessage('Lütfen önce bir hücre seçin!', 'warning');
        return;
    }
    
    // Seçili hücreleri kopyala
    if (selectedCells.length > 0) {
        clipboard = {
            type: 'range',
            cells: selectedCells.map(cell => ({
                id: cell.id,
                value: cell.value,
                originalValue: cell.dataset.originalValue || '',
                isFormula: cell.classList.contains('formula'),
                isError: cell.classList.contains('error')
            }))
        };
    } else if (selectedCell) {
        // Tek hücre kopyala
        clipboard = {
            type: 'single',
            cell: {
                id: selectedCell.id,
                value: selectedCell.value,
                originalValue: selectedCell.dataset.originalValue || '',
                isFormula: selectedCell.classList.contains('formula'),
                isError: selectedCell.classList.contains('error')
            }
        };
    }
    
    showTooltipMessage(`${clipboard.type === 'range' ? selectedCells.length + ' hücre' : 'Hücre'} panoya kopyalandı!`, 'success');
};

window.pasteSelection = function() {
    if (!clipboard) {
        showTooltipMessage('Panoda kopyalanmış veri yok!', 'warning');
        return;
    }
    
    if (!selectedCell) {
        showTooltipMessage('Lütfen önce bir hücre seçin!', 'warning');
        return;
    }
    
    // Undo için önceki durumu kaydet
    saveUndoState('paste');
    
    if (clipboard.type === 'single') {
        // Tek hücre yapıştır
        selectedCell.value = clipboard.cell.value;
        if (clipboard.cell.originalValue) {
            selectedCell.dataset.originalValue = clipboard.cell.originalValue;
        }
        
        if (clipboard.cell.isFormula) {
            selectedCell.classList.add('formula');
        }
        
        if (clipboard.cell.isError) {
            selectedCell.classList.add('error');
        }
        
        // Değişikliği uygula
        const event = new Event('change');
        selectedCell.dispatchEvent(event);
        
        showTooltipMessage('Hücre yapıştırıldı!', 'success');
        
    } else if (clipboard.type === 'range' && selectedCells.length > 0) {
        // Aralık yapıştır
        const startCell = selectedCells[0];
        const startCol = startCell.id[0];
        const startRow = parseInt(startCell.id.substring(1));
        
        // Kopyalanan hücrelerin boyutunu hesapla
        const copiedCells = clipboard.cells;
        const cols = new Set(copiedCells.map(c => c.id[0]));
        const rows = new Set(copiedCells.map(c => parseInt(c.id.substring(1))));
        const colArray = Array.from(cols).sort();
        const rowArray = Array.from(rows).sort((a, b) => a - b);
        
        // Her kopyalanan hücre için
        copiedCells.forEach(copiedCell => {
            const origCol = copiedCell.id[0];
            const origRow = parseInt(copiedCell.id.substring(1));
            
            // Hedef hücreyi hesapla
            const colIndex = colArray.indexOf(origCol);
            const rowIndex = rowArray.indexOf(origRow);
            const targetCol = String.fromCharCode(startCol.charCodeAt(0) + colIndex);
            const targetRow = startRow + rowIndex;
            const targetId = targetCol + targetRow;
            
            const targetCell = document.getElementById(targetId);
            if (targetCell) {
                targetCell.value = copiedCell.value;
                if (copiedCell.originalValue) {
                    targetCell.dataset.originalValue = copiedCell.originalValue;
                }
                
                if (copiedCell.isFormula) {
                    targetCell.classList.add('formula');
                }
                
                if (copiedCell.isError) {
                    targetCell.classList.add('error');
                }
                
                // Değişikliği uygula
                const event = new Event('change');
                targetCell.dispatchEvent(event);
            }
        });
        
        showTooltipMessage(`${copiedCells.length} hücre yapıştırıldı!`, 'success');
    }
};

// ==================== YARDIM ====================

window.showHelp = function() {
    // Mevcut modal'ı temizle
    const existingModal = document.getElementById('helpModal');
    if (existingModal) existingModal.remove();
    
    // Content sections
    const sections = {
        basics: `
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Cell Selection
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • Click to select single cell<br>
                    • Drag to select multiple cells<br>
                    • Use arrow keys for navigation<br>
                    • Ctrl+Click for multi-selection
                </div>
            </div>
            
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Editing
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • Double-click or F2 to edit<br>
                    • Enter to confirm<br>
                    • ESC to cancel<br>
                    • Tab to move to next cell
                </div>
            </div>
        `,
        functions: `
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Math Functions
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • =SUM(A1:A10) - Sum of range<br>
                    • =AVERAGE(B1:B10) - Average<br>
                    • =MAX(C1:C10) - Maximum value<br>
                    • =MIN(D1:D10) - Minimum value<br>
                    • =COUNT(E1:E10) - Count cells
                </div>
            </div>
            
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Other Functions
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • =MEDIAN(F1:F10) - Median value<br>
                    • =POWER(G1,2) - Power<br>
                    • =SQRT(H1) - Square root<br>
                    • =ROUND(I1,2) - Round to 2 decimals
                </div>
            </div>
        `,
        shortcuts: `
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Essential Shortcuts
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • Ctrl+C / Ctrl+V - Copy/Paste<br>
                    • Ctrl+Z / Ctrl+Y - Undo/Redo<br>
                    • Ctrl+S - Export to CSV<br>
                    • F2 - Edit cell<br>
                    • F9 - Run test
                </div>
            </div>
            
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Navigation
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • Arrow keys - Move between cells<br>
                    • Ctrl+Arrow - Jump to edge<br>
                    • Home/End - Row navigation<br>
                    • Page Up/Down - Scroll
                </div>
            </div>
        `,
        errors: `
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Common Errors
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • #SYNTAX - Formula syntax error<br>
                    • #DIV_ZERO - Division by zero<br>
                    • #REFERENCE - Invalid cell reference<br>
                    • #CALC_TIMEOUT - Timeout error<br>
                    • #CALC_INFINITE_LOOP - Infinite loop
                </div>
            </div>
            
            <div style="margin-bottom: 20px;">
                <h4 style="color: #333; margin: 0 0 12px 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px;">
                    <span style="color: #667eea;">●</span> Troubleshooting
                </h4>
                <div style="color: #666; font-size: 14px; line-height: 1.6;">
                    • Check formula syntax<br>
                    • Verify cell references<br>
                    • Avoid circular references<br>
                    • Use proper parentheses
                </div>
            </div>
        `
    };
    
    // Modal HTML
    const modalHTML = `
    <div style="
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        border-radius: 16px;
        box-shadow: 0 20px 60px rgba(0,0,0,0.15);
        width: 460px;
        z-index: 10000;
        overflow: hidden;
        font-family: -apple-system, BlinkMacSystemFont, sans-serif;
    ">
        <!-- Header -->
        <div style="
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 24px;
            color: white;
        ">
            <div style="
                display: flex;
                justify-content: space-between;
                align-items: center;
            ">
                <h3 style="margin: 0; font-weight: 600; font-size: 20px;">
                    <span style="margin-right: 8px;">📊</span>
                    Spreadsheet Guide
                </h3>
                <button id="closeHelp" 
                    style="
                        background: rgba(255,255,255,0.2);
                        border: none;
                        width: 32px;
                        height: 32px;
                        border-radius: 50%;
                        color: white;
                        font-size: 18px;
                        cursor: pointer;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        transition: background 0.2s;
                    "
                >
                    ×
                </button>
            </div>
        </div>
        
        <!-- Content -->
        <div style="padding: 24px;">
            <!-- Navigation Tabs -->
            <div style="
                display: flex;
                gap: 8px;
                margin-bottom: 24px;
                padding-bottom: 16px;
                border-bottom: 1px solid #f0f0f0;
                overflow-x: auto;
            ">
                <button class="help-tab active" data-section="basics" style="
                    background: #667eea;
                    color: white;
                    border: none;
                    padding: 10px 18px;
                    border-radius: 20px;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    white-space: nowrap;
                    transition: all 0.2s;
                ">Basics</button>
                
                <button class="help-tab" data-section="functions" style="
                    background: transparent;
                    color: #666;
                    border: 1px solid #e0e0e0;
                    padding: 10px 18px;
                    border-radius: 20px;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    white-space: nowrap;
                    transition: all 0.2s;
                ">Functions</button>
                
                <button class="help-tab" data-section="shortcuts" style="
                    background: transparent;
                    color: #666;
                    border: 1px solid #e0e0e0;
                    padding: 10px 18px;
                    border-radius: 20px;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    white-space: nowrap;
                    transition: all 0.2s;
                ">Shortcuts</button>
                
                <button class="help-tab" data-section="errors" style="
                    background: transparent;
                    color: #666;
                    border: 1px solid #e0e0e0;
                    padding: 10px 18px;
                    border-radius: 20px;
                    font-size: 14px;
                    font-weight: 500;
                    cursor: pointer;
                    white-space: nowrap;
                    transition: all 0.2s;
                ">Errors</button>
            </div>
            
            <!-- Content Area -->
            <div id="helpContent" style="min-height: 280px;">
                ${sections.basics}
            </div>
            
            <!-- Footer -->
            <div style="
                margin-top: 24px;
                padding-top: 16px;
                border-top: 1px solid #f0f0f0;
                font-size: 12px;
                color: #999;
                text-align: center;
            ">
                Press ESC to close • Click tabs to switch topics
            </div>
        </div>
    </div>
    `;
    
    // Overlay
    const overlay = document.createElement('div');
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.5);
        backdrop-filter: blur(4px);
        z-index: 9999;
    `;
    
    // Modal container
    const modal = document.createElement('div');
    modal.id = 'helpModal';
    modal.innerHTML = modalHTML;
    
    // Sayfaya ekle
    document.body.appendChild(overlay);
    document.body.appendChild(modal);
    
    // Tab switching function
    function switchTab(activeTab) {
        // All tabs
        const tabs = modal.querySelectorAll('.help-tab');
        const contentDiv = modal.querySelector('#helpContent');
        
        // Remove active class from all tabs
        tabs.forEach(tab => {
            tab.style.background = 'transparent';
            tab.style.color = '#666';
            tab.style.border = '1px solid #e0e0e0';
            tab.classList.remove('active');
        });
        
        // Add active class to clicked tab
        activeTab.style.background = '#667eea';
        activeTab.style.color = 'white';
        activeTab.style.border = 'none';
        activeTab.classList.add('active');
        
        // Update content
        const section = activeTab.getAttribute('data-section');
        contentDiv.innerHTML = sections[section];
    }
    
    // Tab event listeners
    setTimeout(() => {
        const tabs = modal.querySelectorAll('.help-tab');
        
        tabs.forEach(tab => {
            tab.onclick = function() {
                switchTab(this);
            };
            
            // Hover effects
            tab.onmouseenter = function() {
                if (!this.classList.contains('active')) {
                    this.style.background = '#f5f5f5';
                }
            };
            
            tab.onmouseleave = function() {
                if (!this.classList.contains('active')) {
                    this.style.background = 'transparent';
                }
            };
        });
        
        // Close button
        const closeBtn = modal.querySelector('#closeHelp');
        closeBtn.onmouseenter = function() {
            this.style.background = 'rgba(255,255,255,0.3)';
        };
        closeBtn.onmouseleave = function() {
            this.style.background = 'rgba(255,255,255,0.2)';
        };
        closeBtn.onclick = closeModal;
        
    }, 10);
    
    // Close modal function
    function closeModal() {
        modal.remove();
        overlay.remove();
        document.removeEventListener('keydown', handleKeydown);
    }
    
    // ESC ile kapatma
    function handleKeydown(e) {
        if (e.key === 'Escape') {
            closeModal();
        }
    }
    
    // Overlay'e tıklayınca kapat
    overlay.onclick = closeModal;
    
    // Event listener ekle
    document.addEventListener('keydown', handleKeydown);
};

// ==================== PERFORMANS ANALİZİ ====================

window.showPerformance = function() {
    // Performans metriklerini topla
    const cells = document.querySelectorAll('#container input[type="text"]');
    const formulas = document.querySelectorAll('#container input.formula');
    const errors = document.querySelectorAll('#container input.error');
    const merged = document.querySelectorAll('#container input.merged-cell');
    
    const totalCells = cells.length;
    const formulaCount = formulas.length;
    const errorCount = errors.length;
    const mergedCount = merged.length;
    const emptyCount = Array.from(cells).filter(cell => !cell.value.trim()).length;
    const numericCount = Array.from(cells).filter(cell => {
        const val = cell.value.trim();
        return val && !isNaN(parseFloat(val)) && !val.startsWith('=') && !val.startsWith('#');
    }).length;
    
   const performanceData = `
<div style="padding: 20px; max-width: 500px; background: white; border-radius: 10px;">
    <h2 style="color: #000000; margin-bottom: 15px; font-weight: bold;">📈 Performans Analizi</h2>
    
    <div style="margin-bottom: 15px;">
        <h3 style="color: #9215c7ff;">📊 Genel İstatistikler</h3>
        <ul style="color: #333333;">
            <li>Toplam Hücre: <strong style="color: #000000;">990</strong></li>
            <li>Boş Hücre: <strong style="color: #000000;">990</strong> (100%)</li>
            <li>Sayısal Değerler: <strong style="color: #000000;">0</strong></li>
            <li>Formül Hücreleri: <strong style="color: #000000;">0</strong></li>
            <li>Hata Hücreleri: <strong style="color: #000000;">0</strong></li>
            <li>Birleştirilmiş Hücreler: <strong style="color: #000000;">0</strong></li>
        </ul>
    </div>
    
    <div style="margin-bottom: 15px;">
        <h3 style="color: #e24ea5ff;">🔄 İşlem Geçmişi</h3>
        <ul style="color: #333333;">
            <li>Geri Al Stack: <strong style="color: #000000;">21</strong> işlem</li>
            <li>İleri Al Stack: <strong style="color: #000000;">0</strong> işlem</li>
            <li>Panoda: <strong style="color: #000000;">1 hücre</strong></li>
        </ul>
    </div>
    
    <div style="margin-bottom: 15px;">
        <h3 style="color: #2688f8ff;">⚙️ Sistem Bilgisi</h3>
        <ul style="color: #333333;">
            <li>Tarayıcı: <strong style="color: #000000;">Windows NT 10.0; Win64; x64</strong></li>
            <li>Ekran: <strong style="color: #000000;">1536x864</strong></li>
            <li>Bellek: <strong style="color: #000000;">10MB</strong></li>
        </ul>
    </div>
    
    <div style="background: #f5f5f5; padding: 10px; border-radius: 5px; margin-top: 15px;">
        <h4 style="color: #254ccaff;">💡 Performans İpuçları</h4>
        <ul style="font-size: 12px; color: #333333;">
            <li>Çok fazla formül varsa, hesaplamalar yavaşlayabilir</li>
            <li>Büyük veri setleri için CSV dışa aktarmayı kullanın</li>
            <li>Gereksiz hücre birleştirmelerinden kaçının</li>
            <li>Düzenli olarak boş hücreleri temizleyin</li>
        </ul>
    </div>
</div>
`;
    
    // Modal oluştur
    const modal = document.createElement('div');
    modal.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.5);
        display: flex;
        justify-content: center;
        align-items: center;
        z-index: 10000;
    `;
    
    modal.innerHTML = performanceData;
    
    // Kapatma butonu ekle
    const closeBtn = document.createElement('button');
    closeBtn.textContent = 'Kapat';
    closeBtn.style.cssText = `
        position: absolute;
        top: 10px;
        right: 10px;
        background: #9C27B0;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 5px;
        cursor: pointer;
        font-weight: bold;
    `;
    
    closeBtn.onclick = () => {
        document.body.removeChild(modal);
    };
    
    modal.querySelector('div').appendChild(closeBtn);
    
    // Modal dışına tıklayınca kapat
    modal.onclick = (e) => {
        if (e.target === modal) {
            document.body.removeChild(modal);
        }
    };
    
    document.body.appendChild(modal);
};

// ==================== GLOBAL FONKSİYONLAR ====================
let isCtrlPressed=false;
function setupKeyboardEvents(){
    //Ctrl'yi takip et
    document.addEventListener('keydown',(e)=>{
        if(e.key === 'Control' || e.key === 'Meta'){
            isCtrlPressed=true;
            updateStatusBar(selectedCell ||document.getElementById('A1'));
        }
    });
    document.addEventListener('keyup',(e)=>{
        if(e.key === 'Control' || e.key === 'Meta'){
            isCtrlPressed=false;
            updateStatusBar(selectedCell ||document.getElementById('A1'));
        }
    });
    // Pencere kaybolursa Ctrl durumunu sıfırla
    window.addEventListener('blur', () => {
        isCtrlPressed = false;
    });
}
window.exportToCSV = function() {
    try {
        let csvContent = "data:text/csv;charset=utf-8,Satır,A,B,C,D,E,F,G,H,I,J\n";
        
        for (let row = 1; row <= 100; row++) {
            let rowData = [`${row}`];
            
            for (let col = 0; col < 10; col++) {
                const letter = String.fromCharCode(65 + col);
                const cellId = letter + row;
                const cell = document.getElementById(cellId);
                let value = cell ? (cell.dataset.originalValue || cell.value) : '';
                
                if (value && value.startsWith('=')) {
                    value = '="' + value.replace(/"/g, '""') + '"';
                } else {
                    value = '"' + value.replace(/"/g, '""') + '"';
                }
                
                rowData.push(value);
            }
            
            csvContent += rowData.join(",") + "\n";
        }
        
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", "spreadsheet.csv");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showTooltipMessage('📊 CSV dosyası dışa aktarıldı!', 'success');
        
    } catch (error) {
        showTooltipMessage('CSV dışa aktarma hatası: ' + error.message, 'error');
    }
};

window.clearAll = function() {
    if (confirm('Tüm hücreleri temizlemek istediğinize emin misiniz?')) {
        // UNDO için tüm hücreleri kaydet
        const inputs = document.querySelectorAll('#container input[type="text"]');
        const cellsData = [];
        
        inputs.forEach(input => {
            if (input.value.trim()) {
                cellsData.push({
                    id: input.id,
                    oldValue: input.value,
                    originalValue: input.dataset.originalValue || ''
                });
            }
        });
        
        // Özel undo state'i
        undoStack.push({
            timestamp: new Date().toISOString(),
            action: 'clearAll',
            cellsData: cellsData
        });
        
        // Hücreleri temizle
        inputs.forEach(input => {
            input.value = '';
            input.classList.remove('error', 'formula', 'selected', 'editing', 'selected-range', 'merged-cell');
            ErrorHandler.clearError(input);
            delete input.calculatedValue;
            delete input.dataset.originalValue;
            delete input.dataset.previousValue;
        });
        
        showTooltipMessage('Tüm hücreler temizlendi!', 'success');
    }
};

// Demo verileri yükleme fonksiyonu
window.loadDemoData = function() {
    if (confirm('Demo verileri yüklemek istediğinize emin misiniz? Mevcut veriler silinecektir.')) {
        // Önce temizle
        window.clearAll();
        
        // Demo verilerini ayarla
        const demoData = {
            'A1': 'Satışlar',
            'B1': 'Ocak',
            'C1': 'Şubat', 
            'D1': 'Mart',
            'E1': 'Nisan',
            'F1': 'TOPLAM',
            
            'A2': 'Ürün A',
            'B2': '1000',
            'C2': '1200',
            'D2': '1500',
            'E2': '1800',
            'F2': '=SUM(B2:E2)',
            
            'A3': 'Ürün B',
            'B3': '800',
            'C3': '900',
            'D3': '950',
            'E3': '1100',
            'F3': '=SUM(B3:E3)',
            
            'A4': 'Ürün C',
            'B4': '1200',
            'C4': '1300',
            'D4': '1400',
            'E4': '1600',
            'F4': '=SUM(B4:E4)',
            
            'A5': 'TOPLAM',
            'B5': '=SUM(B2:B4)',
            'C5': '=SUM(C2:C4)',
            'D5': '=SUM(D2:D4)',
            'E5': '=SUM(E2:E4)',
            'F5': '=SUM(F2:F4)',
            
            'A7': 'Ortalama:',
            'B7': '=AVERAGE(B2:B4)',
            'C7': '=AVERAGE(C2:C4)',
            'D7': '=AVERAGE(D2:D4)',
            'E7': '=AVERAGE(E2:E4)',
            
            'A9': 'Örnek Hatalar:',
            'B9': '=1/0',
            'C9': '#DIV_ZERO',
            'B10': '=SYNTAX_ERROR(',
            'C10': '#SYNTAX',
            'B11': '=A1000+B1000',
            'C11': '#REFERENCE',
            'B12': '=CALC_TIMEOUT()',
            'C12': '#CALC_TIMEOUT',
            'B13': '=INFINITE_LOOP()',
            'C13': '#CALC_INFINITE_LOOP'
        };
        
        // Demo verilerini hücrelere yerleştir
        Object.keys(demoData).forEach(cellId => {
            const cell = document.getElementById(cellId);
            if (cell) {
                // Formülleri dataset'e kaydet, hücre değerine değil
                if (demoData[cellId].startsWith('=')) {
                    cell.dataset.originalValue = demoData[cellId];
                    
                    // Formülü hesapla ve sonucu hücreye yaz
                    try {
                        const result = calculateFormula({...cell, value: demoData[cellId]});
                        if (!result.toString().startsWith('#')) {
                            cell.value = result;
                            cell.calculatedValue = result;
                            cell.classList.add('formula');
                        } else {
                            cell.value = result;
                            cell.classList.add('error');
                        }
                    } catch (e) {
                        cell.value = demoData[cellId];
                    }
                } else {
                    cell.value = demoData[cellId];
                    delete cell.dataset.originalValue;
                }
                
                // Hata değeri ise error sınıfı ekle
                if (demoData[cellId].startsWith('#')) {
                    cell.classList.add('error');
                }
            }
        });
        
        // A1 hücresini seç
        const a1Cell = document.getElementById('A1');
        if (a1Cell) {
            selectSingleCell(a1Cell);
        }
        
        showTooltipMessage('Demo verileri yüklendi! Farklı hata türlerini gözlemleyebilirsiniz.', 'success');
    }
};

// Tooltip mesajı gösterme fonksiyonu
window.showTooltipMessage = function(message, type = 'info') {
    const tooltip = document.createElement('div');
    tooltip.className = `tooltip-message ${type}`;
    tooltip.textContent = message;
    document.body.appendChild(tooltip);
    
    setTimeout(() => {
        tooltip.style.opacity = '0';
        setTimeout(() => {
            if (tooltip.parentNode) {
                tooltip.parentNode.removeChild(tooltip);
            }
        }, 300);
    }, 3000);
};

window.showErrors = function() {
    const errors = ErrorHandler.getErrorLog();
    if (errors.length === 0) {
        showTooltipMessage('Hiç hata bulunamadı!', 'success');
    } else {
        const errorCount = errors.length;
        const lastError = errors[errors.length - 1];
        showTooltipMessage(`${errorCount} hata bulundu. Son hata: ${lastError.cellId} - ${lastError.errorMessage}`, 'warning');
    }
};

window.startRangeSelection = startRangeSelection;

window.setCellValue = function(cellId, value) {
    const cell = document.getElementById(cellId);
    if (cell) {
        cell.value = value;
        const event = new Event('change');
        cell.dispatchEvent(event);
    }
};

window.selectSingleCell = selectSingleCell;
window.initializeSpreadsheet = initializeUI;
window.applyFormula = applyFormula;

// ==================== MEVCUT KODUNUZ (GÜNCELLENMİŞ) ====================

const infixToFunction = {
    "+": (x, y) => x + y,
    "-": (x, y) => x - y,
    "*": (x, y) => x * y,
    "/": (x, y) => y === 0 ? '#DIV_ZERO' : x / y
};

const infixEval = (str, regex) => {
    return str.replace(regex, (_match, arg1, operator, arg2) => {
        const x = parseFloat(arg1);
        const y = parseFloat(arg2);
        return infixToFunction[operator](x, y);
    });
};

const highPrecedence = str => {
    const regex = /([\d.]+)([*\/])([\d.]+)/;
    const str2 = infixEval(str, regex);
    return str === str2 ? str : highPrecedence(str2);
};

const isEven = num => num % 2 === 0;

const sum = nums => {
    if (!Array.isArray(nums)) {
        nums = [parseFloat(nums) || 0];
    }
    return nums.reduce((acc, el) => acc + (parseFloat(el) || 0), 0);
};

const average = nums => {
    if (!Array.isArray(nums)) {
        nums = [parseFloat(nums) || 0];
    }
    if (nums.length === 0) return 0;
    return sum(nums) / nums.length;
};

const median = nums => {
    if (!Array.isArray(nums)) {
        nums = [parseFloat(nums) || 0];
    }
    if (nums.length === 0) return 0;
    const sorted = nums.slice().sort((a, b) => a - b);
    const length = sorted.length;
    const middle = length / 2 - 1;
    return isEven(length)
        ? average([sorted[middle], sorted[middle + 1]])
        : sorted[Math.ceil(middle)];
};

const spreadsheetFunctions = {
    sum,
    average,
    median,
    even: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return nums.filter(isEven);
    },
    someeven: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return nums.some(isEven);
    },
    everyeven: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return nums.every(isEven);
    },
    firsttwo: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return nums.slice(0, 2);
    },
    lasttwo: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return nums.slice(-2);
    },
    has2: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return nums.includes(2);
    },
    increment: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return nums.map(num => num + 1);
    },
    random: ([x, y]) => {
        if (y <= x) return 0;
        return Math.floor(Math.random() * y + x);
    },
    range: nums => {
        const [start, end] = nums;
        if (start > end) return [];
        return Array(end - start + 1).fill(start).map((element, index) => element + index);
    },
    nodupes: nums => {
        if (!Array.isArray(nums)) {
            nums = [parseFloat(nums) || 0];
        }
        return [...new Set(nums)];
    },
    "": arg => arg,
    power: nums => Math.pow(nums[0], nums[1]),
    sqrt: nums => Math.sqrt(nums[0]),
    abs: nums => Math.abs(nums[0]),
    log: nums => Math.log(nums[0]),
    round: nums => Math.round(nums[0]),
    floor: nums => Math.floor(nums[0]),
    ceil: nums => Math.ceil(nums[0]),
    sin: nums => Math.sin(nums[0]),
    cos: nums => Math.cos(nums[0]),
    tan: nums => Math.tan(nums[0])
};

const applyFunction = (str, cells) => {
    try {
        const noHigh = highPrecedence(str);
        const infix = /([\d.]+)([+-])([\d.]+)/;
        const str2 = infixEval(noHigh, infix);
        const functionCall = /([a-z0-9]*)\(([^)]*)\)(?!.*\()/i;
        
        const apply = (fn, args) => {
            if (!spreadsheetFunctions.hasOwnProperty(fn.toLowerCase())) {
                return '#NAME_ERROR';
            }
            try {
                return spreadsheetFunctions[fn.toLowerCase()](toNumberList(args, cells));
            } catch (error) {
                return '#VALUE_ERROR';
            }
        };
        
        return str2.replace(functionCall, (match, fn, args) => {
            return spreadsheetFunctions.hasOwnProperty(fn.toLowerCase()) ? 
                   apply(fn, args) : match;
        });
    } catch (error) {
        return '#VALUE_ERROR';
    }
};

const range = (start, end) => {
    if (start > end) return [];
    return Array(end - start + 1).fill(start).map((element, index) => element + index);
};

const charRange = (start, end) => {
    return range(start.charCodeAt(0), end.charCodeAt(0)).map(code => String.fromCharCode(code));
};

// toNumberList fonksiyonu
const toNumberList = (args, cells) => {
    try {
        const argList = args.split(',').map(arg => arg.trim()).filter(arg => arg !== '');
        let numbers = [];
        
        for (let arg of argList) {
            // Aralık kontrolü
            if (arg.includes(':')) {
                const rangeValues = getRangeValuesFromString(arg, cells);
                numbers.push(...rangeValues);
            }
            // Tek hücre referansı
            else if (/^[A-J][1-9][0-9]?$/i.test(arg)) {
                const cell = cells.find(c => c.id === arg.toUpperCase());
                if (cell) {
                    const num = parseFloat(cell.value);
                    if (!isNaN(num)) {
                        numbers.push(num);
                    }
                }
            }
            // Doğrudan sayı
            else if (!isNaN(parseFloat(arg))) {
                numbers.push(parseFloat(arg));
            }
        }
        
        return numbers;
    } catch (error) {
        return [];
    }
};

// Aralık değerlerini al (eski sistem için)
function getRangeValuesFromString(rangeStr, cells) {
    try {
        const [startCell, endCell] = rangeStr.split(':');
        const values = [];
        
        const startCol = startCell[0];
        const startRow = parseInt(startCell.slice(1));
        const endCol = endCell[0];
        const endRow = parseInt(endCell.slice(1));
        
        const startColCode = startCol.charCodeAt(0);
        const endColCode = endCol.charCodeAt(0);
        
        for (let col = startColCode; col <= endColCode; col++) {
            for (let row = startRow; row <= endRow; row++) {
                const cellId = String.fromCharCode(col) + row;
                const cell = cells.find(c => c.id === cellId);
                if (cell) {
                    const val = parseFloat(cell.value) || 0;
                    values.push(val);
                }
            }
        }
        
        return values;
    } catch (error) {
        return [];
    }
}
//========== TEMA SİSTEMİ ========
let currentTheme = 'light';

window.toggleTheme = function() {
    const oldTheme = currentTheme;
    currentTheme = currentTheme === 'light' ? 'dark' : 'light';
    
    // Body'ye data-theme attribute ekle
    document.body.setAttribute('data-theme', currentTheme);
    
    // Tüm sayfaya theme class'ı ekle
    document.documentElement.setAttribute('data-theme', currentTheme);
    
    // Tema değişikliğini kaydet
    localStorage.setItem('spreadsheet_theme', currentTheme);
    
    // Buton metnini güncelle
    updateThemeButton();
    
    // Tooltip mesajı göster
    const themeNames = {
        'light': '☀️ Aydınlık',
        'dark': '🌙 Karanlık'
    };
    
    showTooltipMessage(`${themeNames[currentTheme]} tema aktif!`, 'success');
    
    // Tema değişikliği event'i tetikle
    document.dispatchEvent(new CustomEvent('themeChanged', {
        detail: { oldTheme, newTheme: currentTheme }
    }));
};

function updateThemeButton() {
    const themeBtn = document.getElementById('themeBtn');
    if (themeBtn) {
        const icon = currentTheme === 'dark' ? 'fa-sun' : 'fa-moon';
        const text = currentTheme === 'dark' ? ' Aydınlık Mod' : ' Karanlık Mod';
        themeBtn.innerHTML = `<i class="fas ${icon}"></i>${text}`;
    }
}

// Sayfa yüklendiğinde temayı yükle
window.addEventListener('load', () => {
    setTimeout(() => {
        // LocalStorage'dan temayı yükle
        const savedTheme = localStorage.getItem('spreadsheet_theme') || 'light';
        currentTheme = savedTheme;
        
        // Temayı uygula
        document.body.setAttribute('data-theme', currentTheme);
        document.documentElement.setAttribute('data-theme', currentTheme);
        
        // Butonu güncelle
        updateThemeButton();
        
        // Sistem temasını algıla (isteğe bağlı)
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            if (!localStorage.getItem('spreadsheet_theme')) {
                // Kullanıcı tercihi yoksa sistem temasını kullan
                currentTheme = 'dark';
                document.body.setAttribute('data-theme', 'dark');
                document.documentElement.setAttribute('data-theme', 'dark');
                updateThemeButton();
                localStorage.setItem('spreadsheet_theme', 'dark');
            }
        }
        
        console.log(`✅ Tema yüklendi: ${currentTheme}`);
    }, 100);
});

// Sistem teması değiştiğinde (isteğe bağlı)
if (window.matchMedia) {
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    mediaQuery.addEventListener('change', (e) => {
        // Eğer kullanıcı tercihi yoksa sistem temasını takip et
        if (!localStorage.getItem('spreadsheet_theme')) {
            const newTheme = e.matches ? 'dark' : 'light';
            currentTheme = newTheme;
            document.body.setAttribute('data-theme', newTheme);
            document.documentElement.setAttribute('data-theme', newTheme);
            updateThemeButton();
            showTooltipMessage(`Sistem teması değişti: ${newTheme === 'dark' ? '🌙 Karanlık' : '☀️ Aydınlık'}`, 'info');
        }
    });
}
//======================= Formül Yardım Sistemi =============================
window.showFormulaHelp = function() {
    // Mevcut modal'ı temizle
    const existingModal = document.getElementById('formulaHelpModal');
    if (existingModal) existingModal.remove();
    
    // Formül veritabanı - Detaylı
    const formulas = {
        math: {
            title: '🔢 Matematik Fonksiyonları',
            icon: 'calculator',
            items: [
                {
                    name: 'SUM',
                    syntax: '=SUM(sayı1, sayı2, ...)',
                    desc: 'Belirtilen sayıların toplamını hesaplar',
                    examples: [
                        { code: '=SUM(A1:A10)', result: 'A1\'den A10\'a kadar olan hücrelerin toplamı' },
                        { code: '=SUM(B2:B5, D2:D5)', result: 'İki farklı aralığın toplamı' },
                        { code: '=SUM(10, 20, 30)', result: '60' }
                    ],
                    tips: [
                        '💡 Birden fazla aralık toplayabilirsiniz',
                        '💡 Hücre referansları ve sayılar karışık kullanılabilir',
                        '💡 Boş hücreler 0 olarak sayılır'
                    ]
                },
                {
                    name: 'AVERAGE',
                    syntax: '=AVERAGE(sayı1, sayı2, ...)',
                    desc: 'Belirtilen sayıların aritmetik ortalamasını hesaplar',
                    examples: [
                        { code: '=AVERAGE(A1:A10)', result: 'A1-A10 aralığının ortalaması' },
                        { code: '=AVERAGE(B2, C2, D2)', result: 'Üç hücrenin ortalaması' },
                        { code: '=AVERAGE(10, 20, 30)', result: '20' }
                    ],
                    tips: [
                        '💡 Sadece sayısal değerler hesaba katılır',
                        '💡 Boş hücreler göz ardı edilir',
                        '💡 Metin içeren hücreler atlanır'
                    ]
                },
                {
                    name: 'MAX',
                    syntax: '=MAX(sayı1, sayı2, ...)',
                    desc: 'Belirtilen sayılar arasındaki en büyük değeri bulur',
                    examples: [
                        { code: '=MAX(A1:A10)', result: 'Aralıktaki en büyük sayı' },
                        { code: '=MAX(100, B2, C3)', result: 'En büyük değer' }
                    ],
                    tips: [
                        '💡 Negatif sayılarda da çalışır',
                        '💡 Metin ve boş hücreler göz ardı edilir'
                    ]
                },
                {
                    name: 'MIN',
                    syntax: '=MIN(sayı1, sayı2, ...)',
                    desc: 'Belirtilen sayılar arasındaki en küçük değeri bulur',
                    examples: [
                        { code: '=MIN(A1:A10)', result: 'Aralıktaki en küçük sayı' },
                        { code: '=MIN(100, B2, C3)', result: 'En küçük değer' }
                    ],
                    tips: [
                        '💡 Sıfır ve negatif sayılar dahildir',
                        '💡 Metin içeren hücreler atlanır'
                    ]
                },
                {
                    name: 'COUNT',
                    syntax: '=COUNT(değer1, değer2, ...)',
                    desc: 'Belirtilen aralıktaki sayısal değerleri sayar',
                    examples: [
                        { code: '=COUNT(A1:A10)', result: 'Sayı içeren hücre sayısı' },
                        { code: '=COUNT(B2:B20)', result: 'Dolu hücre sayısı' }
                    ],
                    tips: [
                        '💡 Sadece sayıları sayar',
                        '💡 Boş hücreler ve metin sayılmaz',
                        '💡 Tarihler sayılır (sayısal değer)'
                    ]
                },
                {
                    name: 'ROUND',
                    syntax: '=ROUND(sayı, basamak)',
                    desc: 'Sayıyı belirtilen ondalık basamağa yuvarlar',
                    examples: [
                        { code: '=ROUND(3.14159, 2)', result: '3.14' },
                        { code: '=ROUND(125.678, 0)', result: '126' },
                        { code: '=ROUND(A1, 1)', result: 'A1 değerini 1 basamağa yuvarlar' }
                    ],
                    tips: [
                        '💡 Pozitif basamak: ondalık kısım',
                        '💡 0: tam sayıya yuvarla',
                        '💡 Negatif basamak: tam kısmı yuvarla'
                    ]
                },
                {
                    name: 'POWER',
                    syntax: '=POWER(taban, üs)',
                    desc: 'Bir sayının üssünü hesaplar',
                    examples: [
                        { code: '=POWER(2, 3)', result: '8 (2³)' },
                        { code: '=POWER(5, 2)', result: '25 (5²)' },
                        { code: '=POWER(A1, 2)', result: 'A1 değerinin karesi' }
                    ],
                    tips: [
                        '💡 Negatif üs: kesirli sonuç (1/taban^üs)',
                        '💡 0.5 üssü: karekök anlamına gelir'
                    ]
                },
                {
                    name: 'SQRT',
                    syntax: '=SQRT(sayı)',
                    desc: 'Bir sayının karekökünü hesaplar',
                    examples: [
                        { code: '=SQRT(16)', result: '4' },
                        { code: '=SQRT(A1)', result: 'A1 değerinin karekökü' },
                        { code: '=SQRT(25)', result: '5' }
                    ],
                    tips: [
                        '💡 Negatif sayılar hata verir',
                        '💡 POWER(sayı, 0.5) ile aynı sonucu verir'
                    ]
                },
                {
                    name: 'ABS',
                    syntax: '=ABS(sayı)',
                    desc: 'Sayının mutlak değerini (işaretsiz halini) döndürür',
                    examples: [
                        { code: '=ABS(-5)', result: '5' },
                        { code: '=ABS(A1)', result: 'A1\'in mutlak değeri' },
                        { code: '=ABS(10-20)', result: '10' }
                    ],
                    tips: [
                        '💡 Negatif sayıları pozitife çevirir',
                        '💡 Pozitif sayıları değiştirmez',
                        '💡 Fark hesaplamalarında kullanışlı'
                    ]
                }
            ]
        },
        logic: {
            title: '🎯 Mantıksal Fonksiyonlar',
            icon: 'code-branch',
            items: [
                {
                    name: 'IF',
                    syntax: '=IF(koşul, doğruysa, yanlışsa)',
                    desc: 'Koşula göre farklı değerler döndürür',
                    examples: [
                        { code: '=IF(A1>10, "Büyük", "Küçük")', result: 'A1 10\'dan büyükse "Büyük", değilse "Küçük"' },
                        { code: '=IF(B2>=50, "Geçti", "Kaldı")', result: 'Not kontrolü' },
                        { code: '=IF(C1="", "Boş", C1)', result: 'Boş hücre kontrolü' }
                    ],
                    tips: [
                        '💡 İç içe IF kullanabilirsiniz',
                        '💡 Metin değerler çift tırnak içinde',
                        '💡 Sayısal karşılaştırmalar: >, <, >=, <=, ='
                    ]
                },
                {
                    name: 'AND',
                    syntax: '=AND(koşul1, koşul2, ...)',
                    desc: 'Tüm koşullar doğruysa TRUE döndürür',
                    examples: [
                        { code: '=AND(A1>10, A1<20)', result: 'A1, 10-20 arasındaysa TRUE' },
                        { code: '=AND(B2>=50, C2>=50)', result: 'Her iki not da 50+ ise TRUE' }
                    ],
                    tips: [
                        '💡 IF ile birlikte kullanılır',
                        '💡 Tüm koşulların sağlanması gerekir',
                        '💡 Örnek: =IF(AND(A1>0, A1<100), "Geçerli", "Geçersiz")'
                    ]
                },
                {
                    name: 'OR',
                    syntax: '=OR(koşul1, koşul2, ...)',
                    desc: 'Koşullardan en az biri doğruysa TRUE döndürür',
                    examples: [
                        { code: '=OR(A1>100, B1>100)', result: 'Herhangi biri 100+ ise TRUE' },
                        { code: '=OR(C1="A", C1="B")', result: 'C1, A veya B ise TRUE' }
                    ],
                    tips: [
                        '💡 En az bir koşul yeterli',
                        '💡 IF ile birlikte kullanışlı',
                        '💡 Örnek: =IF(OR(A1<0, A1>100), "Hata", "OK")'
                    ]
                }
            ]
        },
        statistics: {
            title: '📊 İstatistik Fonksiyonları',
            icon: 'chart-bar',
            items: [
                {
                    name: 'MEDIAN',
                    syntax: '=MEDIAN(sayı1, sayı2, ...)',
                    desc: 'Sayıların ortanca değerini bulur',
                    examples: [
                        { code: '=MEDIAN(1, 2, 3, 4, 5)', result: '3 (ortadaki değer)' },
                        { code: '=MEDIAN(A1:A10)', result: 'Aralığın medyan değeri' }
                    ],
                    tips: [
                        '💡 Sayıları sıralar ve ortadakini alır',
                        '💡 Çift sayıda değer varsa ortadaki ikisinin ortalaması',
                        '💡 Aşırı değerlerden etkilenmez'
                    ]
                },
                {
                    name: 'MODE',
                    syntax: '=MODE(sayı1, sayı2, ...)',
                    desc: 'En sık tekrar eden değeri bulur',
                    examples: [
                        { code: '=MODE(1, 2, 2, 3, 4)', result: '2 (en çok tekrar eden)' },
                        { code: '=MODE(A1:A20)', result: 'En yaygın değer' }
                    ],
                    tips: [
                        '💡 Tekrar eden değer yoksa hata verir',
                        '💡 İlk bulunan modu döndürür',
                        '💡 Frekans analizi için kullanışlı'
                    ]
                },
                {
                    name: 'STDEV',
                    syntax: '=STDEV(sayı1, sayı2, ...)',
                    desc: 'Standart sapmayı hesaplar',
                    examples: [
                        { code: '=STDEV(A1:A10)', result: 'Verilerin standart sapması' },
                        { code: '=STDEV(B2:B100)', result: 'Dağılım ölçüsü' }
                    ],
                    tips: [
                        '💡 Verilerin dağılımını ölçer',
                        '💡 Düşük değer: veriler birbirine yakın',
                        '💡 Yüksek değer: veriler dağınık'
                    ]
                }
            ]
        },
        text: {
            title: '📝 Metin Fonksiyonları',
            icon: 'font',
            items: [
                {
                    name: 'CONCATENATE',
                    syntax: '=CONCATENATE(metin1, metin2, ...)',
                    desc: 'Metinleri birleştirir',
                    examples: [
                        { code: '=CONCATENATE(A1, " ", B1)', result: 'İki hücreyi boşlukla birleştirir' },
                        { code: '=CONCATENATE("Toplam: ", A1)', result: 'Metin ve sayıyı birleştirir' }
                    ],
                    tips: [
                        '💡 & operatörü ile de yapılabilir: A1 & " " & B1',
                        '💡 Metin değerler çift tırnak içinde',
                        '💡 Boşluk eklemek için " " kullanın'
                    ]
                },
                {
                    name: 'LEFT',
                    syntax: '=LEFT(metin, karakter_sayısı)',
                    desc: 'Metnin soldan belirtilen karakter sayısını alır',
                    examples: [
                        { code: '=LEFT("Merhaba", 3)', result: '"Mer"' },
                        { code: '=LEFT(A1, 5)', result: 'A1\'in ilk 5 karakteri' }
                    ],
                    tips: [
                        '💡 Metin ayrıştırmada kullanışlı',
                        '💡 RIGHT: sağdan al',
                        '💡 MID: ortadan al'
                    ]
                },
                {
                    name: 'UPPER',
                    syntax: '=UPPER(metin)',
                    desc: 'Metni büyük harfe çevirir',
                    examples: [
                        { code: '=UPPER("merhaba")', result: '"MERHABA"' },
                        { code: '=UPPER(A1)', result: 'A1\'i büyük harfe çevirir' }
                    ],
                    tips: [
                        '💡 LOWER: küçük harfe çevirir',
                        '💡 PROPER: her kelimenin ilk harfini büyük yapar',
                        '💡 Standartlaştırma için kullanışlı'
                    ]
                },
                {
                    name: 'LEN',
                    syntax: '=LEN(metin)',
                    desc: 'Metnin karakter sayısını döndürür',
                    examples: [
                        { code: '=LEN("Merhaba")', result: '7' },
                        { code: '=LEN(A1)', result: 'A1\'in karakter sayısı' }
                    ],
                    tips: [
                        '💡 Boşluklar da sayılır',
                        '💡 Doğrulama için kullanışlı',
                        '💡 Karakter limiti kontrolü'
                    ]
                }
            ]
        },
        date: {
            title: '📅 Tarih & Saat Fonksiyonları',
            icon: 'calendar',
            items: [
                {
                    name: 'TODAY',
                    syntax: '=TODAY()',
                    desc: 'Bugünün tarihini döndürür',
                    examples: [
                        { code: '=TODAY()', result: 'Bugünün tarihi' },
                        { code: '=TODAY()+7', result: '7 gün sonrası' },
                        { code: '=TODAY()-30', result: '30 gün öncesi' }
                    ],
                    tips: [
                        '💡 Her gün otomatik güncellenir',
                        '💡 Tarih hesaplamalarında kullanışlı',
                        '💡 NOW(): saat ile birlikte döndürür'
                    ]
                },
                {
                    name: 'YEAR',
                    syntax: '=YEAR(tarih)',
                    desc: 'Tarihten yıl bilgisini çıkarır',
                    examples: [
                        { code: '=YEAR(TODAY())', result: 'Bu yıl' },
                        { code: '=YEAR(A1)', result: 'A1 tarihinin yılı' }
                    ],
                    tips: [
                        '💡 MONTH(): ay bilgisi',
                        '💡 DAY(): gün bilgisi',
                        '💡 Yaş hesaplama için kullanılır'
                    ]
                },
                {
                    name: 'DATEDIF',
                    syntax: '=DATEDIF(başlangıç, bitiş, birim)',
                    desc: 'İki tarih arasındaki farkı hesaplar',
                    examples: [
                        { code: '=DATEDIF(A1, TODAY(), "Y")', result: 'Yıl farkı' },
                        { code: '=DATEDIF(A1, B1, "M")', result: 'Ay farkı' },
                        { code: '=DATEDIF(A1, B1, "D")', result: 'Gün farkı' }
                    ],
                    tips: [
                        '💡 "Y": yıl, "M": ay, "D": gün',
                        '💡 Yaş hesaplama için ideal',
                        '💡 Çalışma süresi hesaplama'
                    ]
                }
            ]
        }
    };

    // Modal HTML oluştur
    const modalHTML = `
        <div id="formulaHelpOverlay" style="
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.6);
            backdrop-filter: blur(5px);
            z-index: 9998;
            animation: fadeIn 0.2s ease-out;
        "></div>
        
        <div id="formulaHelpModal" style="
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: 90%;
            max-width: 1000px;
            max-height: 85vh;
            background: white;
            border-radius: 20px;
            box-shadow: 0 25px 80px rgba(0, 0, 0, 0.3);
            z-index: 9999;
            display: flex;
            flex-direction: column;
            overflow: hidden;
            animation: slideUp 0.3s ease-out;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        ">
            <!-- Header -->
            <div style="
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                padding: 24px 30px;
                color: white;
                display: flex;
                justify-content: space-between;
                align-items: center;
            ">
                <div>
                    <h2 style="margin: 0 0 8px 0; font-size: 24px; font-weight: 700;">
                        📚 Formül Kütüphanesi
                    </h2>
                    <p style="margin: 0; opacity: 0.9; font-size: 14px;">
                        Tüm formüller, örnekler ve ipuçları
                    </p>
                </div>
                <button id="closeFormulaHelp" style="
                    background: rgba(255, 255, 255, 0.2);
                    border: none;
                    width: 40px;
                    height: 40px;
                    border-radius: 50%;
                    color: white;
                    font-size: 24px;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    transition: all 0.2s;
                " onmouseover="this.style.background='rgba(255,255,255,0.3)'; this.style.transform='rotate(90deg)';" 
                   onmouseout="this.style.background='rgba(255,255,255,0.2)'; this.style.transform='rotate(0deg)';">
                    ×
                </button>
            </div>

            <!-- Search Bar -->
            <div style="padding: 20px 30px; background: #f8f9fa; border-bottom: 1px solid #e0e0e0;">
                <div style="position: relative;">
                    <input 
                        type="text" 
                        id="formulaSearch" 
                        placeholder="🔍 Formül ara... (örn: toplam, ortalama, koşul)"
                        style="
                            width: 100%;
                            padding: 14px 45px 14px 16px;
                            border: 2px solid #e0e0e0;
                            border-radius: 12px;
                            font-size: 15px;
                            outline: none;
                            transition: all 0.2s;
                            box-sizing: border-box;
                        "
                        onfocus="this.style.borderColor='#667eea'; this.style.boxShadow='0 0 0 3px rgba(102, 126, 234, 0.1)';"
                        onblur="this.style.borderColor='#e0e0e0'; this.style.boxShadow='none';"
                    >
                </div>
            </div>

            <!-- Content Container -->
            <div style="
                display: flex;
                flex: 1;
                overflow: hidden;
            ">
                <!-- Sidebar Categories -->
                <div id="categorySidebar" style="
                    width: 250px;
                    background: #f8f9fa;
                    border-right: 1px solid #e0e0e0;
                    overflow-y: auto;
                    padding: 20px 0;
                ">
                    ${Object.keys(formulas).map((key, index) => `
                        <button class="category-btn ${index === 0 ? 'active' : ''}" data-category="${key}" style="
                            width: 100%;
                            text-align: left;
                            padding: 16px 24px;
                            border: none;
                            background: ${index === 0 ? 'white' : 'transparent'};
                            border-left: 4px solid ${index === 0 ? '#667eea' : 'transparent'};
                            cursor: pointer;
                            transition: all 0.2s;
                            font-size: 15px;
                            font-weight: ${index === 0 ? '600' : '500'};
                            color: ${index === 0 ? '#667eea' : '#666'};
                            display: flex;
                            align-items: center;
                            gap: 12px;
                        " onmouseover="if(!this.classList.contains('active')) {this.style.background='white'; this.style.borderLeftColor='#ddd';}" 
                           onmouseout="if(!this.classList.contains('active')) {this.style.background='transparent'; this.style.borderLeftColor='transparent';}">
                            <i class="fas fa-${formulas[key].icon}" style="width: 20px;"></i>
                            <span>${formulas[key].title}</span>
                        </button>
                    `).join('')}
                </div>

                <!-- Main Content Area -->
                <div id="formulaContent" style="
                    flex: 1;
                    overflow-y: auto;
                    padding: 30px;
                ">
                    <!-- Dinamik olarak doldurulacak -->
                </div>
            </div>

            <!-- Footer -->
            <div style="
                padding: 16px 30px;
                background: #f8f9fa;
                border-top: 1px solid #e0e0e0;
                display: flex;
                justify-content: space-between;
                align-items: center;
                font-size: 13px;
                color: #666;
            ">
                <div>
                    <kbd style="padding: 4px 8px; background: white; border: 1px solid #ddd; border-radius: 4px; font-family: monospace;">ESC</kbd>
                    <span style="margin-left: 8px;">Kapat</span>
                    <span style="margin: 0 16px;">•</span>
                    <kbd style="padding: 4px 8px; background: white; border: 1px solid #ddd; border-radius: 4px; font-family: monospace;">Ctrl+F</kbd>
                    <span style="margin-left: 8px;">Ara</span>
                </div>
                <div>
                    <span style="opacity: 0.7;">💡 Örneklere tıklayarak kopyalayabilirsiniz</span>
                </div>
            </div>
        </div>

        <style>
            @keyframes fadeIn {
                from { opacity: 0; }
                to { opacity: 1; }
            }
            
            @keyframes slideUp {
                from {
                    opacity: 0;
                    transform: translate(-50%, -45%);
                }
                to {
                    opacity: 1;
                    transform: translate(-50%, -50%);
                }
            }

            #categorySidebar::-webkit-scrollbar,
            #formulaContent::-webkit-scrollbar {
                width: 6px;
            }

            #categorySidebar::-webkit-scrollbar-track,
            #formulaContent::-webkit-scrollbar-track {
                background: #f0f0f0;
            }

            #categorySidebar::-webkit-scrollbar-thumb,
            #formulaContent::-webkit-scrollbar-thumb {
                background: #667eea;
                border-radius: 3px;
            }

            .formula-card {
                background: white;
                border: 2px solid #f0f0f0;
                border-radius: 12px;
                padding: 24px;
                margin-bottom: 20px;
                transition: all 0.2s;
            }

            .formula-card:hover {
                border-color: #667eea;
                box-shadow: 0 4px 16px rgba(102, 126, 234, 0.1);
                transform: translateY(-2px);
            }

            .example-code {
                background: #f8f9fa;
                border-left: 4px solid #667eea;
                padding: 12px 16px;
                border-radius: 0 8px 8px 0;
                font-family: 'Courier New', monospace;
                font-size: 14px;
                color: #333;
                cursor: pointer;
                transition: all 0.2s;
                margin: 8px 0;
            }

            .example-code:hover {
                background: #667eea;
                color: white;
                transform: translateX(4px);
            }

            .tip-badge {
                display: inline-block;
                background: #fff3cd;
                color: #856404;
                padding: 6px 12px;
                border-radius: 6px;
                font-size: 13px;
                margin: 4px 8px 4px 0;
                border-left: 3px solid #ffc107;
            }
        </style>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHTML);

    // İçerik gösterme fonksiyonu
    function showCategory(categoryKey) {
        const category = formulas[categoryKey];
        const contentDiv = document.getElementById('formulaContent');
        
        let html = `
            <h3 style="margin: 0 0 24px 0; color: #333; font-size: 22px; font-weight: 700;">
                ${category.title}
            </h3>
        `;

        category.items.forEach(formula => {
            html += `
                <div class="formula-card">
                    <div style="display: flex; justify-content: space-between; align-items: start; margin-bottom: 16px;">
                        <div>
                            <h4 style="margin: 0 0 8px 0; color: #667eea; font-size: 20px; font-weight: 700;">
                                ${formula.name}
                            </h4>
                            <code style="background: #f0f0f0; padding: 4px 10px; border-radius: 6px; font-size: 13px; color: #666;">
                                ${formula.syntax}
                            </code>
                        </div>
                        <button onclick="
                            navigator.clipboard.writeText('${formula.syntax}');
                            this.innerHTML = '<i class=\\'fas fa-check\\'></i> Kopyalandı!';
                            setTimeout(() => this.innerHTML = '<i class=\\'fas fa-copy\\'></i> Kopyala', 2000);
                        " style="
                            background: #667eea;
                            color: white;
                            border: none;
                            padding: 8px 16px;
                            border-radius: 8px;
                            cursor: pointer;
                            font-size: 13px;
                            white-space: nowrap;
                        ">
                            <i class="fas fa-copy"></i> Kopyala
                        </button>
                    </div>

                    <p style="color: #666; margin: 0 0 20px 0; font-size: 15px; line-height: 1.6;">
                        ${formula.desc}
                    </p>

                    <h5 style="margin: 0 0 12px 0; color: #333; font-size: 15px; font-weight: 600;">
                        📌 Örnekler:
                    </h5>
                    ${formula.examples.map(ex => `
                        <div class="example-code" onclick="
                            navigator.clipboard.writeText('${ex.code}');
                            const original = this.innerHTML;
                            this.innerHTML = '<i class=\\'fas fa-check\\'></i> Kopyalandı!';
                            setTimeout(() => this.innerHTML = original, 1500);
                        " title="Kopyalamak için tıkla">
                            <div style="font-weight: 600; margin-bottom: 4px;">${ex.code}</div>
                            <div style="font-size: 12px; opacity: 0.8;">→ ${ex.result}</div>
                        </div>
                    `).join('')}

                    ${formula.tips && formula.tips.length > 0 ? `
                        <h5 style="margin: 20px 0 12px 0; color: #333; font-size: 15px; font-weight: 600;">
                            💡 İpuçları:
                        </h5>
                        <div>
                            ${formula.tips.map(tip => `
                                <span class="tip-badge">${tip}</span>
                            `).join('')}
                        </div>
                    ` : ''}
                </div>
            `;
        });

        contentDiv.innerHTML = html;
    }

    // İlk kategoriyi göster
    showCategory(Object.keys(formulas)[0]);

    // Kategori değiştirme
    document.querySelectorAll('.category-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            // Tüm butonlardan active sınıfını kaldır
            document.querySelectorAll('.category-btn').forEach(b => {
                b.classList.remove('active');
                b.style.background = 'transparent';
                b.style.borderLeftColor = 'transparent';
                b.style.fontWeight = '500';
                b.style.color = '#666';
            });
            
            // Tıklanan butona active ekle
            this.classList.add('active');
            this.style.background = 'white';
            this.style.borderLeftColor = '#667eea';
            this.style.fontWeight = '600';
            this.style.color = '#667eea';
            
            // İçeriği göster
            showCategory(this.getAttribute('data-category'));
        });
    });

    // Arama fonksiyonu
    const searchInput = document.getElementById('formulaSearch');
    searchInput.addEventListener('input', function() {
        const query = this.value.toLowerCase();
        const contentDiv = document.getElementById('formulaContent');
        
        if (query.length < 2) {
            const activeCategory = document.querySelector('.category-btn.active').getAttribute('data-category');
            showCategory(activeCategory);
            return;
        }

        let results = [];
        Object.keys(formulas).forEach(catKey => {
            formulas[catKey].items.forEach(formula => {
                if (formula.name.toLowerCase().includes(query) ||
                    formula.desc.toLowerCase().includes(query) ||
                    formula.syntax.toLowerCase().includes(query)) {
                    results.push({ category: formulas[catKey].title, formula: formula });
                }
            });
        });

        if (results.length === 0) {
            contentDiv.innerHTML = `
                <div style="text-align: center; padding: 60px 20px; color: #999;">
                    <i class="fas fa-search" style="font-size: 48px; margin-bottom: 16px; opacity: 0.3;"></i>
                    <p style="font-size: 16px;">Sonuç bulunamadı</p>
                    <p style="font-size: 14px;">Farklı bir anahtar kelime deneyin</p>
                </div>
            `;
            return;
        }

        let html = `
            <h3 style="margin: 0 0 24px 0; color: #333; font-size: 22px;">
                🔍 Arama Sonuçları (${results.length})
            </h3>
        `;

        results.forEach(result => {
            const formula = result.formula;
            html += `
                <div class="formula-card">
                    <div style="margin-bottom: 12px;">
                        <span style="background: #667eea; color: white; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 600;">
                            ${result.category}
                        </span>
                    </div>
                    
                    <h4 style="margin: 0 0 8px 0; color: #667eea; font-size: 20px; font-weight: 700;">
                        ${formula.name}
                    </h4>
                    <code style="background: #f0f0f0; padding: 4px 10px; border-radius: 6px; font-size: 13px; color: #666; display: inline-block; margin-bottom: 12px;">
                        ${formula.syntax}
                    </code>
                    <p style="color: #666; margin: 0; font-size: 15px;">
                        ${formula.desc}
                    </p>
                    
                    ${formula.examples.length > 0 ? `
                        <div class="example-code" style="margin-top: 12px;" onclick="
                            navigator.clipboard.writeText('${formula.examples[0].code}');
                            this.style.background='#4CAF50'; this.style.color='white';
                            setTimeout(() => {this.style.background='#f8f9fa'; this.style.color='#333';}, 1000);
                        ">
                            ${formula.examples[0].code}
                        </div>
                    ` : ''}
                </div>
            `;
        });

        contentDiv.innerHTML = html;
    });

    // Klavye kısayolları
    document.addEventListener('keydown', function escHandler(e) {
        if (e.key === 'Escape') {
            closeModal();
            document.removeEventListener('keydown', escHandler);
        }
        if (e.ctrlKey && e.key === 'f') {
            e.preventDefault();
            searchInput.focus();
        }
    });

    // Modal kapatma
    function closeModal() {
        const modal = document.getElementById('formulaHelpModal');
        const overlay = document.getElementById('formulaHelpOverlay');
        if (modal) modal.remove();
        if (overlay) overlay.remove();
    }

    document.getElementById('closeFormulaHelp').addEventListener('click', closeModal);
    document.getElementById('formulaHelpOverlay').addEventListener('click', closeModal);

    console.log('✅ Gelişmiş Formül Yardım Sistemi yüklendi!');
};

// F1 kısayolu
document.addEventListener('keydown', (e) => {
    if (e.key === 'F1') {
        e.preventDefault();
        if (typeof window.showFormulaHelp === 'function') {
            window.showFormulaHelp();
        }
    }
});
//==============  GRAFİK FONKSİYONLARI ====================
function addChartButtonToFormulaBar() {
    const formulaBar = document.querySelector('.formula-bar');
    if (!formulaBar) {
        console.log('❌ Formül bar bulunamadı!');
        return;
    }
    
    if (document.getElementById('showChartBtn')) {
        return;
    }
    
    // TEK BUTON - hem ana buton hem dropdown içerir
    const chartButton = document.createElement('button');
    chartButton.id = 'showChartBtn';
    chartButton.innerHTML = '📊 Grafik Göster <span style="font-size: 10px; margin-left: 5px;"></span>';
    chartButton.title = 'Tablodaki verileri grafik olarak göster (tür seçmek için tıklayın)';
    
    chartButton.style.cssText = `
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 6px;
        cursor: pointer;
        font-weight: bold;
        font-size: 14px;
        transition: all 0.3s;
        white-space: nowrap;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-left: 10px;
        position: relative;
    `;
    
    // Hover efekti
    chartButton.onmouseenter = function() {
        this.style.transform = 'translateY(-2px)';
        this.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.3)';
    };
    
    chartButton.onmouseleave = function() {
        this.style.transform = 'translateY(0)';
        this.style.boxShadow = 'none';
    };
    
    // Tek tıklama ile direkt bar grafik göster
    // Uzun tıklama (veya sağ tık) ile grafik türü menüsü göster
    let clickTimer;
    let isLongPress = false;
    
    chartButton.onmousedown = function(e) {
        isLongPress = false;
        clickTimer = setTimeout(() => {
            isLongPress = true;
            showChartTypeDropdown(e);
        }, 500); // 0.5 saniye basılı tutunca menü aç
    };
    
    chartButton.onmouseup = function(e) {
        clearTimeout(clickTimer);
        if (!isLongPress && e.button === 0) { // Sol tık ve kısa basma
            showDataChart('bar'); // Varsayılan bar grafik
        }
    };
    
    // Sağ tık için de menü göster
    chartButton.oncontextmenu = function(e) {
        e.preventDefault();
        showChartTypeDropdown(e);
        return false;
    };
    
    formulaBar.appendChild(chartButton);
    
    console.log('✅ Grafik butonu eklendi');
}

// ==================== GRAFİK TÜRÜ AÇILIR MENÜSÜ ====================
function showChartTypeDropdown(event) {
    // Önceki menüyü temizle
    const existingMenu = document.getElementById('chartTypeDropdown');
    if (existingMenu) {
        existingMenu.remove();
        return;
    }
    
    const dropdown = document.createElement('div');
    dropdown.id = 'chartTypeDropdown';
    dropdown.style.cssText = `
        position: fixed;
        background: white;
        border-radius: 8px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        z-index: 10001;
        min-width: 220px;
        overflow: hidden;
        border: 1px solid #e0e0e0;
    `;
    
    const chartTypes = [
        { id: 'bar', name: 'Bar Grafik', icon: '📊', desc: 'Dikey çubuklarla gösterim', shortcut: 'Varsayılan' },
        { id: 'line', name: 'Çizgi Grafik', icon: '📈', desc: 'Zaman serisi trendi', shortcut: '' },
        { id: 'pie', name: 'Pasta Grafik', icon: '🥧', desc: 'Oranları göster', shortcut: '' },
        { id: 'doughnut', name: 'Halka Grafik', icon: '🍩', desc: 'Pasta grafiğin halka versiyonu', shortcut: '' }
    ];
    
    dropdown.innerHTML = `
        <div style="
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 15px;
            font-weight: bold;
            font-size: 14px;
            display: flex;
            align-items: center;
            gap: 10px;
        ">
            📊 Grafik Türünü Seç
        </div>
        <div style="max-height: 300px; overflow-y: auto;">
            ${chartTypes.map(type => `
                <div onclick="selectChartType('${type.id}')" style="
                    padding: 12px 15px;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    gap: 12px;
                    border-bottom: 1px solid #f0f0f0;
                    transition: all 0.2s;
                    position: relative;
                " onmouseenter="this.style.background='#f8f9fa'" 
                 onmouseleave="this.style.background='white'">
                    <div style="font-size: 20px; width: 30px;">${type.icon}</div>
                    <div style="flex: 1;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div style="font-weight: 600; color: #333;">${type.name}</div>
                            ${type.shortcut ? `<div style="font-size: 11px; color: #667eea; background: #e3f2fd; padding: 2px 6px; border-radius: 10px;">${type.shortcut}</div>` : ''}
                        </div>
                        <div style="font-size: 12px; color: #666; margin-top: 4px;">${type.desc}</div>
                    </div>
                </div>
            `).join('')}
            <div style="padding: 12px 15px; border-top: 1px solid #f0f0f0; background: #f8f9fa; font-size: 12px; color: #666;">
                <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 5px;">
                    <div style="width: 20px; text-align: center;">🖱️</div>
                    <div>Kısa tık: Bar Grafik</div>
                </div>
                <div style="display: flex; align-items: center; gap: 8px;">
                    <div style="width: 20px; text-align: center;">⏱️</div>
                    <div>Uzun tık/Sağ tık: Tür seç</div>
                </div>
            </div>
        </div>
    `;
    
    // Konumlandır (fare pozisyonuna göre)
    const x = event.clientX || event.pageX;
    const y = event.clientY || event.pageY;
    
    dropdown.style.left = Math.min(x, window.innerWidth - 250) + 'px';
    dropdown.style.top = Math.min(y + 5, window.innerHeight - 350) + 'px';
    
    document.body.appendChild(dropdown);
    
    // Dışarı tıklayınca kapat
    setTimeout(() => {
        document.addEventListener('click', closeDropdownOnClickOutside);
    }, 10);
}

function selectChartType(chartType) {
    closeChartTypeDropdown();
    showDataChart(chartType);
}

function closeChartTypeDropdown() {
    const dropdown = document.getElementById('chartTypeDropdown');
    if (dropdown) dropdown.remove();
    document.removeEventListener('click', closeDropdownOnClickOutside);
}

function closeDropdownOnClickOutside(event) {
    const dropdown = document.getElementById('chartTypeDropdown');
    const button = document.getElementById('showChartBtn');
    
    if (dropdown && 
        !dropdown.contains(event.target) && 
        !button.contains(event.target)) {
        closeChartTypeDropdown();
    }
}

// ==================== GRAFİK GÖSTER ====================
function showDataChart(chartType = 'bar') {
    console.log(`📈 ${chartType} grafiği gösteriliyor...`);
    
    const allCells = document.querySelectorAll('#container input[type="text"]');
    if (!allCells || allCells.length === 0) {
        alert('❌ Tablo bulunamadı!');
        return;
    }
    
    const data = {
        numericValues: [],
        textValues: [],
        formulaValues: [],
        cellCount: allCells.length,
        filledCells: 0
    };
    
    allCells.forEach(cell => {
        const value = cell.value.trim();
        if (value === '') return;
        
        data.filledCells++;
        const numValue = parseFloat(value);
        
        if (!isNaN(numValue)) {
            data.numericValues.push({
                id: cell.id,
                value: numValue,
                rawValue: value,
                isFormula: false
            });
        } else if (value.startsWith('=')) {
            data.formulaValues.push({
                id: cell.id,
                value: value,
                isFormula: true
            });
        } else {
            data.textValues.push({
                id: cell.id,
                value: value,
                isFormula: false
            });
        }
    });
    
    if (data.numericValues.length < 2) {
        showSimpleAlert(`
            <div style="text-align: center; padding: 20px;">
                <div style="font-size: 48px; margin-bottom: 10px;">📊</div>
                <h3 style="margin: 0 0 10px 0; color: #333;">Yeterli Veri Yok!</h3>
                <p style="color: #666; margin: 0 0 15px 0;">
                    Grafik oluşturmak için en az <strong>2 sayısal değer</strong> gerekli.
                </p>
                <p style="color: #888; font-size: 14px;">
                    Bulunan: ${data.numericValues.length} sayı, ${data.textValues.length} metin, ${data.formulaValues.length} formül
                </p>
            </div>
        `);
        return;
    }
    
    openChartWindow(data, chartType);
}

// ==================== BASİT UYARI GÖSTERİCİ ====================
function showSimpleAlert(htmlContent) {
    const existing = document.getElementById('simpleAlert');
    if (existing) existing.remove();
    
    const alertDiv = document.createElement('div');
    alertDiv.id = 'simpleAlert';
    alertDiv.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        border-radius: 12px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        z-index: 10000;
        padding: 0;
        overflow: hidden;
        min-width: 300px;
        max-width: 400px;
    `;
    
    alertDiv.innerHTML = `
        <div style="background: linear-gradient(135deg, #FF6B6B 0%, #FF9E6D 100%); padding: 20px; color: white;">
            <h3 style="margin: 0; font-size: 18px;">⚠️ Bilgi</h3>
        </div>
        <div style="padding: 25px;">
            ${htmlContent}
        </div>
        <div style="padding: 15px 25px; background: #f8f9fa; border-top: 1px solid #eee; text-align: center;">
            <button onclick="closeSimpleAlert()" style="
                padding: 8px 25px;
                background: #667eea;
                color: white;
                border: none;
                border-radius: 6px;
                cursor: pointer;
                font-weight: bold;
            ">
                Tamam
            </button>
        </div>
    `;
    
    document.body.appendChild(alertDiv);
    
    const overlay = document.createElement('div');
    overlay.id = 'alertOverlay';
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.5);
        z-index: 9999;
    `;
    overlay.onclick = closeSimpleAlert;
    document.body.appendChild(overlay);
}

function closeSimpleAlert() {
    const alert = document.getElementById('simpleAlert');
    const overlay = document.getElementById('alertOverlay');
    
    if (alert) alert.remove();
    if (overlay) overlay.remove();
}

// ==================== GRAFİK PENCERESİ ====================
function openChartWindow(data, chartType = 'bar') {
    closeChartWindow();
    
    const chartTitles = {
        'bar': 'Bar Grafik',
        'line': 'Çizgi Grafik',
        'pie': 'Pasta Grafik',
        'doughnut': 'Halka Grafik'
    };
    
    const chartIcons = {
        'bar': '📊',
        'line': '📈',
        'pie': '🥧',
        'doughnut': '🍩'
    };
    
    const sortedValues = [...data.numericValues].sort((a, b) => b.value - a.value);
    const topValues = sortedValues.slice(0, Math.min(10, sortedValues.length));
    
    const stats = {
        total: data.numericValues.reduce((sum, item) => sum + item.value, 0),
        average: data.numericValues.reduce((sum, item) => sum + item.value, 0) / data.numericValues.length,
        highest: Math.max(...data.numericValues.map(d => d.value)),
        lowest: Math.min(...data.numericValues.map(d => d.value)),
        count: data.numericValues.length,
        positiveCount: data.numericValues.filter(d => d.value > 0).length,
        negativeCount: data.numericValues.filter(d => d.value < 0).length,
        zeroCount: data.numericValues.filter(d => d.value === 0).length
    };
    
    const chartWindow = document.createElement('div');
    chartWindow.id = 'chartWindow';
    chartWindow.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 900px;
        max-width: 90vw;
        max-height: 85vh;
        background: white;
        border-radius: 16px;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        z-index: 10000;
        overflow: hidden;
        font-family: -apple-system, BlinkMacSystemFont, sans-serif;
    `;
    
    let chartContent = '';
    switch(chartType) {
        case 'bar':
            chartContent = getBarChartContent(topValues, Math.max(...topValues.map(d => d.value)));
            break;
        case 'line':
            chartContent = getLineChartContent(topValues);
            break;
        case 'pie':
            chartContent = getPieChartContent(topValues);
            break;
        case 'doughnut':
            chartContent = getDoughnutChartContent(topValues);
            break;
        default:
            chartContent = getBarChartContent(topValues, Math.max(...topValues.map(d => d.value)));
    }
    
    chartWindow.innerHTML = `
        <!-- Header -->
        <div style="
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 25px 30px;
            color: white;
            display: flex;
            justify-content: space-between;
            align-items: center;
        ">
            <div style="display: flex; align-items: center; gap: 12px;">
                <div style="font-size: 28px;">${chartIcons[chartType]}</div>
                <div>
                    <h2 style="margin: 0 0 8px 0; font-size: 22px; font-weight: 700;">
                        ${chartTitles[chartType]}
                    </h2>
                    <div style="display: flex; gap: 15px; font-size: 13px; opacity: 0.9;">
                        <span>${data.cellCount} hücre</span>
                        <span>•</span>
                        <span>${data.filledCells} dolu</span>
                        <span>•</span>
                        <span>${stats.count} sayı</span>
                    </div>
                </div>
            </div>
            <div style="display: flex; align-items: center; gap: 10px;">
                <div id="chartTypeSelector" style="
                    background: rgba(255,255,255,0.2);
                    border-radius: 8px;
                    padding: 6px 12px;
                    font-size: 13px;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                " onclick="showChartTypeSelectorInWindow()">
                    ${chartIcons[chartType]} ${chartTitles[chartType]}
                    <span style="font-size: 10px;">▼</span>
                </div>
                <button onclick="closeChartWindow()" style="
                    background: rgba(255,255,255,0.2);
                    border: none;
                    width: 36px;
                    height: 36px;
                    border-radius: 50%;
                    color: white;
                    font-size: 20px;
                    cursor: pointer;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    transition: all 0.2s;
                ">×</button>
            </div>
        </div>
        
        <!-- Content -->
        <div style="padding: 25px; max-height: 60vh; overflow-y: auto;">
            <!-- Seçilen Grafik -->
            <div style="margin-bottom: 30px;">
                ${chartContent}
            </div>
            
            <!-- Statistics -->
            <div style="margin-bottom: 25px;">
                <h3 style="margin: 0 0 15px 0; color: #333; font-size: 18px; font-weight: 600;">
                    📈 Detaylı İstatistikler
                </h3>
                <div style="
                    display: grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap: 12px;
                ">
                    <div style="
                        background: linear-gradient(135deg, #667eea, #764ba2);
                        color: white;
                        padding: 15px;
                        border-radius: 10px;
                        text-align: center;
                    ">
                        <div style="font-size: 12px; opacity: 0.9; margin-bottom: 5px;">TOPLAM</div>
                        <div style="font-size: 22px; font-weight: 800;">${stats.total.toFixed(2)}</div>
                    </div>
                    <div style="
                        background: linear-gradient(135deg, #10b981, #059669);
                        color: white;
                        padding: 15px;
                        border-radius: 10px;
                        text-align: center;
                    ">
                        <div style="font-size: 12px; opacity: 0.9; margin-bottom: 5px;">ORTALAMA</div>
                        <div style="font-size: 22px; font-weight: 800;">${stats.average.toFixed(2)}</div>
                    </div>
                    <div style="
                        background: linear-gradient(135deg, #f59e0b, #d97706);
                        color: white;
                        padding: 15px;
                        border-radius: 10px;
                        text-align: center;
                    ">
                        <div style="font-size: 12px; opacity: 0.9; margin-bottom: 5px;">EN YÜKSEK</div>
                        <div style="font-size: 22px; font-weight: 800;">${stats.highest.toFixed(2)}</div>
                    </div>
                    <div style="
                        background: linear-gradient(135deg, #8b5cf6, #7c3aed);
                        color: white;
                        padding: 15px;
                        border-radius: 10px;
                        text-align: center;
                    ">
                        <div style="font-size: 12px; opacity: 0.9; margin-bottom: 5px;">EN DÜŞÜK</div>
                        <div style="font-size: 22px; font-weight: 800;">${stats.lowest.toFixed(2)}</div>
                    </div>
                </div>
            </div>
            
            <!-- Data Table -->
            <div>
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                    <h3 style="margin: 0; color: #333; font-size: 18px; font-weight: 600;">
                        📋 Veri Listesi
                    </h3>
                    <div style="font-size: 13px; color: #666;">
                        Pozitif: ${stats.positiveCount} • Negatif: ${stats.negativeCount} • Sıfır: ${stats.zeroCount}
                    </div>
                </div>
                <div style="
                    background: #f8f9fa;
                    border-radius: 12px;
                    padding: 20px;
                    max-height: 250px;
                    overflow-y: auto;
                ">
                    <div style="
                        display: grid;
                        grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));
                        gap: 12px;
                    ">
                        ${data.numericValues.slice(0, 20).map(item => `
                            <div style="
                                background: white;
                                padding: 15px;
                                border-radius: 8px;
                                border-left: 4px solid ${item.value > 0 ? '#10b981' : item.value < 0 ? '#ef4444' : '#64748b'};
                                box-shadow: 0 2px 8px rgba(0,0,0,0.05);
                            ">
                                <div style="
                                    display: flex;
                                    justify-content: space-between;
                                    align-items: center;
                                    margin-bottom: 8px;
                                ">
                                    <div style="font-weight: 700; color: #333; font-size: 14px;">
                                        ${item.id}
                                    </div>
                                    <div style="
                                        padding: 2px 8px;
                                        background: ${item.value > 0 ? '#d1fae5' : item.value < 0 ? '#fee2e2' : '#e2e8f0'};
                                        color: ${item.value > 0 ? '#065f46' : item.value < 0 ? '#991b1b' : '#475569'};
                                        border-radius: 12px;
                                        font-size: 11px;
                                        font-weight: 600;
                                    ">
                                        ${item.value > 0 ? 'POZİTİF' : item.value < 0 ? 'NEGATİF' : 'SIFIR'}
                                    </div>
                                </div>
                                <div style="font-size: 20px; font-weight: 800; color: ${item.value > 0 ? '#10b981' : item.value < 0 ? '#ef4444' : '#64748b'};">
                                    ${item.value % 1 === 0 ? item.value : item.value.toFixed(2)}
                                </div>
                            </div>
                        `).join('')}
                    </div>
                    ${data.numericValues.length > 20 ? `
                        <div style="text-align: center; margin-top: 15px; padding: 10px; color: #666; font-size: 14px;">
                            + ${data.numericValues.length - 20} daha fazla veri...
                        </div>
                    ` : ''}
                </div>
            </div>
        </div>
        
        <!-- Footer -->
        <div style="
            padding: 20px 30px;
            background: #f8f9fa;
            border-top: 1px solid #e0e0e0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        ">
            <div style="font-size: 13px; color: #666;">
                <span style="font-weight: 600;">💡 İpucu:</span> Grafik türünü değiştirmek için üstteki butonu kullanın
            </div>
            <div style="display: flex; gap: 10px;">
                <button onclick="exportChartData()" style="
                    padding: 10px 20px;
                    background: linear-gradient(135deg, #10b981, #059669);
                    color: white;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-weight: 600;
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    font-size: 14px;
                ">
                    📥 CSV İndir
                </button>
                <button onclick="closeChartWindow()" style="
                    padding: 10px 20px;
                    background: linear-gradient(135deg, #64748b, #475569);
                    color: white;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-weight: 600;
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    font-size: 14px;
                ">
                    Kapat
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(chartWindow);
    
    // Overlay ekle
    const overlay = document.createElement('div');
    overlay.id = 'chartOverlay';
    overlay.style.cssText = `
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(0,0,0,0.5);
        backdrop-filter: blur(4px);
        z-index: 9999;
    `;
    overlay.onclick = closeChartWindow;
    document.body.appendChild(overlay);
    
    // Animasyonları başlat
    setTimeout(() => {
        animateChart(chartType);
    }, 300);
}

// ==================== GRAFİK TÜRLERİ İÇERİKLERİ ====================
function getBarChartContent(values, maxValue) {
    return `
        <h3 style="margin: 0 0 15px 0; color: #333; font-size: 18px; font-weight: 600;">
            Dikey Bar Grafik
        </h3>
        <div id="barChart" style="min-height: 250px;">
            ${values.map((item, index) => {
                const barWidth = (item.value / maxValue) * 100;
                const colors = [
                    '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
                    '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9'
                ];
                const color = colors[index % colors.length];
                
                return `
                    <div style="margin-bottom: 15px;">
                        <div style="display: flex; align-items: center; margin-bottom: 8px;">
                            <div style="
                                width: 24px;
                                height: 24px;
                                background: ${color};
                                color: white;
                                border-radius: 6px;
                                display: flex;
                                align-items: center;
                                justify-content: center;
                                font-weight: bold;
                                font-size: 12px;
                                margin-right: 10px;
                            ">${index + 1}</div>
                            <div style="flex: 1; font-weight: 500; color: #333;">${item.id}</div>
                            <div style="font-weight: bold; color: #667eea; font-size: 16px;">
                                ${item.value % 1 === 0 ? item.value : item.value.toFixed(2)}
                            </div>
                        </div>
                        <div style="
                            width: 100%;
                            height: 28px;
                            background: #f0f0f0;
                            border-radius: 14px;
                            overflow: hidden;
                            position: relative;
                        ">
                            <div class="bar-animation" style="
                                width: 0%;
                                height: 100%;
                                background: ${color};
                                border-radius: 14px;
                                display: flex;
                                align-items: center;
                                padding-left: 15px;
                                color: white;
                                font-weight: bold;
                                font-size: 13px;
                            ">
                                ${item.value % 1 === 0 ? item.value : item.value.toFixed(2)}
                            </div>
                        </div>
                    </div>
                `;
            }).join('')}
        </div>
    `;
}

function getLineChartContent(values) {
    const max = Math.max(...values.map(v => v.value));
    const min = Math.min(...values.map(v => v.value));
    const range = max - min || 1; // Sıfır bölme hatası önlemi
    
    const points = values.map((item, index) => {
        const x = (index / Math.max(values.length - 1, 1)) * 100;
        const y = 100 - ((item.value - min) / range) * 100;
        return `${x}% ${y}%`;
    }).join(', ');
    
    return `
        <h3 style="margin: 0 0 15px 0; color: #333; font-size: 18px; font-weight: 600;">
            Çizgi Grafik
        </h3>
        <div style="position: relative; height: 250px; background: #f8f9fa; border-radius: 12px; padding: 20px;">
            <div style="position: relative; width: 100%; height: 100%;">
                <!-- Grid -->
                <div style="position: absolute; width: 100%; height: 100%; display: grid; grid-template-columns: repeat(10, 1fr); grid-template-rows: repeat(10, 1fr);">
                    ${Array.from({length: 11}).map((_, i) => `
                        <div style="grid-column: ${i + 1} / span 1; grid-row: 1 / span 10; border-right: 1px solid #e0e0e0;"></div>
                        <div style="grid-column: 1 / span 10; grid-row: ${i + 1} / span 1; border-bottom: 1px solid #e0e0e0;"></div>
                    `).join('')}
                </div>
                
                <!-- Line -->
                <svg width="100%" height="100%" style="position: absolute;">
                    <polyline 
                        points="${points}"
                        fill="none"
                        stroke="#667eea"
                        stroke-width="3"
                        stroke-linecap="round"
                        stroke-linejoin="round"
                        class="line-animation"
                    />
                    ${values.map((_, index) => {
                        const x = (index / Math.max(values.length - 1, 1)) * 100;
                        const y = 100 - ((values[index].value - min) / range) * 100;
                        return `
                            <circle 
                                cx="${x}%" 
                                cy="${y}%" 
                                r="6" 
                                fill="#667eea"
                                stroke="white"
                                stroke-width="2"
                                class="circle-animation"
                                style="animation-delay: ${index * 0.1}s"
                            />
                        `;
                    }).join('')}
                </svg>
                
                <!-- Labels -->
                <div style="position: absolute; bottom: -25px; left: 0; right: 0; display: flex; justify-content: space-between;">
                    ${values.map((item, index) => `
                        <div style="
                            font-size: 11px;
                            color: #666;
                            transform: translateX(-50%);
                            position: absolute;
                            left: ${(index / Math.max(values.length - 1, 1)) * 100}%;
                        ">${item.id}</div>
                    `).join('')}
                </div>
                
                <!-- Value Labels -->
                ${values.map((item, index) => {
                    const x = (index / Math.max(values.length - 1, 1)) * 100;
                    const y = 100 - ((item.value - min) / range) * 100;
                    return `
                        <div style="
                            position: absolute;
                            left: ${x}%;
                            top: ${y}%;
                            transform: translate(-50%, -150%);
                            background: #667eea;
                            color: white;
                            padding: 4px 8px;
                            border-radius: 6px;
                            font-size: 12px;
                            font-weight: bold;
                            opacity: 0;
                            animation: fadeIn 0.5s ${index * 0.1 + 1}s forwards;
                        ">
                            ${item.value % 1 === 0 ? item.value : item.value.toFixed(2)}
                        </div>
                    `;
                }).join('')}
            </div>
        </div>
        <style>
            @keyframes fadeIn {
                to { opacity: 1; }
            }
            .line-animation {
                stroke-dasharray: 1000;
                stroke-dashoffset: 1000;
                animation: drawLine 2s forwards;
            }
            .circle-animation {
                opacity: 0;
                animation: fadeInScale 0.5s forwards;
            }
            @keyframes drawLine {
                to { stroke-dashoffset: 0; }
            }
            @keyframes fadeInScale {
                0% { opacity: 0; transform: scale(0); }
                100% { opacity: 1; transform: scale(1); }
            }
        </style>
    `;
}

function getPieChartContent(values) {
    const total = values.reduce((sum, item) => sum + Math.abs(item.value), 0) || 1;
    let cumulativeAngle = 0;
    
    const colors = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
        '#DDA0DD', '#98D8C8', '#F7DC6F', '#BB8FCE', '#85C1E9'
    ];
    
    const segments = values.map((item, index) => {
        const percentage = (Math.abs(item.value) / total) * 100;
        const angle = (percentage / 100) * 360;
        const startAngle = cumulativeAngle;
        cumulativeAngle += angle;
        
        const largeArc = angle > 180 ? 1 : 0;
        const startX = 50 + 40 * Math.cos((startAngle - 90) * Math.PI / 180);
        const startY = 50 + 40 * Math.sin((startAngle - 90) * Math.PI / 180);
        const endX = 50 + 40 * Math.cos((startAngle + angle - 90) * Math.PI / 180);
        const endY = 50 + 40 * Math.sin((startAngle + angle - 90) * Math.PI / 180);
        
        return {
            ...item,
            percentage,
            color: colors[index % colors.length],
            path: `M50,50 L${startX},${startY} A40,40 0 ${largeArc},1 ${endX},${endY} Z`
        };
    });
    
    return `
        <h3 style="margin: 0 0 15px 0; color: #333; font-size: 18px; font-weight: 600;">
            Pasta Grafik
        </h3>
        <div style="display: flex; gap: 30px; align-items: center;">
            <div style="flex: 1; position: relative; height: 250px;">
                <svg width="100%" height="100%" viewBox="0 0 100 100">
                    ${segments.map((seg, index) => `
                        <path 
                            d="${seg.path}" 
                            fill="${seg.color}"
                            stroke="white"
                            stroke-width="2"
                            class="pie-segment"
                            style="animation-delay: ${index * 0.2}s"
                        />
                    `).join('')}
                </svg>
                <div style="
                    position: absolute;
                    top: 50%;
                    left: 50%;
                    transform: translate(-50%, -50%);
                    background: white;
                    width: 60px;
                    height: 60px;
                    border-radius: 50%;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-weight: bold;
                    color: #667eea;
                    font-size: 18px;
                ">
                    ${segments.length}
                </div>
            </div>
            
            <div style="flex: 1; max-height: 250px; overflow-y: auto;">
                <div style="display: flex; flex-direction: column; gap: 10px;">
                    ${segments.map((seg, index) => `
                        <div style="
                            background: white;
                            padding: 12px;
                            border-radius: 8px;
                            border-left: 4px solid ${seg.color};
                            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
                        ">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
                                <div style="font-weight: 700; color: #333; font-size: 14px;">
                                    ${seg.id}
                                </div>
                                <div style="
                                    padding: 2px 8px;
                                    background: ${seg.color};
                                    color: white;
                                    border-radius: 12px;
                                    font-size: 11px;
                                    font-weight: 600;
                                ">
                                    ${seg.percentage.toFixed(1)}%
                                </div>
                            </div>
                            <div style="font-size: 16px; font-weight: 800; color: #333;">
                                ${seg.value % 1 === 0 ? seg.value : seg.value.toFixed(2)}
                            </div>
                            <div style="
                                width: 100%;
                                height: 6px;
                                background: #f0f0f0;
                                border-radius: 3px;
                                margin-top: 8px;
                                overflow: hidden;
                            ">
                                <div style="
                                    width: 0%;
                                    height: 100%;
                                    background: ${seg.color};
                                    border-radius: 3px;
                                    animation: fillBar 1s ${index * 0.1 + 0.5}s forwards;
                                " data-width="${seg.percentage}%"></div>
                            </div>
                        </div>
                    `).join('')}
                </div>
            </div>
        </div>
        <style>
            .pie-segment {
                opacity: 0;
                animation: pieGrow 0.5s forwards;
            }
            @keyframes pieGrow {
                0% { transform: scale(0); opacity: 0; }
                100% { transform: scale(1); opacity: 1; }
            }
            @keyframes fillBar {
                to { width: var(--target-width); }
            }
        </style>
        <script>
            // Dinamik width için CSS değişkeni ayarla
            setTimeout(() => {
                document.querySelectorAll('[data-width]').forEach(bar => {
                    bar.style.setProperty('--target-width', bar.getAttribute('data-width'));
                });
            }, 100);
        </script>
    `;
}

function getDoughnutChartContent(values) {
    return getPieChartContent(values).replace('Pasta Grafik', 'Halka Grafik');
}

// ==================== GRAFİK TÜRÜ SEÇİCİ (PENCERE İÇİ) ====================
function showChartTypeSelectorInWindow() {
    const chartTypes = [
        { id: 'bar', name: 'Bar Grafik', icon: '📊' },
        { id: 'line', name: 'Çizgi Grafik', icon: '📈' },
        { id: 'pie', name: 'Pasta Grafik', icon: '🥧' },
        { id: 'doughnut', name: 'Halka Grafik', icon: '🍩' }
    ];
    
    const selector = document.createElement('div');
    selector.id = 'chartTypeSelectorMenu';
    selector.style.cssText = `
        position: absolute;
        background: white;
        border-radius: 8px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        z-index: 10002;
        min-width: 180px;
        overflow: hidden;
        border: 1px solid #e0e0e0;
    `;
    
    selector.innerHTML = chartTypes.map(type => `
        <div onclick="switchChartType('${type.id}')" style="
            padding: 10px 15px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 10px;
            border-bottom: 1px solid #f0f0f0;
            transition: all 0.2s;
        " onmouseenter="this.style.background='#f8f9fa'" 
         onmouseleave="this.style.background='white'">
            <div style="font-size: 18px;">${type.icon}</div>
            <div style="font-weight: 500; color: #333;">${type.name}</div>
        </div>
    `).join('');
    
    const trigger = document.getElementById('chartTypeSelector');
    const rect = trigger.getBoundingClientRect();
    selector.style.top = (rect.bottom + window.scrollY) + 'px';
    selector.style.left = (rect.left + window.scrollX) + 'px';
    
    document.body.appendChild(selector);
    
    setTimeout(() => {
        document.addEventListener('click', closeChartTypeSelectorInWindow);
    }, 10);
}

function switchChartType(chartType) {
    closeChartTypeSelectorInWindow();
    
    const allCells = document.querySelectorAll('#container input[type="text"]');
    const data = {
        numericValues: [],
        textValues: [],
        formulaValues: [],
        cellCount: allCells.length,
        filledCells: 0
    };
    
    allCells.forEach(cell => {
        const value = cell.value.trim();
        if (value === '') return;
        
        data.filledCells++;
        const numValue = parseFloat(value);
        
        if (!isNaN(numValue)) {
            data.numericValues.push({
                id: cell.id,
                value: numValue,
                rawValue: value,
                isFormula: false
            });
        } else if (value.startsWith('=')) {
            data.formulaValues.push({
                id: cell.id,
                value: value,
                isFormula: true
            });
        } else {
            data.textValues.push({
                id: cell.id,
                value: value,
                isFormula: false
            });
        }
    });
    
    closeChartWindow();
    setTimeout(() => openChartWindow(data, chartType), 50);
}

function closeChartTypeSelectorInWindow() {
    const selector = document.getElementById('chartTypeSelectorMenu');
    if (selector) selector.remove();
    document.removeEventListener('click', closeChartTypeSelectorInWindow);
}

// ==================== ANİMASYONLAR ====================
function animateChart(chartType) {
    switch(chartType) {
        case 'bar':
            const bars = document.querySelectorAll('.bar-animation');
            bars.forEach((bar, index) => {
                const computedStyle = window.getComputedStyle(bar.parentElement);
                const parentWidth = bar.parentElement.offsetWidth;
                const targetWidth = (bar.getAttribute('data-width') || 
                    (parseFloat(bar.parentElement.querySelector('.bar-animation').textContent) / 
                     parseFloat(bars[0].parentElement.querySelector('.bar-animation').textContent) * 100)) + '%';
                
                bar.style.width = '0%';
                setTimeout(() => {
                    bar.style.width = targetWidth;
                }, index * 150);
            });
            break;
    }
}

// ==================== YARDIMCI FONKSİYONLAR ====================
function closeChartWindow() {
    const chartWindow = document.getElementById('chartWindow');
    const overlay = document.getElementById('chartOverlay');
    
    if (chartWindow) chartWindow.remove();
    if (overlay) overlay.remove();
    
    // Ayrıca pencere içi seçiciyi de temizle
    closeChartTypeSelectorInWindow();
}

function exportChartData() {
    const allCells = document.querySelectorAll('#container input[type="text"]');
    let csv = 'Hücre,Değer,Tip\n';
    
    allCells.forEach(cell => {
        const value = cell.value.trim();
        if (value === '') return;
        
        const numValue = parseFloat(value);
        let type = 'Metin';
        
        if (!isNaN(numValue)) {
            type = 'Sayı';
        } else if (value.startsWith('=')) {
            type = 'Formül';
        }
        
        csv += `"${cell.id}","${value.replace(/"/g, '""')}","${type}"\n`;
    });
    
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `tablo-verileri-${new Date().toISOString().slice(0,10)}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // İndirme bildirimi
    const alert = document.createElement('div');
    alert.innerHTML = '✅ CSV dosyası indirildi!';
    alert.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #10b981;
        color: white;
        padding: 12px 20px;
        border-radius: 6px;
        z-index: 10001;
        font-weight: bold;
        box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        animation: slideIn 0.3s ease-out;
    `;
    
    document.body.appendChild(alert);
    setTimeout(() => {
        alert.style.animation = 'slideOut 0.3s ease-in forwards';
        setTimeout(() => alert.remove(), 300);
    }, 3000);
}

// ==================== SAYFA YÜKLENİNCE ÇALIŞTIR ====================
document.addEventListener('DOMContentLoaded', function() {
    setTimeout(addChartButtonToFormulaBar, 2000);
});

window.addEventListener('load', function() {
    setTimeout(addChartButtonToFormulaBar, 1000);
});

setInterval(function() {
    if (!document.getElementById('showChartBtn')) {
        addChartButtonToFormulaBar();
    }
}, 3000);

// Animasyon CSS'leri ekle
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }
    @keyframes slideOut {
        from { transform: translateX(0); opacity: 1; }
        to { transform: translateX(100%); opacity: 0; }
    }
`;
document.head.appendChild(style);
