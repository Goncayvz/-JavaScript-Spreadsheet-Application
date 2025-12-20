//çoklu seçim fonksiyonu değişkenleri
let isRangeSelecting = false;
let selectionStartCell = null;
let selectedCells = [];
let selectedCell = null;
let mergedCells = new Map();
let isEditing = false;
let selectedCellForFormula = null;


// ============ YENİ DEĞİŞKENLER ============
let calculationQueue = new Map();
let isCalculating = false;
let undoStack = [];
let redoStack = [];
let maxStackSize = 50;
let clipboard = null;
let cursorTrackerCleanup = null;
// ==================== EXCEL FORMÜL SİSTEMİ - BASİT VE ÇALIŞAN ====================

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
    if (!cell) return 0;
    
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

// ==================== SPREADSHEET OLUŞTURMA ====================

window.onload = () => {
    const container = document.getElementById("container");
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
    range(1, 99).forEach(number => {
        // Satır numarası
        createLabel(number);

        // Hücreler
        letters.forEach(letter => {
            const input = document.createElement("input");
            input.type = "text";
            input.id = letter + number;
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
    }, 100);
};

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
    const currentCellDisplay = document.getElementById('currentCellDisplay');
    
    if (formulaInput) {
        // Formül bar'da ORİJİNAL formülü göster
        formulaInput.value = cell.dataset.originalValue || cell.value || '';
    }
    
    if (currentCellDisplay) {
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
        const cellDisplay = document.createElement('div');
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

