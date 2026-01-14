// HATA TİPLERİ - GLOBAL
const ERROR_TYPES = {
    SYNTAX: '#SYNTAX',
    REFERENCE: '#REFERENCE',
    DIV_ZERO: '#DIV_ZERO',
    VALUE: '#VALUE',
    NAME: '#NAME',
    CIRCULAR: '#CIRCULAR',
    CALC_TIMEOUT: '#CALC_TIMEOUT',
    CALC_INFINITE_LOOP: '#CALC_INFINITE_LOOP'
};

// Hata Mesajları - GLOBAL
const ERROR_MESSAGES = {
    [ERROR_TYPES.SYNTAX]: 'Formül syntax hatası',
    [ERROR_TYPES.REFERENCE]: 'Geçersiz hücre referansı',
    [ERROR_TYPES.DIV_ZERO]: 'Sıfıra bölme hatası',
    [ERROR_TYPES.VALUE]: 'Geçersiz değer',
    [ERROR_TYPES.NAME]: 'Tanımlanmamış fonksiyon',
    [ERROR_TYPES.CIRCULAR]: 'Döngüsel referans',
    [ERROR_TYPES.CALC_TIMEOUT]: 'Hesaplama zaman aşımına uğradı',
    [ERROR_TYPES.CALC_INFINITE_LOOP]: 'Sonsuz döngü tespit edildi'
};

// ErrorHandler sınıfı - GLOBAL
class ErrorHandler {
    static errorLog = [];
    static maxErrorLogSize = 100;
    
    static handleError(cell, errorType, details = '') {
        const errorMessage = ERROR_MESSAGES[errorType] || 'Bilinmeyen hata';
        
        // Eğer hücre yoksa
        if (!cell) {
            console.error(`Hata: ${errorType} - ${errorMessage}`, details);
            return errorType;
        }
        
        // Hata değerini hücreye yaz
        cell.value = errorType;
        cell.title = `${errorMessage}${details ? ': ' + details : ''}`;
        cell.classList.add('error');
        
        // Eğer formül hatası ise formül sınıfını kaldır
        cell.classList.remove('formula');
        
        // Hesaplanmış değeri sil
        delete cell.calculatedValue;
        
        // Orijinal değeri kaydet
        if (!cell.dataset.originalValue) {
            cell.dataset.originalValue = cell.value;
        }
        
        // Log'a ekle
        this.errorLog.push({
            timestamp: new Date().toISOString(),
            cellId: cell.id || 'unknown',
            errorType: errorType,
            errorMessage: errorMessage,
            details: details,
            originalValue: cell.dataset.originalValue
        });
        
        // Log boyutunu kontrol et
        if (this.errorLog.length > this.maxErrorLogSize) {
            this.errorLog.shift();
        }
        
        console.warn(`Hata: ${cell.id} - ${errorType} - ${errorMessage}`, details);
        return errorType;
    }

    static clearError(cell) {
        if (!cell) return;
        cell.classList.remove('error');
        cell.title = '';
        
        // Eğer orijinal değer varsa geri yükle
        if (cell.dataset.originalValue && cell.dataset.originalValue.startsWith('=')) {
            cell.value = cell.dataset.originalValue;
            cell.classList.add('formula');
            delete cell.dataset.originalValue;
        }
    }

    static isValidError(value) {
        if (!value || typeof value !== 'string') return false;
        return Object.values(ERROR_TYPES).includes(value);
    }
    
    static getErrorLog() {
        return [...this.errorLog];
    }
    
    static clearErrorLog() {
        this.errorLog = [];
    }
    
    static getErrorCount() {
        return this.errorLog.length;
    }
}

// FormulaValidator sınıfı - GLOBAL
class FormulaValidator {
    static sanitizeInput(value) {
        if (value === null || value === undefined) return '';
        if (typeof value !== 'string') return String(value);
        
        // Potansiyel tehlikeli karakterleri temizle
        const dangerousChars = /[<>\\]/g;
        let sanitized = value.replace(dangerousChars, '');
        
        // Baştaki ve sondaki boşlukları temizle
        sanitized = sanitized.trim();
        
        // Uzunluğu sınırla
        return sanitized.length > 1000 ? sanitized.substring(0, 1000) : sanitized;
    }
    
    static validateCellId(id) {
        if (!id || typeof id !== 'string') return false;
        const regex = /^[A-J]([1-9]|[1-9][0-9]|100)$/i;
        return regex.test(id);
    }
    
    static validateFormulaSyntax(formula) {
        if (!formula || !formula.startsWith('=')) {
            return { valid: true };
        }
        
        const expr = formula.substring(1).trim();
        
        // Boş formül kontrolü
        if (!expr) {
            return { valid: false, error: ERROR_TYPES.SYNTAX, message: 'Boş formül' };
        }
        
        // Parantez kontrolü
        const openParen = (expr.match(/\(/g) || []).length;
        const closeParen = (expr.match(/\)/g) || []).length;
        if (openParen !== closeParen) {
            return { valid: false, error: ERROR_TYPES.SYNTAX, message: 'Eşleşmeyen parantez' };
        }
        
        return { valid: true };
    }
}

// Global değişkenleri window nesnesine ekleyelim
if (typeof window !== 'undefined') {
    window.ERROR_TYPES = ERROR_TYPES;
    window.ErrorHandler = ErrorHandler;
    window.FormulaValidator = FormulaValidator;
    console.log('✅ ErrorHandler global olarak yüklendi!');
}
