// HATA TİPLERİ
const ERROR_TYPES = {
    SYNTAX: '#SYNTAX_ERROR',
    REFERENCE: '#REFERENCE_ERROR',
    DIV_ZERO: '#DIV_ZERO',
    VALUE: '#VALUE_ERROR',
    NAME: '#NAME_ERROR',
    CIRCULAR: '#CIRCULAR_REF',
    RANGE: '#RANGE_ERROR',
    CALC_TIMEOUT: '#CALC_TIMEOUT',
    CALC_INFINITE_LOOP: '#CALC_INFINITE_LOOP'
};

// Hata FFonksiyonları
const ERROR_MESSAGES = {
    [ERROR_TYPES.SYNTAX]: 'Formül syntax hatası',
    [ERROR_TYPES.REFERENCE]: 'Geçersiz hücre referansı',
    [ERROR_TYPES.DIV_ZERO]: 'Sıfıra bölme hatası',
    [ERROR_TYPES.VALUE]: 'Geçersiz değer',
    [ERROR_TYPES.NAME]: 'Tanımlanmamış fonksiyon',
    [ERROR_TYPES.CIRCULAR]: 'Döngüsel referans',
    [ERROR_TYPES.RANGE]: 'Geçersiz aralık',
    [ERROR_TYPES.CALC_TIMEOUT]: 'Hesaplama zaman aşımına uğradı',
    [ERROR_TYPES.CALC_INFINITE_LOOP]: 'Sonsuz döngü tespit edildi'
};

class ErrorHandler {
    static errorLog = [];
    static maxErrorLogSize = 100;
    
    static handleError(cell, errorType, details = '') {
        const errorMessage = ERROR_MESSAGES[errorType];
        
        // Eğer hücre yoksa (ilk yükleme sırasında)
        if (!cell) {
            console.error(`Hata: ${errorType} - ${errorMessage}`, details);
            return errorType;
        }
        
        cell.value = errorType;
        cell.title = `${errorMessage}${details ? ': ' + details : ''}`;
        cell.classList.add('error');
        
        // Log error
        this.errorLog.push({
            timestamp: new Date().toISOString(),
            cellId: cell.id || 'unknown',
            errorType,
            errorMessage,
            details
        });
        
        // Keep log size limited
        if (this.errorLog.length > this.maxErrorLogSize) {
            this.errorLog.shift();
        }
        
        console.warn(`Hata: ${errorType} - ${errorMessage}`, details);
        return errorType;
    }

    static clearError(cell) {
        if (!cell) return;
        cell.classList.remove('error');
        cell.title = '';
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
}

// Hata Validation Fonksiyonları
class FormulaValidator {
    static validateSyntax(formula) {
        try {
            // Boş formül kontrolü
            if (!formula || formula.trim() === '') {
                throw new Error('Boş formül');
            }

            // Parantez kontrolü
            const openParen = (formula.match(/\(/g) || []).length;
            const closeParen = (formula.match(/\)/g) || []).length;
            if (openParen !== closeParen) {
                throw new Error('Eşleşmeyen parantez');
            }

            // Geçersiz karakter kontrolü (daha esnek)
            const invalidChars = formula.match(/[^A-J0-9\s.,:+*\/\-()=<>!&|"']/gi);
            if (invalidChars) {
                throw new Error(`Geçersiz karakter: ${invalidChars[0]}`);
            }

            return true;
        } 
        catch (error) {
            throw new Error(ERROR_TYPES.SYNTAX);
        }
    }

    static checkCircularReference(cellId, formula, cells, visited = new Set()) {
        try {
            const referencedCells = formula.match(/[A-J][1-9][0-9]?/gi) || [];
            
            // Kendi kendine referans
            if (referencedCells.some(ref => ref.toUpperCase() === cellId)) {
                throw new Error(ERROR_TYPES.CIRCULAR);
            }
            
            // Derinlik kontrolü (max 10 seviye)
            if (visited.size > 10) {
                throw new Error(ERROR_TYPES.CIRCULAR);
            }
            
            // Referanslanan hücreleri kontrol et
            for (const refCell of referencedCells) {
                const targetCell = cells.find(cell => cell.id === refCell.toUpperCase());
                if (targetCell && targetCell.value && targetCell.value.startsWith('=')) {
                    if (visited.has(refCell.toUpperCase())) {
                        throw new Error(ERROR_TYPES.CIRCULAR);
                    }
                    
                    const newVisited = new Set([...visited, cellId]);
                    try {
                        FormulaValidator.checkCircularReference(
                            cellId, 
                            targetCell.value.slice(1), 
                            cells, 
                            newVisited
                        );
                    } catch (error) {
                        if (error.message === ERROR_TYPES.CIRCULAR) {
                            throw error;
                        }
                    }
                }
            }
            return false;
        } catch (error) {
            throw error;
        }
    }

    static validateCellReference(cellId, cells) {
        const cell = cells.find(c => c.id === cellId.toUpperCase());
        if (!cell) {
            throw new Error(ERROR_TYPES.REFERENCE);
        }
        return cell;
    }

    static validateNumericOperations(value){
        if(typeof value ==='number' && !isFinite(value)){
            throw new Error(ERROR_TYPES.VALUE);
        }
        return true;
    }

    static validateFunctionArguments(args, expectedCount, functionName){
        if(args.length !== expectedCount){
            throw new Error(`${ERROR_TYPES.VALUE}: ${functionName} fonksiyonu ${expectedCount} argüman bekliyor`);
        }
        return true;
    }
    
    // Input sanitization
    static sanitizeInput(value) {
        if (typeof value !== 'string') return value;
        
        // Potansiyel tehlikeli karakterleri temizle (daha az agresif)
        const dangerousChars = /[<>\\]/g;
        const sanitized = value.replace(dangerousChars, '');
        
        // Uzunluğu sınırla
        return sanitized.length > 1000 ? sanitized.substring(0, 1000) : sanitized;
    }
    
    // Cell ID validation
    static validateCellId(id) {
        const regex = /^[A-J]([1-9]|[1-9][0-9])$/i;
        return regex.test(id);
    }
}