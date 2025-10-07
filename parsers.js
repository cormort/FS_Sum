// parsers.js: 包含所有解析 Excel 檔案的函式

import { FULL_CONFIG } from './config.js';

// --- Main Exported Function ---
export function processFile(file, selectedFundType) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const workbook = XLSX.read(e.target.result, { type: 'array', cellStyles: true });
            let fundName = extractFundName(workbook) || file.name.replace(/\.xlsx?$/, '');
            let extractedData = {};
            const activeConfig = FULL_CONFIG[selectedFundType];

            for (const reportKey in activeConfig) {
                const config = activeConfig[reportKey];
                const sheetName = findSheet(workbook, config.sheetKeyword);
                if (!sheetName) continue;

                const sheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                let records;

                switch (config.parser) {
                    case 'dynamic_normal':
                        records = parseNormalTable(data, config, fundName, sheet);
                        if (records.length > 0) extractedData[reportKey] = records;
                        break;
                    case 'balance_sheet':
                        records = parseBalanceSheet(data, fundName, sheet);
                        for (const key in records) {
                            if (records[key].length > 0) {
                                if (!extractedData[key]) extractedData[key] = [];
                                extractedData[key].push(...records[key]);
                            }
                        }
                        break;
                    case 'fixed_yuchu':
                        records = parseFixedYuchuBiao(data, config, fundName, sheet);
                        if (records.length > 0) extractedData[reportKey] = records;
                        break;
                    case 'fixed_xianliu':
                        records = parseFixedXianliuBiao(data, config, fundName, sheet);
                        if (records.length > 0) extractedData[reportKey] = records;
                        break;
                    case 'fixed_shouzhi':
                        records = parseFixedShouzhiBiao(data, config, fundName, sheet);
                        if (records.length > 0) extractedData[reportKey] = records;
                        break;
                    case 'fixed_business_profitloss':
                        records = parseFixedBusinessProfitLoss(data, config, fundName, sheet);
                        if (records.length > 0) extractedData[reportKey] = records;
                        break;
                    case 'fixed_business_appropriation':
                        records = parseFixedBusinessAppropriation(data, config, fundName, sheet);
                        if (records.length > 0) extractedData[reportKey] = records;
                        break;
                    case 'fixed_business_cashflow':
                        records = parseFixedBusinessCashFlow(data, config, fundName, sheet);
                        if (records.length > 0) extractedData[reportKey] = records;
                        break;
                    case 'fixed_business_balancesheet':
                        records = parseFixedBusinessBalanceSheet(data, fundName, sheet);
                        for (const key in records) {
                            if (records[key].length > 0) {
                                if (!extractedData[key]) extractedData[key] = [];
                                extractedData[key].push(...records[key]);
                            }
                        }
                        break;
                }
            }
            const fileHasData = Object.keys(extractedData).length > 0;
            resolve(fileHasData ? { fundName, fileName: file.name, data: extractedData } : null);
        };
        reader.readAsArrayBuffer(file);
    });
}


// --- Helper and Dynamic Parser Functions (Internal to this module) ---
function extractFundName(workbook) {
    let longestName = "";
    for(let i=0; i < Math.min(workbook.SheetNames.length, 3); i++) {
        const sheetName = workbook.SheetNames[i];
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        for(let j=0; j < Math.min(data.length, 5); j++) {
            const row = data[j];
            if(Array.isArray(row)) {
                const fullRowText = row.join('');
                if (fullRowText.includes('基金')) {
                    for(const cell of row) {
                        const cellText = String(cell).trim();
                        if (cellText.includes('基金') && cellText.length > longestName.length) longestName = cellText;
                    }
                }
            }
        }
    }
    if (longestName) {
        const removableParts = ['收支餘絀表', '餘絀撥補表', '現金流量表', '平衡表', '資產負債表', '損益表', '盈虧撥補表']; 
        removableParts.forEach(part => { longestName = longestName.replace(part, ''); });
        return longestName.trim();
    }
    return null;
}
function findSheet(workbook, keyword) { return workbook.SheetNames.find(name => name.includes(keyword)); }
function findHeaderRowIndex(data, columns) {
    let bestMatch = { rowIndex: -1, score: 0 };
    for (let i = 0; i < Math.min(data.length, 10); i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const rowAsString = row.join(' ').replace(/\s/g, '');
        let score = 0;
        columns.forEach(col => {
            const cleanCol = col.replace(/\s/g, '');
            if (rowAsString.includes(cleanCol)) score++;
        });
        if (score > bestMatch.score) bestMatch = { rowIndex: i, score };
    }
    return bestMatch.score > 1 ? bestMatch.rowIndex : -1;
}
function getHeaderMapping(headerRow, columns, startCol = 0) {
    const mapping = {};
    const assignedCols = new Set();
    columns.forEach(colName => {
        let bestMatchColIndex = -1;
        const cleanColName = colName.replace(/\s/g, '');
        for (let i = startCol; i < headerRow.length; i++) {
            if (assignedCols.has(i)) continue;
            const cellContent = String(headerRow[i] || '').trim().replace(/\s/g, '');
            if (cellContent === cleanColName) { bestMatchColIndex = i; break; }
            if (cellContent.includes(cleanColName) && bestMatchColIndex === -1) { bestMatchColIndex = i; }
        }
        if (bestMatchColIndex !== -1) { mapping[colName] = bestMatchColIndex; assignedCols.add(bestMatchColIndex); }
    });
    return mapping;
}
function parseNormalTable(data, config, fundName, sheet) {
    const headerRowIndex = findHeaderRowIndex(data, config.columns);
    if (headerRowIndex === -1) return [];
    const headerMapping = getHeaderMapping(data[headerRowIndex], config.columns);
    const keyColIndex = headerMapping[config.keyColumn];
    if (keyColIndex === undefined) return [];
    const records = [];
    for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const keyText = String(row[keyColIndex] || '').trim();
        if (!keyText || keyText.startsWith('註')) continue;
        const record = { '基金名稱': fundName };
        let hasMeaningfulData = false;
        let indentLevel = 0;
        try {
            const cellAddress = XLSX.utils.encode_cell({ r: i, c: keyColIndex });
            if (sheet[cellAddress]?.s?.alignment?.indent) indentLevel = sheet[cellAddress].s.alignment.indent;
        } catch (e) {}
        record.indent_level = indentLevel;
        for (const colName in headerMapping) {
            const value = row[headerMapping[colName]];
            record[colName] = value;
            if (colName !== config.keyColumn && value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) hasMeaningfulData = true;
        }
        if (hasMeaningfulData || keyText) records.push(record);
    }
    return records;
}

function parseBalanceSheet(data, fundName, sheet) {
    const assetConfig = { keyColumn: '科目', columns: ['科目', '本年度決算核定數', '上年度決算審定數', '比較增減'], subTableIdentifier: '資產' };
    const liabilityConfig = { keyColumn: '科目', columns: ['科目', '本年度決算核定數', '上年度決算審定數', '比較增減'], subTableIdentifier: '負債' };
    
    const assetRecords = _parseSideBySide(data, assetConfig, fundName, sheet);
    const liabilityRecords = _parseSideBySide(data, liabilityConfig, fundName, sheet);

    return { '平衡表_資產': assetRecords, '平衡表_負債及權益': liabilityRecords };
}

function _parseSideBySide(data, config, fundName, sheet) {
    const headerRowIndex = findHeaderRowIndex(data, config.columns);
    if (headerRowIndex === -1) return [];
    let tableStartCol = 0;
    if (config.subTableIdentifier === '負債') {
        let foundFirst = false;
        for (let i = 0; i < data[headerRowIndex].length; i++) {
            if (String(data[headerRowIndex][i]).replace(/\s|　/g, '').includes(config.keyColumn)) {
                if (foundFirst) { tableStartCol = i; break; }
                foundFirst = true;
            }
        }
        if (tableStartCol === 0 && data[headerRowIndex].filter(h => String(h).replace(/\s|　/g, '').includes(config.keyColumn)).length > 1) return [];
    }
    const headerMapping = getHeaderMapping(data[headerRowIndex], config.columns, tableStartCol);
    const keyColIndex = headerMapping[config.keyColumn];
    if (keyColIndex === undefined) return [];
    const records = [];
    let dataStartRowIndex = data.findIndex(row => Array.isArray(row) && String(row[tableStartCol] || '').includes(config.subTableIdentifier));
    if (dataStartRowIndex === -1) dataStartRowIndex = headerRowIndex + 1;
    for (let i = dataStartRowIndex; i < data.length; i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const keyText = String(row[keyColIndex] || '').trim();
        if (!keyText || keyText.startsWith('註')) continue;
        const record = { '基金名稱': fundName };
        let hasMeaningfulData = false;
        let indentLevel = 0;
        try {
            const cellAddress = XLSX.utils.encode_cell({ r: i, c: keyColIndex });
            if (sheet[cellAddress]?.s?.alignment?.indent) indentLevel = sheet[cellAddress].s.alignment.indent;
        } catch (e) {}
        record.indent_level = indentLevel;
        for (const colName in headerMapping) {
            const value = row[headerMapping[colName]];
            record[colName] = value;
            if (colName !== config.keyColumn && value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) hasMeaningfulData = true;
        }
        if (hasMeaningfulData || keyText) records.push(record);
    }
    return records;
}

// --- Fixed-Layout Parsers ---
function _parseFixed(data, config, fundName, sheet, startRow, colMap) {
    const records = [];
    for (let i = startRow; i < data.length; i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const keyColIndex = colMap[config.keyColumn];
        if (keyColIndex === undefined) continue;

        const keyText = String(row[keyColIndex] || '').trim();
        if (!keyText || keyText.startsWith('註') || keyText.startsWith('附註')) continue;
        const record = { '基金名稱': fundName };
        let hasMeaningfulData = false;
        let indentLevel = 0;
        try {
            const cellAddress = XLSX.utils.encode_cell({ r: i, c: keyColIndex });
            if (sheet[cellAddress]?.s?.alignment?.indent) indentLevel = sheet[cellAddress].s.alignment.indent;
        } catch(e) {}
        record.indent_level = indentLevel;

        config.columns.forEach(colName => {
             const colIndex = colMap[colName];
             if (colIndex !== undefined) {
                const value = row[colIndex];
                record[colName] = value;
                if (colName !== config.keyColumn && value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
                    hasMeaningfulData = true;
                }
             } else {
                record[colName] = '';
             }
        });
        if (hasMeaningfulData || keyText) {
            records.push(record);
        }
    }
    return records;
}

function parseFixedYuchuBiao(data, config, fundName, sheet) {
    const colMap = { '項目': 0, '預算數': 3, '原列決算數': 5, '修正數': 8, '決算核定數': 9, '預算與決算核定數比較增減': 11 };
    return _parseFixed(data, config, fundName, sheet, 4, colMap);
}
function parseFixedXianliuBiao(data, config, fundName, sheet) {
    const colMap = { '項目': 0, '決算核定數': 3 };
    return _parseFixed(data, config, fundName, sheet, 5, colMap);
}
function parseFixedShouzhiBiao(data, config, fundName, sheet) {
    const colMap = { '科目': 0, '原列決算數': 3, '修正數': 4, '決算核定數': 5 };
    return _parseFixed(data, config, fundName, sheet, 4, colMap);
}

function parseFixedBusinessProfitLoss(data, config, fundName, sheet) {
    const colMap = { '科目': 2, '上年度決算數': 0, '本年度預算數': 3, '原列決算數': 5, '修正數': 6, '決算核定數': 7 };
    return _parseFixed(data, config, fundName, sheet, 6, colMap);
}
function parseFixedBusinessAppropriation(data, config, fundName, sheet) {
    const colMap = { '項目': 2, '上年度決算數': 0, '本年度預算數': 3, '原列決算數': 5, '修正數': 6, '決算核定數': 7 };
    return _parseFixed(data, config, fundName, sheet, 6, colMap);
}
function parseFixedBusinessCashFlow(data, config, fundName, sheet) {
    const colMap = { '項目': 0, '本年度預算數': 1, '原列決算數': 2, '修正數': 3, '決算核定數': 4 };
    return _parseFixed(data, config, fundName, sheet, 5, colMap);
}
function parseFixedBusinessBalanceSheet(data, fundName, sheet) {
    const assetConfig = { keyColumn: '科目', columns: ['科目', '上年度決算數', '原列決算數', '修正數', '決算核定數'] };
    const assetColMap = { '科目': 3, '上年度決算數': 1, '原列決算數': 4, '修正數': 5, '決算核定數': 6 };
    const assetRecords = _parseFixed(data, assetConfig, fundName, sheet, 6, assetColMap);

    const liabilityConfig = { keyColumn: '科目', columns: ['科目', '上年度決算數', '原列決算數', '修正數', '決算核定數'] };
    const liabilityColMap = { '科目': 10, '上年度決算數': 8, '原列決算數': 11, '修正數': 12, '決算核定數': 13 };
    const liabilityRecords = _parseFixed(data, liabilityConfig, fundName, sheet, 6, liabilityColMap);

    return { '資產負債表_資產': assetRecords, '資產負債表_負債及權益': liabilityRecords };
}