// parsers.js

import { FULL_CONFIG } from './config.js';
import { findSheet, extractFundName, findHeaderRowIndex, getHeaderMapping } from './utils.js';

// ★★★ 核心修改：專門為現金流量表添加後綴的輔助函式 ★★★
function applyCashFlowSuffixes(records, keyColumn) {
    const itemCounter = {
        '收取利息': 0,
        '收取股利': 0,
        '支付利息': 0
    };
    const targetItems = Object.keys(itemCounter);
    // 使用正規表示式移除括號內容和空白，以進行準確匹配
    const normalize = (name) => String(name || '').replace(/\s|　/g, '').replace(/（.*）|\(.*\)/, '');

    return records.map(record => {
        const keyText = record[keyColumn] || '';
        const normalizedKey = normalize(keyText);

        if (targetItems.includes(normalizedKey)) {
            itemCounter[normalizedKey]++;
            const count = itemCounter[normalizedKey];
            let suffix = '';

            if (count === 2) { // 只處理第二次出現的項目
                if (normalizedKey === '收取利息' || normalizedKey === '收取股利') {
                    suffix = ' (投資活動)';
                } else if (normalizedKey === '支付利息') {
                    suffix = ' (籌資活動)';
                }
            }

            if (suffix) {
                // 回傳一個新物件，避免直接修改原始 record
                return { ...record, [keyColumn]: keyText + suffix };
            }
        }
        return record;
    });
}


function getIndentLevelFromWhitespace(rawText) {
    const text = String(rawText || '');
    const match = text.match(/^(\s|　)+/); 
    if (!match) {
        return 0; 
    }
    const leadingWhitespace = match[0];
    let visualWidth = 0;
    for (const char of leadingWhitespace) {
        visualWidth += (char === '　' ? 2 : 1);
    }
    const indentUnit = 2; 
    const indentLevel = Math.floor(visualWidth / indentUnit);
    return indentLevel;
}

function _parseFixed(data, config, fundName, sheet, startRow, colMap) {
    const records = [];
    for (let i = startRow; i < data.length; i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const keyColIndex = colMap[config.keyColumn];
        if (keyColIndex === undefined) continue;

        const keyTextRaw = String(row[keyColIndex] || '');
        const keyText = keyTextRaw.trim();
        if (!keyText || keyText.startsWith('註') || keyText.startsWith('附註')) continue;
        
        const record = { '基金名稱': fundName };
        let hasMeaningfulData = false;
        
        record.indent_level = getIndentLevelFromWhitespace(keyTextRaw);

        config.columns.forEach(colName => {
            const colIndex = colMap[colName];
            const value = (colIndex !== undefined) ? row[colIndex] : '';
            record[colName] = colName === config.keyColumn ? keyText : value; 
            if (colName !== config.keyColumn && value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
                hasMeaningfulData = true;
            }
        });

        if (hasMeaningfulData || keyText) {
            records.push(record);
        }
    }
    return records;
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
        
        const keyTextRaw = String(row[keyColIndex] || '');
        const keyText = keyTextRaw.trim();
        if (!keyText || keyText.startsWith('註')) continue;

        const record = { '基金名稱': fundName };
        let hasMeaningfulData = false;
        
        record.indent_level = getIndentLevelFromWhitespace(keyTextRaw);
        
        for (const colName in headerMapping) {
            const value = row[headerMapping[colName]];
            record[colName] = colName === config.keyColumn ? keyText : value;
            if (colName !== config.keyColumn && value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) hasMeaningfulData = true;
        }

        if (hasMeaningfulData || keyText) records.push(record);
    }
    return records;
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

        const keyTextRaw = String(row[keyColIndex] || '');
        const keyText = keyTextRaw.trim();
        if (!keyText || keyText.startsWith('註')) continue;

        const record = { '基金名稱': fundName };
        let hasMeaningfulData = false;

        record.indent_level = getIndentLevelFromWhitespace(keyTextRaw);
        
        for (const colName in headerMapping) {
            const value = row[headerMapping[colName]];
            record[colName] = colName === config.keyColumn ? keyText : value;
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

function parseBusinessProfitLoss_Stateful(data, config, fundName, sheet) {
    const records = [];
    const colMap = { '科目': 2, '上年度決算數': 0, '本年度預算數': 3, '原列決算數': 5, '修正數': 6, '決算核定數': 7 };
    const keyColIndex = colMap[config.keyColumn];
    const numericCols = config.columns.filter(c => c !== config.keyColumn);
    const normalize = (name) => String(name || '').replace(/\s|　/g, '').split('(')[0];
    let inNonOperatingSection = false;
    for (let i = 6; i < data.length; i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const keyTextRaw = String(row[keyColIndex] || '');
        let keyText = keyTextRaw.trim();
        if (!keyText || keyText.startsWith('註')) continue;
        if (normalize(keyText) === '營業外收入') { inNonOperatingSection = true; }
        const normalizedKey = normalize(keyText);
        if (normalizedKey === '採用權益法認列之關聯企業及合資利益之份額' || normalizedKey === '採用權益法認列之關聯企業及合資損失之份額') {
            if (inNonOperatingSection) { keyText += ' (營業外)'; }
        }
        const record = { '基金名稱': fundName, [config.keyColumn]: keyText };
        let hasMeaningfulData = false;
        numericCols.forEach(colName => {
            const colIndex = colMap[colName];
            const value = (colIndex !== undefined) ? row[colIndex] : '';
            record[colName] = value;
            if (value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
                hasMeaningfulData = true;
            }
        });
        record.indent_level = getIndentLevelFromWhitespace(keyTextRaw);
        if (hasMeaningfulData || keyText) { records.push(record); }
    }
    return records;
}

function parseBusinessAppropriation_Stateful(data, config, fundName, sheet) {
    const records = [];
    const colMap = { '項目': 2, '上年度決算數': 0, '本年度預算數': 3, '原列決算數': 5, '修正數': 6, '決算核定數': 7 };
    const keyColIndex = colMap[config.keyColumn];
    const numericCols = config.columns.filter(c => c !== config.keyColumn);
    const normalize = (name) => String(name || '').replace(/\s|　/g, '').split('(')[0];
    let inDeficitSection = false; 
    for (let i = 6; i < data.length; i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const keyTextRaw = String(row[keyColIndex] || '');
        let keyText = keyTextRaw.trim();
        if (!keyText || keyText.startsWith('註')) continue;
        if (normalize(keyText) === '虧損之部') { inDeficitSection = true; }
        if (normalize(keyText) === '其他綜合損益轉入數') {
            keyText += inDeficitSection ? ' (虧損之部)' : ' (盈餘之部)';
        }
        const record = { '基金名稱': fundName, [config.keyColumn]: keyText };
        let hasMeaningfulData = false;
        numericCols.forEach(colName => {
            const colIndex = colMap[colName];
            const value = (colIndex !== undefined) ? row[colIndex] : '';
            record[colName] = value;
            if (value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
                hasMeaningfulData = true;
            }
        });
        record.indent_level = getIndentLevelFromWhitespace(keyTextRaw);
        if (hasMeaningfulData || keyText) { records.push(record); }
    }
    return records;
}

function parseFixedBusinessCashFlow(data, config, fundName, sheet) {
    const records = [];
    const colMap = { '項目': 0, '本年度預算數': 1, '原列決算數': 2, '修正數': 3, '決算核定數': 4 };
    const keyColIndex = colMap[config.keyColumn];
    const startRow = 5;
    for (let i = startRow; i < data.length; i++) {
        const row = data[i];
        if (!Array.isArray(row) || row.length === 0) continue;
        const keyTextRaw = String(row[keyColIndex] || '');
        const keyText = keyTextRaw.trim();
        if (!keyText || keyText.startsWith('註') || keyText.startsWith('附註')) continue;
        const record = { '基金名稱': fundName };
        let hasMeaningfulData = false;
        config.columns.forEach(colName => {
            const colIndex = colMap[colName];
            const value = colIndex !== undefined ? row[colIndex] : '';
            record[colName] = colName === config.keyColumn ? keyText : value;
            if (colName !== config.keyColumn && value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
                hasMeaningfulData = true;
            }
        });
        record.indent_level = getIndentLevelFromWhitespace(keyTextRaw);
        if (hasMeaningfulData || keyText) { records.push(record); }
        if (keyText.includes('期末現金及約當現金')) { break; }
    }
    return records;
}

function parseBusinessBalanceSheet_SideBySide(data, config, fundName, sheet) {
    const assetConfig = { ...config };
    const assetColMap = { '科目': 3, '上年度決算數': 1, '原列決算數': 4, '修正數': 5, '決算核定數': 6 };
    const assetRecords = _parseFixed(data, assetConfig, fundName, sheet, 6, assetColMap);
    const liabilityConfig = { ...config };
    const liabilityColMap = { '科目': 10, '上年度決算數': 8, '原列決算數': 11, '修正數': 12, '決算核定數': 13 };
    const liabilityRecords = _parseFixed(data, liabilityConfig, fundName, sheet, 6, liabilityColMap);
    return { '資產負債表_資產': assetRecords, '資產負債表_負債及權益': liabilityRecords };
}

const PARSERS = {
    dynamic_normal: parseNormalTable,
    balance_sheet: parseBalanceSheet,
    fixed_yuchu: parseFixedYuchuBiao,
    fixed_xianliu: parseFixedXianliuBiao,
    fixed_shouzhi: parseFixedShouzhiBiao,
    fixed_business_profitloss_stateful: parseBusinessProfitLoss_Stateful,
    fixed_business_appropriation_stateful: parseBusinessAppropriation_Stateful,
    fixed_business_cashflow: parseFixedBusinessCashFlow,
    fixed_business_balancesheet_sidebyside: parseBusinessBalanceSheet_SideBySide,
};

export function processFile(file, selectedFundType) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'array' });
                let fundName = extractFundName(workbook, selectedFundType) || file.name.replace(/\.xlsx?$/, '');
                let extractedData = {};
                const activeConfig = FULL_CONFIG[selectedFundType];
                for (const reportKey in activeConfig) {
                    const config = activeConfig[reportKey];
                    const sheetName = findSheet(workbook, config.sheetKeyword);
                    if (!sheetName) continue;
                    const sheet = workbook.Sheets[sheetName];
                    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                    const parserFunc = PARSERS[config.parser];
                    if (parserFunc) {
                        let records = parserFunc(data, config, fundName, sheet);
                        
                        // ★★★ 核心修改：在此處應用後綴 ★★★
                        if (reportKey === '現金流量表' && selectedFundType === 'business' && records && records.length > 0) {
                            records = applyCashFlowSuffixes(records, config.keyColumn);
                        }

                        if (records) {
                            if (reportKey === '平衡表' || reportKey === '資產負債表') {
                                for (const key in records) {
                                    if (records[key].length > 0) {
                                        if (!extractedData[key]) extractedData[key] = [];
                                        extractedData[key].push(...records[key]);
                                    }
                                }
                            } else if (records.length > 0) {
                                if (!extractedData[reportKey]) extractedData[reportKey] = [];
                                extractedData[reportKey].push(...records);
                            }
                        }
                    }
                }
                const fileHasData = Object.keys(extractedData).length > 0;
                resolve(fileHasData ? { fundName, fileName: file.name, data: extractedData } : null);
            } catch (error) {
                console.error("Error processing file:", file.name, error);
                reject(error);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}
