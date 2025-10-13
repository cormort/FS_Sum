// parsers.js

import { FULL_CONFIG } from './config.js';
import { findSheet, extractFundName, findHeaderRowIndex, getHeaderMapping } from './utils.js';

// --- 通用解析器 ---
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

// --- 作業基金解析器 ---
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
function parseBalanceSheet(data, fundName, sheet) {
    const assetConfig = { keyColumn: '科目', columns: ['科目', '本年度決算核定數', '上年度決算審定數', '比較增減'], subTableIdentifier: '資產' };
    const liabilityConfig = { keyColumn: '科目', columns: ['科目', '本年度決算核定數', '上年度決算審定數', '比較增減'], subTableIdentifier: '負債' };
    
    const assetRecords = _parseSideBySide(data, assetConfig, fundName, sheet);
    const liabilityRecords = _parseSideBySide(data, liabilityConfig, fundName, sheet);

    return { '平衡表_資產': assetRecords, '平衡表_負債及權益': liabilityRecords };
}

// --- 政事基金解析器 ---
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

// --- 營業基金解析器 ---
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

        let keyText = String(row[keyColIndex] || '').trim();
        if (!keyText || keyText.startsWith('註')) continue;

        if (normalize(keyText) === '營業外收入') {
            inNonOperatingSection = true;
        }

        const normalizedKey = normalize(keyText);
        if (
            normalizedKey === '採用權益法認列之關聯企業及合資利益之份額' ||
            normalizedKey === '採用權益法認列之關聯企業及合資損失之份額'
        ) {
            if (inNonOperatingSection) {
                keyText += ' (營業外)';
            }
        }

        const record = { '基金名稱': fundName, [config.keyColumn]: keyText };
        let hasMeaningfulData = false;

        numericCols.forEach(colName => {
            const colIndex = colMap[colName];
            if (colIndex !== undefined) {
                const value = row[colIndex];
                record[colName] = value;
                if (value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
                    hasMeaningfulData = true;
                }
            }
        });
        
        let indentLevel = 0;
        try {
            const cellAddress = XLSX.utils.encode_cell({ r: i, c: keyColIndex });
            if (sheet[cellAddress]?.s?.alignment?.indent) {
                indentLevel = sheet[cellAddress].s.alignment.indent;
            }
        } catch(e) {}
        record.indent_level = indentLevel;

        if (hasMeaningfulData || keyText) {
            records.push(record);
        }
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

        let keyText = String(row[keyColIndex] || '').trim();
        if (!keyText || keyText.startsWith('註')) continue;

        if (normalize(keyText) === '虧損之部') {
            inDeficitSection = true;
        }

        if (normalize(keyText) === '其他綜合損益轉入數') {
            keyText += inDeficitSection ? ' (虧損之部)' : ' (盈餘之部)';
        }

        const record = { '基金名稱': fundName, [config.keyColumn]: keyText };
        let hasMeaningfulData = false;

        numericCols.forEach(colName => {
            const colIndex = colMap[colName];
            if (colIndex !== undefined) {
                const value = row[colIndex];
                record[colName] = value;
                if (value != null && value !== '' && !isNaN(parseFloat(String(value).replace(/,/g, '')))) {
            _B Financial Statements in Machine-Readable Format
                                                                                               
### XBRL (eXtensible Business Reporting Language)

XBRL is the global standard for exchanging business information. It provides a way to "tag" financial data so it can be processed automatically by computers.

* **How it works:** Each piece of financial data (like "Revenue" or "Net Income") is given a unique, computer-readable tag from a standard dictionary, or "taxonomy." This eliminates ambiguity and makes data comparison and analysis much easier.
* **Adoption:** Public companies in many countries, including the U.S. (mandated by the SEC), are required to file their financial statements in XBRL format.
* **Benefits:**
    * **Automation:** Greatly simplifies data extraction and analysis.
    * **Accuracy:** Reduces errors from manual data entry.
    * **Comparability:** Easy to compare financials across different companies, industries, and time periods.

### APIs (Application Programming Interfaces)

Many financial data providers offer APIs that allow developers to programmatically access their vast databases of financial information.

* **How it works:** You write code to send a request to the provider's server (e.g., "get Apple Inc.'s quarterly revenue for the last 5 years"). The server then sends back the requested data in a structured format like JSON or XML.
* **Providers:**
    * **Bloomberg:** A top choice for institutional investors, offering real-time and historical data.
    * **Refinitiv (formerly Thomson Reuters):** Another major player with comprehensive financial data.
    * **FactSet:** Provides integrated financial information and analytical applications.
    * **Alpha Vantage, IEX Cloud:** More accessible and often cheaper options, popular with individual investors and startups.
* **Benefits:**
    * **Real-time data:** Many APIs provide up-to-the-minute information.
    * **Flexibility:** You can request the exact data points you need.
    * **Integration:** Easy to integrate financial data into your own applications, models, and websites.

### Financial Data Aggregators

These are services that collect, clean, and standardize financial data from various sources (including SEC filings and press releases) and make it available in a more user-friendly format, often through web platforms or their own APIs.

* **Examples:**
    * **S&P Capital IQ:** A web-based platform providing deep data and analytics on public and private companies.
    * **PitchBook:** Focuses on private market data, including venture capital, private equity, and M&A.
* **Benefits:**
    * **Cleaned data:** Aggregators do the hard work of standardizing data from different sources.
    * **Value-added information:** Often include analyst ratings, estimates, and other proprietary data.
    * **User-friendly interfaces:** Make it easy to search, screen, and analyze data without needing to code.

### Direct from Company Websites

Most publicly traded companies have an "Investor Relations" section on their website where they post their quarterly and annual reports, SEC filings, and press releases.

* **Formats:** Usually available as PDFs, and sometimes as interactive web pages or spreadsheets (Excel).
* **Benefits:**
    * **Source documents:** You get the information directly from the source.
    * **Free:** This information is publicly available at no cost.
* **Drawbacks:**
    * **Not machine-readable:** Data is often "locked" in PDFs, requiring manual extraction.
    * **Inconsistent formats:** Each company may present its data differently.

### Summary: Choosing the Right Format

| Format/Source           | Best For                                                              | Pros                                                     | Cons                                                      |
| ----------------------- | --------------------------------------------------------------------- | -------------------------------------------------------- | --------------------------------------------------------- |
| **XBRL** | Large-scale, automated analysis of public company filings.            | Standardized, accurate, highly comparable.               | Can have a learning curve; requires specific tools to parse. |
| **APIs** | Integrating live or historical data into custom applications.         | Flexible, real-time, broad data coverage.                | Can be expensive; requires programming skills.             |
| **Financial Aggregators** | In-depth research and analysis without programming.                   | Cleaned data, user-friendly, value-added analytics.      | Often subscription-based and can be costly.               |
| **Company Websites (PDF)** | Verifying specific figures or understanding the company's narrative. | Free, direct from the source, includes management commentary. | Not machine-readable, manual effort, inconsistent.        |

The trend is overwhelmingly towards more structured, machine-readable formats like XBRL and APIs, as they enable more powerful, efficient, and accurate financial analysis.
