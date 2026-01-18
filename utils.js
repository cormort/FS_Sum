// utils.js

// --- 檔案與工作表相關 ---
export function findSheet(workbook, keyword) { 
    return workbook.SheetNames.find(name => name.includes(keyword)); 
}

export function extractFundName(workbook) {
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

// --- 動態表格解析相關 (for 作業基金) ---
// 尋找表頭所在的列索引
// 支援空白正規化與部分關鍵字匹配
export function findHeaderRowIndex(data, columns) {
    if (!data || data.length === 0) return -1;
    let bestMatch = { rowIndex: -1, score: 0 };
    
    // 預先處理要尋找的 columns，移除空白以利比對
    const normalizedTargets = columns.map(c => c.replace(/\s|　/g, ''));

    // 掃描前 20 列 (避免表頭過深)
    const scanLimit = Math.min(data.length, 20);
    
    for (let i = 0; i < scanLimit; i++) {
        const row = data[i];
        if (!Array.isArray(row)) continue;
        
        // 正規化該列所有儲存格內容
        const normalizedRow = row.map(cell => String(cell || '').replace(/\s|　/g, ''));
        
        let score = 0;
        normalizedTargets.forEach(target => {
            // 檢查列中是否包含目標欄位 (雙向包含：目標在儲存格內，或儲存格是目標的子字串)
            // 例如：Target="本年度預算數", Cell="預算數" -> Match
            // 例如：Target="科目", Cell="科　　目" -> Match (經正規化後)
            const match = normalizedRow.some(cellVal => {
                if (!cellVal) return false;
                return cellVal.includes(target) || target.includes(cellVal);
            });
            if (match) score++;
        });

        if (score > bestMatch.score) {
            bestMatch = { rowIndex: i, score };
        }
    }

    // 只要有匹配到任何欄位，就視為找到 (由 >1 放寬為 >0，因有些表可能極簡)
    // 但為避免誤判，若 score 過低仍需謹慎，先維持 > 0 即可，因為我們會取最佳解
    return bestMatch.score > 0 ? bestMatch.rowIndex : -1;
}

export function getHeaderMapping(headerRow, columns, startCol = 0) {
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

// --- 資料匯出相關 ---
function getTableDataAsJSON(table) {
    const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent.trim());
    const rows = Array.from(table.querySelectorAll('tbody tr'));
    return rows.map(row => {
        const rowData = {};
        Array.from(row.children).forEach((cell, i) => {
            rowData[headers[i]] = cell.textContent.trim();
        });
        return rowData;
    });
}

export function exportData(reportKey, format) {
    const tabContent = document.getElementById(reportKey);
    if (!tabContent) return;
    const table = tabContent.querySelector('table');
    if (!table) return;

    const now = new Date();
    const timestamp = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}_${now.getHours().toString().padStart(2,'0')}${now.getMinutes().toString().padStart(2,'0')}${now.getSeconds().toString().padStart(2,'0')}`;
    const filename = `${reportKey}_${timestamp}`;

    if (format === 'xlsx') {
        const wb = XLSX.utils.table_to_book(table, { raw: true });
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else {
        let data, mimeType, fileExtension;
        if (format === 'json') {
            const jsonData = getTableDataAsJSON(table);
            data = JSON.stringify(jsonData, null, 2);
            mimeType = 'application/json';
            fileExtension = 'json';
        } else if (format === 'html') {
            const htmlTemplate = `
                <!DOCTYPE html>
                <html lang="zh-Hant">
                <head>
                    <meta charset="UTF-8">
                    <title>匯出資料: ${reportKey}</title>
                    <style>
                        body { font-family: sans-serif; }
                        table { border-collapse: collapse; width: 100%; }
                        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                        thead { background-color: #f2f2f2; }
                        tr:nth-child(even) { background-color: #f9f9f9; }
                    </style>
                </head>
                <body>
                    <h1>${reportKey}</h1>
                    ${table.outerHTML}
                </body>
                </html>`;
            data = htmlTemplate;
            mimeType = 'text/html';
            fileExtension = 'html';
        }

        if (data) {
            const blob = new Blob([data], { type: mimeType });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${filename}.${fileExtension}`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
    }
}
