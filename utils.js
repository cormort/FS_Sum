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
export function findHeaderRowIndex(data, columns) {
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
