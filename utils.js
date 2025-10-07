// utils.js: 存放輔助工具函式

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
        const wb = XLSX.utils.table_to_book(table, { raw: true }); // Use raw to export numbers not strings
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