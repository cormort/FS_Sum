// main.js: 應用程式主入口，處理 UI 互動與狀態管理

import { FULL_CONFIG, CONFIG_BUSINESS, BUSINESS_PROFIT_LOSS_ACCOUNTS } from './config.js';
import { processFile } from './parsers.js';
import { exportData } from './utils.js';

// --- Global Variables & UI Elements ---
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const statusDiv = document.getElementById('status');
const outputContainer = document.getElementById('output-container');
const controlsContainer = document.querySelector('.controls');
const typeSelector = document.getElementById('type-selector');
const mainContent = document.getElementById('main-content');
const resetButton = document.getElementById('reset-button');
const fundTypeDisplay = document.getElementById('current-fund-type-display');

const DEBUG_MODE = true; 
let allExtractedData = {};
let fundNames = [];
let selectedFundType = null;
let fundFileMap = {};

// --- Event Listeners ---
typeSelector.addEventListener('change', (e) => {
    if (e.target.name === 'fund-type') {
        selectedFundType = e.target.value;
        const labelText = document.querySelector(`label[for=${e.target.id}]`).textContent;
        fundTypeDisplay.textContent = `目前選擇：${labelText}`;
        mainContent.style.display = 'block';
        typeSelector.style.display = 'none';
    }
});

resetButton.addEventListener('click', () => resetState(true));

dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => handleFiles(e.target.files));
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('drag-over'); });
dropZone.addEventListener('drop', (e) => { e.preventDefault(); dropZone.classList.remove('drag-over'); handleFiles(e.dataTransfer.files); });

function handleFiles(files) {
    if (!selectedFundType) {
        alert('請先返回並選擇基金類型！');
        return;
    }
    if (!files || files.length === 0) return;
    if (DEBUG_MODE) console.clear();
    resetState(false);
    statusDiv.textContent = `讀取中... 偵測到 ${files.length} 個檔案。`;
    
    const filePromises = Array.from(files).map(file => processFile(file, selectedFundType));

    Promise.all(filePromises).then(results => {
        const successfulResults = results.filter(r => r);
        successfulResults.sort((a, b) => a.fileName.localeCompare(b.fileName, 'zh-Hant'));
        
        successfulResults.forEach(r => { fundFileMap[r.fundName] = r.fileName; });
        fundNames = successfulResults.map(r => r.fundName);
        
        successfulResults.forEach(result => {
            for (const reportKey in result.data) {
                if (!allExtractedData[reportKey]) allExtractedData[reportKey] = [];
                allExtractedData[reportKey].push(...result.data[reportKey]);
            }
        });

        if (fundNames.length > 0) {
            statusDiv.textContent = `處理完成！共 ${fundNames.length} 個基金。`;
            renderControls();
            refreshView();
        } else {
            statusDiv.textContent = '未找到任何可辨識的資料。請檢查檔案內容是否符合所選基金類型的格式。';
        }
    }).catch(error => {
        statusDiv.textContent = `處理檔案時發生嚴重錯誤：${error.message}`;
        console.error(error);
    });
}

// --- UI and State Management Functions ---
function resetState(fullReset = true) {
    allExtractedData = {};
    fundNames = [];
    fundFileMap = {};
    outputContainer.innerHTML = '';
    controlsContainer.innerHTML = '';
    statusDiv.textContent = '';
    fileInput.value = '';
    if (fullReset) {
        fundTypeDisplay.textContent = '';
        mainContent.style.display = 'none';
        typeSelector.style.display = 'block';
        const radios = document.querySelectorAll('input[name="fund-type"]');
        radios.forEach(radio => radio.checked = false);
        selectedFundType = null;
    }
}

function renderControls() {
    controlsContainer.innerHTML = `<div class="control-row"><div class="control-group"><label>檢視模式：</label><div class="mode-selector"><input type="radio" id="mode-individual" name="view-mode" value="individual" checked><label for="mode-individual">個別基金</label><input type="radio" id="mode-sum" name="view-mode" value="sum"><label for="mode-sum">所有基金加總</label><input type="radio" id="mode-compare" name="view-mode" value="compare"><label for="mode-compare">單項比較</label></div></div></div><div class="control-row" id="dynamic-controls"></div>`;
    controlsContainer.querySelector('.mode-selector').addEventListener('change', refreshView);
}

function refreshView() {
    const selectedMode = document.querySelector('input[name="view-mode"]:checked').value;
    const dynamicControlsContainer = document.getElementById('dynamic-controls');
    dynamicControlsContainer.innerHTML = '';
    if (selectedMode === 'individual') {
        dynamicControlsContainer.innerHTML = `<div class="control-group"><label for="fund-select">選擇基金：</label><select id="fund-select">${fundNames.map(name => `<option value="${name}">${name}</option>`).join('')}</select></div>`;
        document.getElementById('fund-select').addEventListener('change', displayIndividualFund);
        displayIndividualFund();
    } else if (selectedMode === 'sum') {
        displayAggregated();
    } else if (selectedMode === 'compare') {
        displayComparison();
    }
}

function displayIndividualFund() {
    const selectedFund = document.getElementById('fund-select')?.value;
    if (!selectedFund) return;
    const fundData = {};
    for (const reportKey in allExtractedData) {
        if(allExtractedData[reportKey]) {
            fundData[reportKey] = allExtractedData[reportKey].filter(r => r['基金名稱'] === selectedFund);
        }
    }
    outputContainer.innerHTML = createTabsAndTables(fundData);
    initTabs();
    initExportButtons();
}

function displayAggregated() {
    const summaryData = {};
    for(const reportKey in allExtractedData) {
        if(!allExtractedData[reportKey]) continue;
        
        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = FULL_CONFIG[selectedFundType][baseKey];
        if (!config) continue;

        const keyColumn = config.keyColumn;
        const numericCols = config.columns.filter(c => c !== keyColumn);

        const grouped = allExtractedData[reportKey].reduce((acc, row) => {
            const originalKeyText = row[keyColumn];
            if(!originalKeyText) return acc;
            const key = String(originalKeyText).replace(/\s|　/g, '');
            if (!acc[key]) {
                acc[key] = { [keyColumn]: originalKeyText.trim() };
                numericCols.forEach(col => acc[key][col] = 0);
                acc[key].indent_level = row.indent_level || 0;
            }
            numericCols.forEach(col => {
                let val = parseFloat(String(row[col] || '0').replace(/,/g, ''));
                if (!isNaN(val)) {
                    if (selectedFundType === 'business') {
                        val = Math.round(val);
                    }
                    acc[key][col] += val;
                }
            });
            return acc;
        }, {});
        summaryData[reportKey] = Object.values(grouped);
    }
    outputContainer.innerHTML = createTabsAndTables(summaryData, {}, 'sum');
    initTabs();
    initExportButtons();
}

function displayComparison() {
    const dynamicControlsContainer = document.getElementById('dynamic-controls');
    let optionsHtml = '';
    const activeConfig = FULL_CONFIG[selectedFundType];
    const reportKeysInData = Object.keys(allExtractedData).sort();

    reportKeysInData.forEach(reportKey => {
        if (!allExtractedData[reportKey] || allExtractedData[reportKey].length === 0) return;
        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = activeConfig[baseKey];
        if (!config || !config.keyColumn) return;

        const keyColumn = config.keyColumn;
        const items = [...new Set(allExtractedData[reportKey].map(r => r[keyColumn]))].filter(Boolean).sort((a,b) => a.localeCompare(b, 'zh-Hant'));
        const tabName = reportKey.replace(/_/g, ' ');

        if (items.length > 0) {
             optionsHtml += `<optgroup label="${tabName}">${items.map(item => `<option value="${reportKey}::${item}">${item}</option>`).join('')}</optgroup>`;
        }
    });

    dynamicControlsContainer.innerHTML = `<div class="control-group"><label for="item-select">選擇比較項目：</label><select id="item-select">${optionsHtml}</select></div>`;
    const itemSelect = document.getElementById('item-select');
    
    const updateComparisonView = () => {
        const selectedItem = itemSelect?.value;
        if (!selectedItem) {
            outputContainer.innerHTML = '';
            return;
        };
        const [reportKey, itemName] = selectedItem.split('::');
        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = FULL_CONFIG[selectedFundType][baseKey];
        const keyColumn = config.keyColumn;
        const columns = config.columns;
        const dataForCompare = allExtractedData[reportKey].filter(r => r[keyColumn] === itemName);
        const finalData = { [reportKey]: [] };
        
        const totals = {};
        const numericHeaders = columns.filter(c => c !== keyColumn);
        numericHeaders.forEach(h => totals[h] = 0);

        fundNames.forEach(fund => {
            const fundRow = dataForCompare.find(r => r['基金名稱'] === fund);
            const newRow = { '基金名稱': fund };
            columns.forEach(col => {
                if (col !== keyColumn) {
                    const val = fundRow ? fundRow[col] : '-';
                    newRow[col] = val;
                    const numVal = parseFloat(String(val).replace(/,/g, '')) || 0;
                    if (totals[col] !== undefined) {
                        totals[col] += numVal;
                    }
                }
            });
            finalData[reportKey].push(newRow);
        });

        const totalRow = { '基金名稱': `${itemName}合計` };
        Object.assign(totalRow, totals);
        finalData[reportKey].unshift(totalRow);

        const headers = ['基金名稱', ...columns.filter(c => c !== keyColumn)];
        outputContainer.innerHTML = createTabsAndTables(finalData, { [reportKey]: headers }, 'comparison');
        initTabs();
        initSortableTables();
        initExportButtons();
    };

    itemSelect.addEventListener('change', updateComparisonView);
    if (itemSelect.options.length > 0) {
        updateComparisonView();
    }
}

// --- HTML Rendering Functions ---
function createTabsAndTables(data, customHeaders = {}, mode = 'default') {
    let tabsHtml = '<div class="report-tabs">';
    let contentHtml = '';
    let isFirst = true;
    const activeConfig = FULL_CONFIG[selectedFundType];
    const reportKeysInOrder = Object.keys(activeConfig);

    const allPossibleKeys = [];
    reportKeysInOrder.forEach(k => {
        if (k === '平衡表' || k === '資產負債表') {
            allPossibleKeys.push(k + '_資產', k + '_負債及權益');
        } else {
            allPossibleKeys.push(k);
        }
    });

    const dataKeys = Object.keys(data);
    const orderedDataKeys = allPossibleKeys.filter(k => dataKeys.includes(k));

    orderedDataKeys.forEach(reportKey => {
        if (data[reportKey] && data[reportKey].length > 0) {
            const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
            const config = activeConfig[baseKey];
            if (!config) return;

            const columns = config.columns;
            const tabName = reportKey.replace(/_/g, ' ');
            tabsHtml += `<button class="tab-link ${isFirst ? 'active' : ''}" data-tab="${reportKey}">${tabName}</button>`;
            
            let tableContent;
            if (selectedFundType === 'governmental' && reportKey === '餘絀表' && mode === 'sum') {
                tableContent = createGovernmentalYuchuSummaryTable();
            } else if (selectedFundType === 'business' && reportKey === '損益表' && mode === 'sum') {
                tableContent = createBusinessProfitLossSummaryTable(data[reportKey]);
            } else {
                tableContent = createTableHtml(data[reportKey], customHeaders[reportKey] || ['基金名稱', ...columns], mode);
            }

            const exportButtons = `
                <div class="export-buttons">
                    <button class="export-btn json" data-format="json" data-report-key="${reportKey}">匯出 JSON</button>
                    <button class="export-btn xlsx" data-format="xlsx" data-report-key="${reportKey}">匯出 XLSX</button>
                    <button class="export-btn" data-format="html" data-report-key="${reportKey}">匯出 HTML</button>
                </div>`;

            contentHtml += `<div id="${reportKey}" class="tab-content ${isFirst ? 'active' : ''}">
                <div class="tab-header">
                    <h2>${tabName}</h2>
                    ${exportButtons}
                </div>
                ${tableContent}
            </div>`;
            isFirst = false;
        }
    });
    tabsHtml += '</div>';
    if (contentHtml === '') {
        return '<p>無資料可顯示。</p>';
    }
    return tabsHtml + contentHtml;
}

function createTableHtml(records, headers, mode = 'default') {
    let table = '<table><thead><tr>';
    const keyColumns = ['科目', '項目'];
    table += headers.map(h => {
        const isSortable = mode === 'comparison' && !['基金名稱', ...keyColumns].includes(h);
        return `<th class="${isSortable ? 'sortable' : ''}" data-column-key="${h}">${h} <span class="sort-arrow"></span></th>`;
    }).join('');
    table += '</tr></thead><tbody>';

    records.forEach(record => {
        table += '<tr>';
        headers.forEach(header => {
            const val = record[header];
            let displayVal = (val === null || val === undefined) ? '' : val;
            const isKeyColumn = keyColumns.includes(header);
            const indentLevel = record.indent_level || 0;
            let style = '';
            if (isKeyColumn && indentLevel > 0) {
                style = `padding-left: ${1 + indentLevel * 1.5}em;`;
            }
            const isNumericField = !['基金名稱', ...keyColumns].includes(header);
            if (isNumericField && val != null && val !== '' && !isNaN(Number(String(val).replace(/,/g, '')))) {
                 displayVal = Number(String(val).replace(/,/g, '')).toLocaleString();
            }
            table += `<td style="${style}">${displayVal}</td>`;
        });
        table += '</tr>';
    });
    return table + '</tbody></table>';
}

function initTabs() {
    const tabs = document.querySelectorAll('.report-tabs .tab-link');
    const contents = document.querySelectorAll('.tab-content');
    tabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            tabs.forEach(t => t.classList.remove('active'));
            contents.forEach(c => c.classList.remove('active'));
            const tabId = e.currentTarget.dataset.tab;
            e.currentTarget.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
    if(tabs.length > 0) {
        const firstActiveTab = document.querySelector('.report-tabs .tab-link.active');
        if(!firstActiveTab && contents.length > 0) {
            tabs[0].classList.add('active');
            contents[0].classList.add('active');
        }
    }
}

function initSortableTables() {
    document.querySelectorAll('.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const table = th.closest('table');
            const tbody = table.querySelector('tbody');
            const headerIndex = Array.from(th.parentNode.children).indexOf(th);
            const currentIsDesc = th.classList.contains('sort-desc');

            table.querySelectorAll('th').forEach(h => {
                h.classList.remove('sort-asc', 'sort-desc');
                const arrow = h.querySelector('.sort-arrow');
                if(arrow) arrow.textContent = '';
            });

            let direction = currentIsDesc ? 'asc' : 'desc';
            th.classList.add(direction === 'asc' ? 'sort-asc' : 'sort-desc');
            
            const arrow = th.querySelector('.sort-arrow');
            if(arrow) arrow.textContent = direction === 'asc' ? '▲' : '▼';

            const rows = Array.from(tbody.querySelectorAll('tr'));
            const headerRow = rows.shift(); // Keep total row at top

            rows.sort((a, b) => {
                const aValText = a.children[headerIndex].textContent;
                const bValText = b.children[headerIndex].textContent;
                const aVal = parseFloat(String(aValText).replace(/,/g, '')) || Number.NEGATIVE_INFINITY;
                const bVal = parseFloat(String(bValText).replace(/,/g, '')) || Number.NEGATIVE_INFINITY;
                return direction === 'asc' ? aVal - bVal : bVal - aVal;
            })
            .forEach(row => tbody.appendChild(row));
            tbody.insertBefore(headerRow, tbody.firstChild);
        });
    });
}

function initExportButtons() {
    document.querySelectorAll('.export-btn').forEach(button => {
        button.addEventListener('click', (e) => {
            const reportKey = e.target.dataset.reportKey;
            const format = e.target.dataset.format;
            exportData(reportKey, format);
        });
    });
}

// --- Special Report Generation ---
function createGovernmentalYuchuSummaryTable() {
    const yuchuData = allExtractedData['餘絀表'];
    if (!yuchuData) return '<p>無餘絀表資料可顯示。</p>';

    const dataByFund = yuchuData.reduce((acc, row) => {
        const fundName = row['基金名稱'];
        if (!acc[fundName]) acc[fundName] = [];
        acc[fundName].push(row);
        return acc;
    }, {});

    const sortedFundNames = Object.keys(dataByFund).sort((a, b) => {
        return fundFileMap[a].localeCompare(fundFileMap[b], 'zh-Hant');
    });

    const findValue = (fundData, itemName, colName) => {
        const cleanItemName = itemName.replace(/\s|　/g, '');
        const row = fundData.find(r => String(r['項目']).replace(/\s|　/g, '') === cleanItemName);
        const value = row ? row[colName] : 0;
        if (value != null && value !== '' && !isNaN(Number(String(value).replace(/,/g, '')))) {
            return Number(String(value).replace(/,/g, '')).toLocaleString();
        }
        return '0';
    };

    let table = `<table>
        <thead>
            <tr>
                <th rowspan="2">基金別</th>
                <th colspan="3">預算數</th>
                <th colspan="3">決算核定數</th>
                <th colspan="3">決算核定數與預算數比較</th>
                <th rowspan="2">期初基金餘額</th>
                <th rowspan="2">本期繳庫數</th>
                <th rowspan="2">期末基金餘額</th>
            </tr>
            <tr>
                <th>基金來源</th><th>基金用途</th><th>賸餘(短絀)</th>
                <th>基金來源</th><th>基金用途</th><th>賸餘(短絀)</th>
                <th>基金來源</th><th>基金用途</th><th>賸餘(短絀)</th>
            </tr>
        </thead>
        <tbody>`;

    sortedFundNames.forEach(fundName => {
        const fundData = dataByFund[fundName];
        table += `<tr>
            <td>${fundName}</td>
            <td>${findValue(fundData, '基金來源', '預算數')}</td>
            <td>${findValue(fundData, '基金用途', '預算數')}</td>
            <td>${findValue(fundData, '本期賸餘(短絀)', '預算數')}</td>
            <td>${findValue(fundData, '基金來源', '決算核定數')}</td>
            <td>${findValue(fundData, '基金用途', '決算核定數')}</td>
            <td>${findValue(fundData, '本期賸餘(短絀)', '決算核定數')}</td>
            <td>${findValue(fundData, '基金來源', '預算與決算核定數比較增減')}</td>
            <td>${findValue(fundData, '基金用途', '預算與決算核定數比較增減')}</td>
            <td>${findValue(fundData, '本期賸餘(短絀)', '預算與決算核定數比較增減')}</td>
            <td>${findValue(fundData, '期初基金餘額', '決算核定數')}</td>
            <td>${findValue(fundData, '本期繳庫數', '決算核定數')}</td>
            <td>${findValue(fundData, '期末基金餘額', '決算核定數')}</td>
        </tr>`;
    });

    table += '</tbody></table>';
    return table;
}

function createBusinessProfitLossSummaryTable(aggregatedData) {
    if (!aggregatedData) return '<p>無損益表資料可顯示。</p>';

    const config = CONFIG_BUSINESS['損益表'];
    const keyColumn = config.keyColumn;
    const numericCols = config.columns.filter(c => c !== keyColumn);
    const normalize = (name) => String(name || '').replace(/\s|　/g, '').split('(')[0];
    
    const results = new Map();
    const structuredAccounts = BUSINESS_PROFIT_LOSS_ACCOUNTS.filter(acc => [1, 2, 4].includes(acc.code.length));

    structuredAccounts.forEach(accDef => {
        const displayName = (accDef.merge_target || accDef.name).split('(')[0].trim();
        if (!results.has(displayName)) {
            const newRecord = { [keyColumn]: displayName };
            numericCols.forEach(col => newRecord[col] = 0);
            if (accDef.code.length === 2) newRecord.indent_level = 1;
            else if (accDef.code.length === 4) newRecord.indent_level = 2;
            else newRecord.indent_level = 0;
            results.set(displayName, newRecord);
        }
    });

    aggregatedData.forEach(row => {
        const accountDef = structuredAccounts.find(def => normalize(def.name) === normalize(row[keyColumn]));
        if (accountDef && !accountDef.type) {
            const displayName = (accountDef.merge_target || accountDef.name).split('(')[0].trim();
            const recordToUpdate = results.get(displayName);
            if (recordToUpdate) {
                numericCols.forEach(col => {
                    recordToUpdate[col] += row[col] || 0;
                });
            }
        }
    });
    
    const事業投資利益 = aggregatedData.find(r => normalize(r[keyColumn]) === '事業投資利益');
    const事業投資損益 = aggregatedData.find(r => normalize(r[keyColumn]) === '事業投資損益');
    const targetProfitAcc = results.get('採用權益法認列之關聯企業及合資利益之份額');
    const targetLossAcc = results.get('採用權益法認列之關聯企業及合資損失之份額');

    if (事業投資利益 && targetProfitAcc) {
        numericCols.forEach(col => { targetProfitAcc[col] += 事業投資利益[col] || 0; });
    }
    if (事業投資損益 && targetLossAcc) {
        numericCols.forEach(col => { targetLossAcc[col] += 事業投資損益[col] || 0; });
    }

    const getVal = (name, col) => results.get(name)?.[col] || 0;
    
    const incomeAccounts = ['銷售收入', '勞務收入', '金融保險收入', '採用權益法認列之關聯企業及合資利益之份額', '其他營業收入'];
    const costAccounts = ['銷售成本', '勞務成本', '金融保險成本', '採用權益法認列之關聯企業及合資損失之份額', '其他營業成本'];
    const expenseAccounts = ['行銷費用', '業務費用', '管理費用', '其他營業費用'];
    const nonOpIncomeAccounts = ['財務收入', '採用權益法認列之關聯企業及合資利益之份額', '其他營業外收入'];
    const nonOpExpenseAccounts = ['財務成本', '採用權益法認列之關聯企業及合資損失之份額', '其他營業外費用'];

    numericCols.forEach(col => {
        results.get('營業收入')[col] = incomeAccounts.reduce((sum, acc) => sum + getVal(acc, col), 0);
        results.get('營業成本')[col] = costAccounts.reduce((sum, acc) => sum + getVal(acc, col), 0);
        results.get('營業費用')[col] = expenseAccounts.reduce((sum, acc) => sum + getVal(acc, col), 0);
        results.get('營業外收入')[col] = nonOpIncomeAccounts.reduce((sum, acc) => sum + getVal(acc, col), 0);
        results.get('營業外費用')[col] = nonOpExpenseAccounts.reduce((sum, acc) => sum + getVal(acc, col), 0);

        results.get('營業毛利(毛損)')[col] = getVal('營業收入', col) - getVal('營業成本', col);
        results.get('營業利益(損失)')[col] = getVal('營業毛利(毛損)', col) - getVal('營業費用', col);
        results.get('營業外利益(損失)')[col] = getVal('營業外收入', col) - getVal('營業外費用', col);
        results.get('稅前淨利(淨損)')[col] = getVal('營業利益(損失)', col) + getVal('營業外利益(損失)', col);
        results.get('繼續營業單位本期淨利(淨損)')[col] = getVal('稅前淨利(淨損)', col) - getVal('所得稅費用(利益)', col);
        results.get('本期淨利(淨損)')[col] = getVal('繼續營業單位本期淨利(淨損)', col) + getVal('停業單位損益', col);
        results.get('本期淨利(淨損)歸屬於:')[col] = getVal('本期淨利(淨損)', col);
        
        results.get('收入')[col] = getVal('營業收入', col) + getVal('營業外收入', col);
        results.get('支出')[col] = getVal('營業成本', col) + getVal('營業費用', col) + getVal('營業外費用', col) + getVal('所得稅費用(利益)', col);
    });

    const finalRecords = Array.from(results.values());
    return createTableHtml(finalRecords, config.columns);
}