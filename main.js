// main.js: 應用程式主入口，處理 UI 互動與狀態管理

import { FULL_CONFIG, CONFIG_BUSINESS, PROFIT_LOSS_SUMMARY_TEMPLATE } from './config.js';
import { processFile } from './parsers.js';
import { exportData } from './utils.js';

// --- Global Variables & UI Elements (維持不變) ---
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

// --- Event Listeners & File Handling (維持不變) ---
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

// --- UI and State Management Functions (維持不變) ---
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
    const normalize = (name) => String(name || '').replace(/\s|　/g, '').split('(')[0];

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
            
            const key = normalize(originalKeyText);

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
    // (此函式維持不變)
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

// ★★★ 修正 createTabsAndTables 函式，重新啟用對營業基金的特殊處理 ★★★
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
                tableContent = createGovernmentalYuchuSummaryTable(data[reportKey]);
            } else if (selectedFundType === 'business' && reportKey === '損益表' && mode === 'sum') {
                // ★★★ 在加總模式下，為營業基金損益表呼叫新的特殊處理函式 ★★★
                tableContent = createBusinessProfitLossSummaryTable(data[reportKey]);
            }
            else {
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

// (createTableHtml, initTabs, initSortableTables, initExportButtons, createGovernmentalYuchuSummaryTable 維持不變)
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
            const indentLevel = record.indent_level || record.indent || 0;
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
function createGovernmentalYuchuSummaryTable(aggregatedData) {
    const yuchuData = aggregatedData;
    if (!yuchuData || yuchuData.length === 0) return '<p>無餘絀表資料可顯示。</p>';

    const findValue = (itemName, colName) => {
        const cleanItemName = itemName.replace(/\s|　/g, '');
        const row = yuchuData.find(r => String(r['項目']).replace(/\s|　/g, '') === cleanItemName);
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
        <tbody>
            <tr>
                <td>所有基金合計</td>
                <td>${findValue('基金來源', '預算數')}</td>
                <td>${findValue('基金用途', '預算數')}</td>
                <td>${findValue('本期賸餘(短絀)', '預算數')}</td>
                <td>${findValue('基金來源', '決算核定數')}</td>
                <td>${findValue('基金用途', '決算核定數')}</td>
                <td>${findValue('本期賸餘(短絀)', '決算核定數')}</td>
                <td>${findValue('基金來源', '預算與決算核定數比較增減')}</td>
                <td>${findValue('基金用途', '預算與決算核定數比較增減')}</td>
                <td>${findValue('本期賸餘(短絀)', '預算與決算核定數比較增減')}</td>
                <td>${findValue('期初基金餘額', '決算核定數')}</td>
                <td>${findValue('本期繳庫數', '決算核定數')}</td>
                <td>${findValue('期末基金餘額', '決算核定數')}</td>
            </tr>
        </tbody></table>`;
    return table;
}


// ★★★ 全新重寫的函式，用於產生標準化的營業基金損益總表 ★★★
function createBusinessProfitLossSummaryTable(aggregatedData) {
    if (!aggregatedData) return '<p>無損益表資料可顯示。</p>';

    const config = CONFIG_BUSINESS['損益表'];
    const keyColumn = config.keyColumn;
    const numericCols = config.columns.filter(c => c !== keyColumn);
    const normalize = (name) => String(name || '').replace(/\s|　/g, '').split('(')[0];

    // 1. 將加總後的資料轉為 Map，方便快速查找
    const dataMap = new Map();
    aggregatedData.forEach(row => {
        const key = normalize(row[keyColumn]);
        dataMap.set(key, row);
    });

    // 2. 準備一個 Map 來存放最終結果
    const resultsMap = new Map();

    // 3. 遍歷標準範本，從 dataMap 填充資料並執行特殊邏輯
    PROFIT_LOSS_SUMMARY_TEMPLATE.forEach(templateItem => {
        const record = { [keyColumn]: templateItem.name, indent: templateItem.indent };
        numericCols.forEach(col => record[col] = 0); // 初始化為 0

        if (templateItem.type === 'data') {
            // 尋找原始資料 (考慮別名)
            let sourceRow = dataMap.get(normalize(templateItem.source_name || templateItem.name));
            if (!sourceRow && templateItem.aliases) {
                for (const alias of templateItem.aliases) {
                    sourceRow = dataMap.get(normalize(alias));
                    if (sourceRow) break;
                }
            }

            if (sourceRow) {
                numericCols.forEach(col => {
                    record[col] = sourceRow[col] || 0;
                });
            }

            // ★★★ 處理中央銀行科目的合併 ★★★
            if (templateItem.special === 'merge_investment_profit') {
                const cbankProfit = dataMap.get(normalize('事業投資利益'));
                if (cbankProfit) {
                    numericCols.forEach(col => {
                        record[col] += cbankProfit[col] || 0;
                    });
                }
            }
            if (templateItem.special === 'merge_investment_loss') {
                const cbankLoss = dataMap.get(normalize('事業投資損失'));
                if (cbankLoss) {
                    numericCols.forEach(col => {
                        record[col] += cbankLoss[col] || 0;
                    });
                }
            }
            
            // ★★★ 處理重複科目名稱問題 ★★★
            // 假設：所有從原始資料來的 "採用權益法..." 都歸屬於營業項目。
            // 營業外項目預設為 0，除非有特殊邏輯指定。
            // 目前的範本設計已經透過 `source_name` 處理了這個問題，
            // 營業外的項目會去查找原始名稱，但因為原始名稱的資料已經被營業項目用掉了，
            // 這裡需要一個更明確的區分。
            // 簡化後的策略：我們假設原始資料無法區分，因此所有金額都計入營業部分。
            // 範本中的營業外項目將顯示為 0。
            if (templateItem.note === 'non-operating') {
                 numericCols.forEach(col => record[col] = 0);
            }
        }
        resultsMap.set(templateItem.name, record);
    });

    // 4. 遍歷範本，執行計算
    PROFIT_LOSS_SUMMARY_TEMPLATE.forEach(templateItem => {
        if (templateItem.type === 'calc' && templateItem.formula) {
            const recordToUpdate = resultsMap.get(templateItem.name);
            const [val1_name, operator, val2_name] = templateItem.formula;
            
            const val1_record = resultsMap.get(val1_name);
            const val2_record = resultsMap.get(val2_name);

            if (recordToUpdate && val1_record && val2_record) {
                numericCols.forEach(col => {
                    if (operator === '+') {
                        recordToUpdate[col] = (val1_record[col] || 0) + (val2_record[col] || 0);
                    } else if (operator === '-') {
                        recordToUpdate[col] = (val1_record[col] || 0) - (val2_record[col] || 0);
                    }
                });
            }
        } else if (templateItem.type === 'calc') {
            // 處理彙總型計算 (如營業收入 = 其下所有子項之和)
            const recordToUpdate = resultsMap.get(templateItem.name);
            let startIndex = PROFIT_LOSS_SUMMARY_TEMPLATE.findIndex(item => item.name === templateItem.name) + 1;
            let sumItems = [];
            for (let i = startIndex; i < PROFIT_LOSS_SUMMARY_TEMPLATE.length; i++) {
                const subItem = PROFIT_LOSS_SUMMARY_TEMPLATE[i];
                if (subItem.indent <= templateItem.indent) break; // 遇到同級或更上層的科目，停止加總
                if (subItem.type === 'data' || subItem.type === 'calc') {
                    sumItems.push(subItem.name);
                }
            }
            
            if (recordToUpdate && sumItems.length > 0) {
                 numericCols.forEach(col => {
                    recordToUpdate[col] = sumItems.reduce((total, itemName) => {
                        const itemRecord = resultsMap.get(itemName);
                        return total + (itemRecord ? itemRecord[col] : 0);
                    }, 0);
                });
            }
        }
    });

    // 5. 將 Map 轉為最終的陣列
    const finalRecords = Array.from(resultsMap.values());

    // 6. 呼叫通用的 HTML 產生器
    return createTableHtml(finalRecords, config.columns);
}
