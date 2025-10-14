// main.js

import { FULL_CONFIG, PROFIT_LOSS_ACCOUNT_ORDER, APPROPRIATION_ACCOUNT_ORDER } from './config.js';
import { processFile } from './parsers.js';
import { exportData } from './utils.js';

// ★★★ 內建公版資產負債表科目順序樣板 ★★★
const PUBLIC_ASSET_ORDER = [
    "資產", "流動資產", "現金", "存放銀行同業", "存放央行", "流動金融資產", "應收款項", 
    "本期所得稅資產", "黃金與白銀", "存貨", "消耗性生物資產－流動", "生產性生物資產－流動", 
    "預付款項", "短期墊款", "待出售非流動資產", "合約資產－流動", "其他流動資產", 
    "押匯貼現及放款", "押匯及貼現", "短期放款及透支", "短期擔保放款及透支", "中期放款", 
    "中期擔保放款", "長期放款", "長期擔保放款", "銀行業融通", "基金、投資及長期應收款", 
    "基金", "非流動金融資產", "採用權益法之投資", "其他長期投資", "長期應收款項", 
    "再保險準備資產", "合約資產－非流動", "不動產、廠房及設備", "土地", "土地改良物", 
    "房屋及建築", "機械及設備", "交通及運輸設備", "什項設備", "租賃權益改良", 
    "購建中固定資產", "核能燃料", "生產性植物", "使用權資產", "投資性不動產", 
    "投資性不動產－土地", "投資性不動產－土地改良物", "投資性不動產－房屋及建築", 
    "投資性不動產－租賃權益改良", "建造中之投資性不動產", "無形資產", "累計減損－無形資產", 
    "生物資產", "消耗性生物資產－非流動", "生產性生物資產－非流動", "其他資產", 
    "遞延資產", "遞延所得稅資產", "待整理資產", "什項資產", "合　　計"
];

const PUBLIC_LIABILITY_ORDER = [
    "負債", "流動負債", "短期債務", "央行存款", "銀行同業存款", "國際金融機構存款", 
    "應付款項", "本期所得稅負債", "發行券幣", "預收款項", "流動金融負債", 
    "與待出售處分群組直接相關之負債", "合約負債－流動", "其他流動負債", "存款、匯款及金融債券", 
    "支票存款", "活期存款", "定期存款", "儲蓄存款", "匯款", "金融債券", "央行及同業融資", 
    "央行融資", "同業融資", "長期負債", "長期債務", "租賃負債", "非流動金融負債", "其他負債", 
    "負債準備", "遞延負債", "遞延所得稅負債", "待整理負債", "合約負債－非流動", "什項負債", 
    "權益", "資本", "預收資本", "資本公積", "保留盈餘（或累積虧損）", "已指撥保留盈餘", 
    "未指撥保留盈餘", "累積虧損", "累積其他綜合損益", "國外營運機構財務報表換算之兌換差額", 
    "現金流量避險中屬有效避險部分之避險工具利益（損失）", 
    "國外營運機構淨投資避險中屬有效避險部分之避險工具利益（損失）", "不動產重估增值", 
    "與待出售非流動資產直接相關之權益", "確定福利計畫之再衡量數", 
    "指定為透過損益按公允價值衡量之金融負債其變動金額來自信用風險", 
    "透過其他綜合損益按公允價值衡量之金融資產損益", "採用覆蓋法重分類之其他綜合損益", 
    "其他權益－其他", "庫藏股票", "首次採用國際財務報導準則調整數", "非控制權益", "合　　計"
];


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

    for (const reportKey in allExtractedData) {
        if (!allExtractedData[reportKey] || allExtractedData[reportKey].length === 0) continue;

        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = FULL_CONFIG[selectedFundType]?.[baseKey];
        if (!config) continue;

        const keyColumn = config.keyColumn;
        const numericCols = config.columns.filter(c => c !== keyColumn);

        // --- 步驟一：精確累加所有資料 ---
        // 使用 "科目名稱::層級" 作為唯一鍵，確保父子科目被視為不同實體，其資料被分開並正確累加。
        const initialDataMap = new Map();
        allExtractedData[reportKey].forEach(row => {
            const keyText = row[keyColumn]?.trim();
            if (!keyText) return;
            const indent = row.indent_level || row.indent || 0;
            const compositeKey = `${keyText}::${indent}`;

            if (!initialDataMap.has(compositeKey)) {
                const newRow = { ...row, [keyColumn]: keyText, 'indent_level': indent };
                numericCols.forEach(col => {
                    newRow[col] = parseFloat(String(newRow[col] || '0').replace(/,/g, '')) || 0;
                });
                initialDataMap.set(compositeKey, newRow);
            } else {
                const existingRow = initialDataMap.get(compositeKey);
                numericCols.forEach(col => {
                    const val = parseFloat(String(row[col] || '0').replace(/,/g, ''));
                    if (!isNaN(val)) existingRow[col] += val;
                });
            }
        });
        
        // 對營業基金的數值進行四捨五入
        if (selectedFundType === 'business') {
             initialDataMap.forEach(row => {
                numericCols.forEach(col => row[col] = Math.round(row[col]));
            });
        }
        
        // --- 步驟二：建立最終報表容器，並以公版順序為基礎 ---
        // reportMap 的鍵是簡單的科目名稱，值是完整的資料行物件。
        const reportMap = new Map();
        const standardOrder = {
            '損益表': PROFIT_LOSS_ACCOUNT_ORDER,
            '盈虧撥補表': APPROPRIATION_ACCOUNT_ORDER,
            '資產負債表_資產': PUBLIC_ASSET_ORDER,
            '資產負債表_負債及權益': PUBLIC_LIABILITY_ORDER
        }[reportKey] || [];

        // 預先填入所有公版科目，確保順序
        standardOrder.forEach(key => {
            const newRow = { [keyColumn]: key };
            numericCols.forEach(col => newRow[col] = 0);
            reportMap.set(key, newRow);
        });

        // 將步驟一累加好的資料填入 reportMap。因為父子科目已分開，這裡會正確地將它們各自的金額加總。
        initialDataMap.forEach(row => {
            const accountName = row[keyColumn];
            // 如果公版中沒有，也新增進去，確保不遺漏
            if (!reportMap.has(accountName)) {
                const newRow = { [keyColumn]: accountName, indent_level: row.indent_level };
                numericCols.forEach(col => newRow[col] = 0);
                reportMap.set(accountName, newRow);
            }
            const targetRow = reportMap.get(accountName);
            numericCols.forEach(col => {
                targetRow[col] += row[col] || 0;
            });
            // 繼承原始的 indent level
            targetRow.indent_level = row.indent_level;
        });

        // --- 步驟三：應用您提供的所有例外規則 ---
        const mergeRules = {
            '損益表': [
                { target: '採用權益法認列之關聯企業及合資利益之份額', sources: ['事業投資利益'], type: 'accumulator' },
                { target: '採用權益法認列之關聯企業及合資損失之份額', sources: ['事業投資損失'], type: 'accumulator' },
            ],
            '資產負債表_資產': [
                { target: '存放銀行同業', sources: ['存放銀行業'], type: 'accumulator' },
                { target: '採用權益法之投資', sources: ['事業投資', '其他長期投資'], type: 'accumulator' },
                { target: '押匯貼現及放款', sources: ['融通', '銀行業融通', '押匯及貼現', '短期放款及透支', '短期擔保放款及透支', '中期放款', '中期擔保放款', '長期放款', '長期擔保放款'], type: 'summary' },
            ],
            '資產負債表_負債及權益': [
                { target: '銀行同業存款', sources: ['銀行業存款'], type: 'accumulator' },
                { target: '存款、匯款及金融債券', sources: ['存款'], type: 'accumulator' },
                { target: '支票存款', sources: ['公庫及政府機關存款'], type: 'accumulator' },
                { target: '儲蓄存款', sources: ['儲蓄存款及儲蓄券'], type: 'accumulator' }
            ]
        };

        const activeMergeRules = mergeRules[reportKey] || [];
        const sourceKeysToHide = new Set(); // 用於後續隱藏被合併的來源科目

        activeMergeRules.forEach(rule => {
            const targetRow = reportMap.get(rule.target);
            if (!targetRow) return;

            if (rule.type === 'accumulator') { // 累加型
                rule.sources.forEach(sourceKey => {
                    sourceKeysToHide.add(sourceKey);
                    const sourceRow = reportMap.get(sourceKey);
                    if (sourceRow) {
                        numericCols.forEach(col => targetRow[col] += (sourceRow[col] || 0));
                    }
                });
            } else if (rule.type === 'summary') { // 彙總型
                // 先清空目標科目的原始值
                numericCols.forEach(col => targetRow[col] = 0);
                rule.sources.forEach(sourceKey => {
                    sourceKeysToHide.add(sourceKey);
                    const sourceRow = reportMap.get(sourceKey);
                    if (sourceRow) {
                        numericCols.forEach(col => targetRow[col] += (sourceRow[col] || 0));
                    }
                });
            }
        });

        // --- 步驟四：組裝最終結果 ---
        const finalRows = [];
        const processedKeys = new Set();

        // 首先，按照公版順序添加科目
        standardOrder.forEach(key => {
            if (!sourceKeysToHide.has(key) && reportMap.has(key)) {
                finalRows.push(reportMap.get(key));
                processedKeys.add(key);
            }
        });

        // 其次，添加不在公版中、也未被隱藏的額外科目
        reportMap.forEach((row, key) => {
            if (!processedKeys.has(key) && !sourceKeysToHide.has(key)) {
                finalRows.push(row);
            }
        });

        summaryData[reportKey] = finalRows;
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
    const keyColumns = ['科目', '項目', '基金名稱']; 
    table += headers.map(h => {
        const isSortable = mode === 'comparison' && !keyColumns.includes(h);
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
            let className = '';

            if (isKeyColumn) {
                if (header !== '基金名稱' && indentLevel > 0) {
                    style = `padding-left: ${1 + indentLevel * 1.5}em;`;
                }
            } else {
                const rawVal = String(val).replace(/,/g, '');
                const isNumericField = val != null && val !== '' && !isNaN(Number(rawVal)) && isFinite(Number(rawVal));
                
                if (isNumericField) {
                    const numVal = Number(rawVal);
                    displayVal = numVal.toLocaleString();
                    className = 'numeric-data';
                    if (numVal < 0) {
                        className += ' negative-value';
                    }
                } else if (val != null && val !== '') {
                    // Non-numeric but non-empty stays left-aligned
                } else {
                    displayVal = '';
                }
            }
            
            table += `<td class="${className}" style="${style}">${displayVal}</td>`;
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
                <td class="numeric-data">${findValue('基金來源', '預算數')}</td>
                <td class="numeric-data">${findValue('基金用途', '預算數')}</td>
                <td class="numeric-data">${findValue('本期賸餘(短絀)', '預算數')}</td>
                <td class="numeric-data">${findValue('基金來源', '決算核定數')}</td>
                <td class="numeric-data">${findValue('基金用途', '決算核定數')}</td>
                <td class="numeric-data">${findValue('本期賸餘(短絀)', '決算核定數')}</td>
                <td class="numeric-data">${findValue('基金來源', '預算與決算核定數比較增減')}</td>
                <td class="numeric-data">${findValue('基金用途', '預算與決算核定數比較增減')}</td>
                <td class="numeric-data">${findValue('本期賸餘(短絀)', '預算與決算核定數比較增減')}</td>
                <td class="numeric-data">${findValue('期初基金餘額', '決算核定數')}</td>
                <td class="numeric-data">${findValue('本期繳庫數', '決算核定數')}</td>
                <td class="numeric-data">${findValue('期末基金餘額', '決算核定數')}</td>
            </tr>
        </tbody></table>`;
    return table;
}
