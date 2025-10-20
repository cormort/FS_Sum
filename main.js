// main.js

import { FULL_CONFIG, PROFIT_LOSS_ACCOUNT_ORDER, APPROPRIATION_ACCOUNT_ORDER, CASH_FLOW_ACCOUNT_ORDER } from './config.js';
import { processFile } from './parsers.js';
import { exportData } from './utils.js';

// --- 公版順序樣板 ---
const PUBLIC_ASSET_ORDER = [ "資產", "流動資產", "現金", "存放銀行同業", "存放央行", "流動金融資產", "應收款項", "本期所得稅資產", "黃金與白銀", "存貨", "消耗性生物資產－流動", "生產性生物資產－流動", "預付款項", "短期墊款", "待出售非流動資產", "合約資產－流動", "其他流動資產", "押匯貼現及放款", "押匯及貼現", "短期放款及透支", "短期擔保放款及透支", "中期放款", "中期擔保放款", "長期放款", "長期擔保放款", "銀行業融通", "基金、投資及長期應收款", "基金", "非流動金融資產", "採用權益法之投資", "其他長期投資", "長期應收款項", "再保險準備資產", "合約資產－非流動", "不動產、廠房及設備", "土地", "土地改良物", "房屋及建築", "機械及設備", "交通及運輸設備", "什項設備", "租賃權益改良", "購建中固定資產", "核能燃料", "生產性植物", "使用權資產", "投資性不動產", "投資性不動產－土地", "投資性不動產－土地改良物", "投資性不動產－房屋及建築", "投資性不動產－租賃權益改良", "建造中之投資性不動產", "無形資產", "累計減損－無形資產", "生物資產", "消耗性生物資產－非流動", "生產性生物資產－非流動", "其他資產", "遞延資產", "遞延所得稅資產", "待整理資產", "什項資產", "合　　計" ];
const PUBLIC_LIABILITY_ORDER = [ "負債", "流動負債", "短期債務", "央行存款", "銀行同業存款", "國際金融機構存款", "應付款項", "本期所得稅負債", "發行券幣", "預收款項", "流動金融負債", "與待出售處分群組直接相關之負債", "合約負債－流動", "其他流動負債", "存款、匯款及金融債券", "支票存款", "活期存款", "定期存款", "儲蓄存款", "匯款", "金融債券", "央行及同業融資", "央行融資", "同業融資", "長期負債", "長期債務", "租賃負債", "非流動金融負債", "其他負債", "負債準備", "遞延負債", "遞延所得稅負債", "待整理負債", "合約負債－非流動", "什項負債", "權益", "資本", "預收資本", "資本公積", "保留盈餘（或累積虧損）", "已指撥保留盈餘", "未指撥保留盈餘", "累積虧損", "累積其他綜合損益", "國外營運機構財務報表換算之兌換差額", "現金流量避險中屬有效避險部分之避險工具利益（損失）", "國外營運機構淨投資避險中屬有效避險部分之避險工具利益（損失）", "不動產重估增值", "與待出售非流動資產直接相關之權益", "確定福利計畫之再衡量數", "指定為透過損益按公允價值衡量之金融負債其變動金額來自信用風險", "透過其他綜合損益按公允價值衡量之金融資產損益", "採用覆蓋法重分類之其他綜合損益", "其他權益－其他", "庫藏股票", "首次採用國際財務報導準則調整數", "非控制權益", "合　　計" ];

// --- Global Variables & UI Elements ---
const dropZone = document.getElementById('drop-zone'), fileInput = document.getElementById('file-input'), statusDiv = document.getElementById('status'), outputContainer = document.getElementById('output-container'), controlsContainer = document.querySelector('.controls'), typeSelector = document.getElementById('type-selector'), mainContent = document.getElementById('main-content'), resetButton = document.getElementById('reset-button'), fundTypeDisplay = document.getElementById('current-fund-type-display');
let allExtractedData = {}, fundNames = [], selectedFundType = null, fundFileMap = {};
typeSelector.addEventListener('change', e => { if (e.target.name === 'fund-type') { selectedFundType = e.target.value; const labelText = document.querySelector(`label[for=${e.target.id}]`).textContent; fundTypeDisplay.textContent = `目前選擇：${labelText}`; mainContent.style.display = 'block'; typeSelector.style.display = 'none'; } });
resetButton.addEventListener('click', () => resetState(true));
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', e => handleFiles(e.target.files));
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('drag-over'); handleFiles(e.dataTransfer.files); });

function handleFiles(files) {
    if (!selectedFundType) return alert('請先返回並選擇基金類型！');
    if (!files || files.length === 0) return;
    console.clear();
    resetState(false);
    statusDiv.textContent = `讀取中... 偵測到 ${files.length} 個檔案。`;
    const filePromises = Array.from(files).map(file => processFile(file, selectedFundType));
    Promise.all(filePromises).then(results => {
        const successfulResults = results.filter(r => r);
        successfulResults.forEach(r => fundFileMap[r.fundName] = r.fileName);
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
    }).catch(error => { statusDiv.textContent = `處理檔案時發生嚴重錯誤：${error.message}`; console.error(error); });
}
function resetState(fullReset = true) {
    allExtractedData = {}; fundNames = []; fundFileMap = {};
    outputContainer.innerHTML = ''; controlsContainer.innerHTML = '';
    statusDiv.textContent = ''; fileInput.value = '';
    if (fullReset) {
        fundTypeDisplay.textContent = ''; mainContent.style.display = 'none';
        typeSelector.style.display = 'block';
        document.querySelectorAll('input[name="fund-type"]').forEach(radio => radio.checked = false);
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
        const fundSelect = document.getElementById('fund-select');
        fundSelect.addEventListener('change', displayIndividualFund);
        if (fundNames.length > 0) displayIndividualFund({ target: { value: fundNames[0] } });
    } else if (selectedMode === 'sum') {
        displayAggregated();
    } else if (selectedMode === 'compare') {
        displayComparison();
    }
}

// 移除 buildTree 和 flattenTree 函數，避免引入階層計算的數值覆蓋問題。
// class Node 保持不變
class Node {
    constructor(name, indent, data = {}) { this.name = name.trim(); this.indent = indent; this.data = data; this.children = []; }
}
function buildTree(records, keyColumn, numericCols) {
    const root = new Node('Root', -1);
    const stack = [root];
    records.forEach(record => {
        const name = record[keyColumn]?.trim();
        if (!name) return;
        const indent = record.indent_level || 0;
        const data = { '基金名稱': record['基金名稱'] };
        numericCols.forEach(col => { data[col] = parseFloat(String(record[col] || '0').replace(/,/g, '')) || 0; });
        const node = new Node(name, indent, data);
        while (stack.length > 1 && stack[stack.length - 1].indent >= indent) { stack.pop(); }
        stack[stack.length - 1].children.push(node);
        stack.push(node);
    });
    return root;
}
function flattenTree(node, result, keyColumn) {
    if (node.name !== 'Root') {
        const row = { ...node.data, [keyColumn]: node.name, 'indent_level': node.indent, 'isSummary': node.children.length > 0 };
        result.push(row);
    }
    node.children.forEach(child => flattenTree(child, result, keyColumn));
}


function displayIndividualFund(e) {
    const selectedFund = e.target.value;
    if (!selectedFund) return;
    const fundData = {};
    for (const reportKey in allExtractedData) {
        const sourceReportData = allExtractedData[reportKey];
        if (sourceReportData) {
            const fundRecords = sourceReportData.filter(r => r['基金名稱'] === selectedFund);
            const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
            const config = FULL_CONFIG[selectedFundType]?.[baseKey];
            if (config && fundRecords.length > 0) {
                const keyColumn = config.keyColumn;
                const numericCols = config.columns.filter(c => c !== keyColumn && c !== '基金名稱');
                const tree = buildTree(fundRecords, keyColumn, numericCols);
                const flattened = [];
                flattenTree(tree, flattened, keyColumn);
                fundData[reportKey] = flattened;
            }
        }
    }
    outputContainer.innerHTML = createTabsAndTables(fundData, {}, 'individual');
    initTabs();
    initExportButtons();
}

function displayAggregated() {
    const summaryData = {};
    for (const reportKey in allExtractedData) {
        const sourceReportData = allExtractedData[reportKey];
        if (!sourceReportData || sourceReportData.length === 0) continue;
        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = FULL_CONFIG[selectedFundType]?.[baseKey];
        if (!config) continue;
        const keyColumn = config.keyColumn;
        const numericCols = config.columns.filter(c => c !== keyColumn && c !== '基金名稱');

        const otherFundsData = sourceReportData.filter(r => r['基金名稱'] !== '中央銀行');
        const centralBankData = sourceReportData.filter(r => r['基金名稱'] === '中央銀行');
        const ledger = new Map();

        // 1. 先加總所有「其他基金」
        otherFundsData.forEach(row => {
            const keyText = row[keyColumn]?.trim();
            if (!keyText) return;
            const indent = row.indent_level || 0;
            const compositeKey = `${keyText}::${indent}`;
            if (!ledger.has(compositeKey)) {
                ledger.set(compositeKey, { [keyColumn]: keyText, 'indent_level': indent, 'isSummary': false, ...Object.fromEntries(numericCols.map(c => [c, 0])) });
            }
            const summaryRow = ledger.get(compositeKey);
            numericCols.forEach(col => {
                const val = parseFloat(String(row[col] || '0').replace(/,/g, '')) || 0;
                summaryRow[col] += val;
            });
        });

        // Helper: 查找 ledger 中是否存在該名稱的科目，並返回其縮排
        const getExistingIndent = (name) => {
            const existingKey = [...ledger.keys()].find(k => k.startsWith(`${name}::`));
            return existingKey ? parseInt(existingKey.split('::')[1]) : null;
        };


        // 2. 處理「中央銀行」的數據，並應用例外規則
        const mergeRules = {
            '損益表': { '事業投資利益': '採用權益法認列之關聯企業及合資利益之份額', '事業投資損失': '採用權益法認列之關聯企業及合資損失之份額' },
            '資產負債表_資產': { '存放銀行業': '存放銀行同業', '事業投資': '採用權益法之投資', '其他長期投資': '採用權益法之投資' },
            '資產負債表_負債及權益': { '銀行業存款': '銀行同業存款', '存款': '存款、匯款及金融債券', '公庫及政府機關存款': '支票存款', '儲蓄存款及儲蓄券': '儲蓄存款' }
        };
        const summaryRuleSources = { '押匯貼現及放款': ['融通', '銀行業融通', '押匯及貼現', '短期放款及透支', '短期擔保放款及透支', '中期放款', '中期擔保放款', '長期放款', '長期擔保放款'] };
        const activeRules = (selectedFundType === 'business' && mergeRules[reportKey]) ? mergeRules[reportKey] : {};
        const sourceToTargetMap = new Map(Object.entries(activeRules));
        const centralBankOnlyItems = new Set();

        // 彙總型規則：記錄中央銀行的來源科目，稍後不顯示
        const summaryRuleTargets = new Map();
        for (const targetName in summaryRuleSources) {
            summaryRuleSources[targetName].forEach(sourceName => {
                sourceToTargetMap.set(sourceName, targetName);
                centralBankOnlyItems.add(sourceName);
            });
            summaryRuleTargets.set(targetName, summaryRuleSources[targetName]);
        }
                
        // 記錄所有被合併的中央銀行專屬科目
        Object.keys(activeRules).forEach(item => centralBankOnlyItems.add(item));
                
        centralBankData.forEach(row => {
            let keyText = row[keyColumn]?.trim();
            if (!keyText) return;
            let indent = row.indent_level || 0;
            const originalRowValue = Object.fromEntries(numericCols.map(col => [col, parseFloat(String(row[col] || '0').replace(/,/g, '')) || 0]));

            let targetName = keyText;
            let targetIndent = indent;
            
            // 檢查是否為需要合併的來源科目
            if (sourceToTargetMap.has(keyText)) {
                targetName = sourceToTargetMap.get(keyText);
                
                // 對於合併的目標科目，查找其他基金中是否存在，如果存在，則強制使用該縮排
                const existingTargetIndent = getExistingIndent(targetName); 
                
                if (existingTargetIndent !== null) {
                    targetIndent = existingTargetIndent;
                } else {
                    targetIndent = indent; // 如果沒有，就用中央銀行的原始縮排
                }
            } else {
                // 處理非合併的正常/父級科目 (關鍵修正點)
                const existingIndent = getExistingIndent(keyText);
                
                // 如果該科目已存在於 ledger (來自其他基金)，強制使用該縮排以避免數據分散
                if (existingIndent !== null) {
                    targetIndent = existingIndent;
                }
                // 否則，使用中央銀行數據的原始縮排 (targetIndent = indent)
            }

            // A. 處理合併後的目標科目 / 正常科目
            const compositeKey = `${targetName}::${targetIndent}`;
            if (!ledger.has(compositeKey)) {
                ledger.set(compositeKey, { [keyColumn]: targetName, 'indent_level': targetIndent, 'isSummary': false, ...Object.fromEntries(numericCols.map(c => [c, 0])) });
            }
            
            const summaryRow = ledger.get(compositeKey);
            // 累加數值到目標/正常科目。
            numericCols.forEach(col => {
                summaryRow[col] += originalRowValue[col];
            });

        });
        
        // 3. 準備最終輸出：從 ledger 轉換為可排序的列表
        // 由於移除了 buildTree/flattenTree，我們直接使用 ledger.values()
        let finalRowsForSort = Array.from(ledger.values());
        
        // 為了讓表格看起來像階層，這裡可以根據公版定義簡單標記 isSummary
        // 我們將所有縮排 < 2 的科目視為彙總父級
        finalRowsForSort.forEach(row => {
            if (row['indent_level'] < 2 || row[keyColumn] === '合　　計') {
                row.isSummary = true;
            }
        });

        // 4. 按公版順序排序
        const standardOrderMap = { '損益表': PROFIT_LOSS_ACCOUNT_ORDER, '盈虧撥補表': APPROPRIATION_ACCOUNT_ORDER, '現金流量表': CASH_FLOW_ACCOUNT_ORDER, '資產負債表_資產': PUBLIC_ASSET_ORDER, '資產負債表_負債及權益': PUBLIC_LIABILITY_ORDER };
        const order = standardOrderMap[reportKey] || [];
        const finalRows = [];
        const processedItems = new Set();
        // 隱藏的來源科目
        const sourceKeysToHide = centralBankOnlyItems;
        
        // 將 ledger 轉換為 {name: [row1, row2, ...]} 以便快速查找
        const ledgerByName = finalRowsForSort.reduce((acc, row) => {
            const name = row[keyColumn];
            if (!acc[name]) acc[name] = [];
            acc[name].push(row);
            return acc;
        }, {});


        // 從公版順序開始，逐一查找並添加科目
        order.forEach(accountName => {
            if (sourceKeysToHide.has(accountName)) return; 
            
            const matchingRows = ledgerByName[accountName] || [];
            
            if (matchingRows.length > 0) {
                 // 確保所有同名科目按縮排層級正確排序
                 matchingRows.sort((a,b) => a.indent_level - b.indent_level);
                 
                 matchingRows.forEach(row => {
                    const compositeKey = `${row[keyColumn]}::${row.indent_level}`;
                    if (!processedItems.has(compositeKey)) {
                       row['基金名稱'] = '所有基金加總';
                       if (selectedFundType === 'business') {
                           numericCols.forEach(col => row[col] = Math.round(row[col]));
                       }
                       finalRows.push(row);
                       processedItems.add(compositeKey);
                    }
                 });
            }
        });

        // 將不在公版中的項目也加回來 (如果公版不包含該科目，且它不是被隱藏的來源科目)
        finalRowsForSort.forEach(row => {
            const compositeKey = `${row[keyColumn]}::${row.indent_level}`;
            if (!processedItems.has(compositeKey) && !sourceKeysToHide.has(row[keyColumn])) {
                row['基金名稱'] = '所有基金加總';
                if (selectedFundType === 'business') {
                    numericCols.forEach(col => row[col] = Math.round(row[col]));
                }
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
        const items = [...new Set(allExtractedData[reportKey].map(r => r[keyColumn]))].filter(Boolean).sort((a, b) => a.localeCompare(b, 'zh-Hant'));
        const tabName = reportKey.replace(/_/g, ' ');
        if (items.length > 0) {
            optionsHtml += `<optgroup label="${tabName}">${items.map(item => `<option value="${reportKey}::${item}">${item}</option>`).join('')}</optgroup>`;
        }
    });
    dynamicControlsContainer.innerHTML = `<div class="control-group"><label for="item-select">選擇比較項目：</label><select id="item-select">${optionsHtml}</select></div>`;
    const itemSelect = document.getElementById('item-select');
    const updateComparisonView = () => {
        const selectedValue = itemSelect?.value;
        if (!selectedValue) { outputContainer.innerHTML = ''; return; };
        const [reportKey, itemName] = selectedValue.split('::');
        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = FULL_CONFIG[selectedFundType][baseKey];
        const keyColumn = config.keyColumn;
        const columns = config.columns;
        const dataForCompare = allExtractedData[reportKey].filter(r => r[keyColumn] === itemName);
        const finalData = { [reportKey]: [] };
        const totals = {};
        const numericHeaders = columns.filter(c => c !== keyColumn && c !== '基金名稱');
        numericHeaders.forEach(h => totals[h] = 0);
        fundNames.forEach(fund => {
            const fundRow = dataForCompare.find(r => r['基金名稱'] === fund);
            const newRow = { '基金名稱': fund };
            columns.forEach(col => {
                if (col !== keyColumn) {
                    const val = fundRow ? fundRow[col] : '-';
                    newRow[col] = val;
                    const numVal = parseFloat(String(val).replace(/,/g, '')) || 0;
                    if (totals[col] !== undefined) totals[col] += numVal;
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
    if (itemSelect.options.length > 0) updateComparisonView();
}


function createTabsAndTables(data, customHeaders = {}, mode = 'default') {
    let tabsHtml = '<div class="report-tabs">';
    let contentHtml = '';
    let isFirst = true;
    const activeConfig = FULL_CONFIG[selectedFundType];
    const allPossibleKeys = Object.keys(activeConfig).flatMap(k => 
        (k === '平衡表' || k === '資產負債表') ? [`${k}_資產`, `${k}_負債及權益`] : [k]
    );
    const orderedDataKeys = allPossibleKeys.filter(key => data[key] && data[key].length > 0);

    orderedDataKeys.forEach(reportKey => {
        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = activeConfig[baseKey];
        if (!config) return;
        const columns = config.columns;
        const tabName = reportKey.replace(/_/g, ' ');
        tabsHtml += `<button class="tab-link ${isFirst ? 'active' : ''}" data-tab="${reportKey}">${tabName}</button>`;
        const headers = customHeaders[reportKey] || (mode === 'individual' ? columns : ['基金名稱', ...columns]);
        let tableContent = createTableHtml(data[reportKey], headers, mode);
        const exportButtons = `<div class="export-buttons"><button class="export-btn json" data-format="json" data-report-key="${reportKey}">匯出 JSON</button><button class="export-btn xlsx" data-format="xlsx" data-report-key="${reportKey}">匯出 XLSX</button><button class="export-btn" data-format="html" data-report-key="${reportKey}">匯出 HTML</button></div>`;
        contentHtml += `<div id="${reportKey}" class="tab-content ${isFirst ? 'active' : ''}"><div class="tab-header"><h2>${tabName}</h2>${exportButtons}</div>${tableContent}</div>`;
        isFirst = false;
    });
    tabsHtml += '</div>';
    if (contentHtml === '') return '<p>無資料可顯示。</p>';
    return tabsHtml + contentHtml;
}

function createTableHtml(records, headers, mode = 'default') {
    let table = '<table><thead><tr>';
    const keyColumns = ['科目', '項目', '基金名稱'];
    table += headers.map(h => `<th data-column-key="${h}">${h}</th>`).join('');
    table += '</tr></thead><tbody>';

    records.forEach(record => {
        table += '<tr>';
        headers.forEach(header => {
            const val = record[header];
            let displayVal = (val === null || val === undefined) ? '' : val;
            const isKeyColumn = keyColumns.includes(header);
            const indentLevel = record.indent_level || 0;
            let style = '';
            let className = '';
            const isSummaryRow = record.isSummary === true;

            if (isKeyColumn) {
                if (header !== '基金名稱' && indentLevel > 0) {
                    style = `padding-left: ${1 + indentLevel * 1.5}em;`;
                }
                if (isSummaryRow) {
                    displayVal = `<strong>${displayVal}</strong>`;
                }
            } else {
                const rawVal = String(val).replace(/,/g, '');
                const isNumericField = !isNaN(parseFloat(rawVal)) && isFinite(rawVal) && rawVal.trim() !== '';
                if (isNumericField) {
                    const numVal = Number(rawVal);
                    displayVal = numVal.toLocaleString();
                    if (isSummaryRow) {
                        displayVal = `<strong>${displayVal}</strong>`;
                    }
                    className = 'numeric-data';
                    if (numVal < 0) { className += ' negative-value'; }
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
        tab.addEventListener('click', e => {
            tabs.forEach(t => t.classList.remove('active'));
            contents.forEach(c => c.classList.remove('active'));
            const tabId = e.currentTarget.dataset.tab;
            e.currentTarget.classList.add('active');
            document.getElementById(tabId).classList.add('active');
        });
    });
    if (tabs.length > 0) {
        const firstActiveTab = document.querySelector('.report-tabs .tab-link.active');
        if (!firstActiveTab && contents.length > 0) {
            tabs[0].classList.add('active');
            contents[0].classList.add('active');
        }
    }
}
function initSortableTables() {
    document.querySelectorAll('th.sortable').forEach(th => {
        th.addEventListener('click', () => {
            const table = th.closest('table');
            const tbody = table.querySelector('tbody');
            const headerIndex = Array.from(th.parentNode.children).indexOf(th);
            const currentIsDesc = th.classList.contains('sort-desc');
            table.querySelectorAll('th').forEach(h => {
                h.classList.remove('sort-asc', 'sort-desc');
                h.querySelector('.sort-arrow')?.remove();
            });
            const direction = currentIsDesc ? 'asc' : 'desc';
            th.classList.add(direction === 'asc' ? 'sort-asc' : 'sort-desc');
            th.innerHTML += `<span class="sort-arrow">${direction === 'asc' ? '▲' : '▼'}</span>`;
            const rows = Array.from(tbody.querySelectorAll('tr'));
            const headerRow = rows.shift();
            rows.sort((a, b) => {
                const aVal = parseFloat(String(a.children[headerIndex].textContent).replace(/,/g, '')) || Number.NEGATIVE_INFINITY;
                const bVal = parseFloat(String(b.children[headerIndex].textContent).replace(/,/g, '')) || Number.NEGATIVE_INFINITY;
                return direction === 'asc' ? aVal - bVal : bVal - aVal;
            }).forEach(row => tbody.appendChild(row));
            tbody.insertBefore(headerRow, tbody.firstChild);
        });
    });
}
function initExportButtons() {
    document.querySelectorAll('.export-btn').forEach(button => {
        button.addEventListener('click', e => {
            const reportKey = e.target.dataset.reportKey;
            const format = e.target.dataset.format;
            exportData(reportKey, format);
        });
    });
}
function createGovernmentalYuchuSummaryTable(aggregatedData) {
    // This function can be filled in if needed
}
