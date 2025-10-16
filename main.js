// main.js

import { FULL_CONFIG, PROFIT_LOSS_ACCOUNT_ORDER, APPROPRIATION_ACCOUNT_ORDER, CASH_FLOW_ACCOUNT_ORDER } from './config.js';
import { processFile } from './parsers.js';
import { exportData } from './utils.js';

// --- 公版順序樣板 ---
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
        const fundSelect = document.getElementById('fund-select');
        fundSelect.addEventListener('change', displayIndividualFund);
        if (fundNames.length > 0) {
           displayIndividualFund({ target: { value: fundNames[0] } });
        }
    } else if (selectedMode === 'sum') {
        displayAggregated();
    } else if (selectedMode === 'compare') {
        displayComparison();
    }
}

// --- 階層樹與格式化核心邏輯 ---

class Node {
    constructor(name, indent, data = {}) {
        this.name = name.trim();
        this.indent = indent;
        this.data = data;
        this.children = [];
    }
}

function buildTree(records, keyColumn, numericCols) {
    const root = new Node('Root', -1);
    const stack = [root];
    records.forEach(record => {
        const name = record[keyColumn]?.trim();
        if (!name) return;
        const indent = record.indent_level || 0;
        const data = {};
        numericCols.forEach(col => {
            data[col] = parseFloat(String(record[col] || '0').replace(/,/g, '')) || 0;
        });
        // 把基金名稱也存入 data，以便後續使用
        if (record['基金名稱']) {
            data['基金名稱'] = record['基金名稱'];
        }
        const node = new Node(name, indent, data);
        while (stack.length > 1 && stack[stack.length - 1].indent >= indent) {
            stack.pop();
        }
        stack[stack.length - 1].children.push(node);
        stack.push(node);
    });
    return root;
}

function mergeTrees(targetNode, sourceNode, numericCols) {
    sourceNode.children.forEach(sourceChild => {
        let found = false;
        for (const targetChild of targetNode.children) {
            if (targetChild.name === sourceChild.name && targetChild.indent === sourceChild.indent) {
                numericCols.forEach(col => {
                    targetChild.data[col] = (targetChild.data[col] || 0) + (sourceChild.data[col] || 0);
                });
                mergeTrees(targetChild, sourceChild, numericCols);
                found = true;
                break;
            }
        }
        if (!found) {
            targetNode.children.push(JSON.parse(JSON.stringify(sourceChild)));
        }
    });
}

function recalculateTotals(node, numericCols) {
    if (!node.children || node.children.length === 0) return;
    node.children.forEach(child => recalculateTotals(child, numericCols));
    const isSummaryNode = node.children.some(child => child.data[numericCols[0]] !== undefined);
    if (isSummaryNode) {
        numericCols.forEach(col => {
            node.data[col] = 0;
            node.children.forEach(child => {
                node.data[col] += (child.data[col] || 0);
            });
        });
    }
}

function flattenTree(node, result, keyColumn, numericCols) {
    if (node.name !== 'Root') {
        const row = {
            ...node.data, // ★ 核心修改：將節點的所有 data (包含基金名稱) 展開到 row 中
            [keyColumn]: node.name,
            'indent_level': node.indent,
            'isSummary': node.children.length > 0
        };
        // numericCols.forEach(col => { row[col] = node.data[col] || 0; });
        result.push(row);
    }
    node.children.forEach(child => flattenTree(child, result, keyColumn, numericCols));
}

// --- 顯示模式切換函式 ---

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
            if(config && fundRecords.length > 0) {
                const keyColumn = config.keyColumn;
                const numericCols = config.columns.filter(c => c !== keyColumn && c !== '基金名稱');
                const tree = buildTree(fundRecords, keyColumn, numericCols);
                const flattened = [];
                flattenTree(tree, flattened, keyColumn, numericCols);
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
        
        const fundTrees = fundNames.map(name => {
            const fundRecords = sourceReportData.filter(r => r['基金名稱'] === name);
            return buildTree(fundRecords, keyColumn, numericCols);
        });

        const summaryTree = buildTree([], keyColumn, numericCols);
        fundTrees.forEach(tree => mergeTrees(summaryTree, tree, numericCols));
        summaryTree.data['基金名稱'] = '所有基金加總';

        const mergeRules = {
            '損益表': [ { target: '採用權益法認列之關聯企業及合資利益之份額', sources: ['事業投資利益'], type: 'additive' }, { target: '採用權益法認列之關聯企業及合資損失之份額', sources: ['事業投資損失'], type: 'additive' } ],
            '資產負債表_資產': [ { target: '存放銀行同業', sources: ['存放銀行業'], type: 'additive' }, { target: '押匯貼現及放款', sources: ['融通', '銀行業融通', '押匯及貼現', '短期放款及透支', '短期擔保放款及透支', '中期放款', '中期擔保放款', '長期放款', '長期擔保放款'], type: 'summary' }, { target: '採用權益法之投資', sources: ['事業投資', '其他長期投資'], type: 'additive' } ],
            '資產負債表_負債及權益': [ { target: '銀行同業存款', sources: ['銀行業存款'], type: 'additive' }, { target: '存款、匯款及金融債券', sources: ['存款'], type: 'additive' }, { target: '支票存款', sources: ['公庫及政府機關存款'], type: 'additive' }, { target: '儲蓄存款', sources: ['儲蓄存款及儲蓄券'], type: 'additive' } ]
        };
        const activeMergeRules = (selectedFundType === 'business' && mergeRules[reportKey]) ? mergeRules[reportKey] : [];
        const sourceKeysToHide = new Set(activeMergeRules.flatMap(rule => rule.sources));
        const centralBankData = sourceReportData.filter(r => r['基金名稱'] === '中央銀行');
        
        activeMergeRules.forEach(rule => {
             function findNode(node, name) {
                 for(const child of node.children) {
                     if(child.name === name) return child;
                     const found = findNode(child, name);
                     if(found) return found;
                 }
                 return null;
             }
             const targetNode = findNode(summaryTree, rule.target);
             if(!targetNode) return;
             
             const sourceValues = {};
             numericCols.forEach(col => sourceValues[col] = 0);
             rule.sources.forEach(sourceKey => {
                 centralBankData.forEach(row => {
                     if (row[keyColumn]?.trim() === sourceKey) {
                         numericCols.forEach(col => {
                             sourceValues[col] += parseFloat(String(row[col] || '0').replace(/,/g, '')) || 0;
                         });
                     }
                 });
             });
             
             if (rule.type === 'additive') {
                 numericCols.forEach(col => targetNode.data[col] += sourceValues[col]);
             } else if (rule.type === 'summary') {
                 const cbTargetContribution = {};
                 numericCols.forEach(col => cbTargetContribution[col] = 0);
                 centralBankData.forEach(row => {
                     if (row[keyColumn]?.trim() === rule.target && (row.indent_level || 0) === targetNode.indent) {
                          numericCols.forEach(col => {
                            cbTargetContribution[col] += parseFloat(String(row[col] || '0').replace(/,/g, '')) || 0;
                          });
                     }
                 });
                 numericCols.forEach(col => {
                    targetNode.data[col] = (targetNode.data[col] - cbTargetContribution[col]) + sourceValues[col];
                 });
             }
        });
        
        recalculateTotals(summaryTree, numericCols);

        const flattenedData = [];
        flattenTree(summaryTree, flattenedData, keyColumn, numericCols);

        const standardOrderMap = {
             '損益表': PROFIT_LOSS_ACCOUNT_ORDER,
             '盈虧撥補表': APPROPRIATION_ACCOUNT_ORDER,
             '現金流量表': CASH_FLOW_ACCOUNT_ORDER,
             '資產負債表_資產': PUBLIC_ASSET_ORDER,
             '資產負債表_負債及權益': PUBLIC_LIABILITY_ORDER
        };
        const order = standardOrderMap[reportKey] || [];
        
        const finalRows = [];
        const processedItems = new Set();

        order.forEach(accountName => {
            if (sourceKeysToHide.has(accountName)) return;
            const matchingRows = flattenedData.filter(row => row[keyColumn] === accountName);
            if (matchingRows.length > 0) {
                 matchingRows.sort((a,b) => a.indent_level - b.indent_level);
                 matchingRows.forEach(row => {
                    row['基金名稱'] = '所有基金加總'; // 確保加總模式下有名稱
                    if (selectedFundType === 'business') {
                        numericCols.forEach(col => row[col] = Math.round(row[col]));
                    }
                    finalRows.push(row);
                    processedItems.add(`${row[keyColumn]}::${row.indent_level}`);
                 })
            }
        });
        
        flattenedData.forEach(row => {
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

function displayComparison() { /* ... 此函式保持不變 ... */ }

// --- HTML渲染與UI互動函式 ---

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

    const dataKeys = Object.keys(data).filter(key => data[key] && data[key].length > 0);
    const orderedDataKeys = allPossibleKeys.filter(k => dataKeys.includes(k));

    orderedDataKeys.forEach(reportKey => {
        const baseKey = reportKey.replace(/_資產|_負債及權益/, '');
        const config = activeConfig[baseKey];
        if (!config) return;

        const columns = config.columns;
        const tabName = reportKey.replace(/_/g, ' ');
        tabsHtml += `<button class="tab-link ${isFirst ? 'active' : ''}" data-tab="${reportKey}">${tabName}</button>`;
        
        const headers = customHeaders[reportKey] || (mode === 'individual' ? columns : ['基金名稱', ...columns]);
        let tableContent = createTableHtml(data[reportKey], headers, mode);
        
        const exportButtons = `
            <div class="export-buttons">
                <button class="export-btn json" data-format="json" data-report-key="${reportKey}">匯出 JSON</button>
                <button class="export-btn xlsx" data-format="xlsx" data-report-key="${reportKey}">匯出 XLSX</button>
                <button class="export-btn" data-format="html" data-report-key="${reportKey}">匯出 HTML</button>
            </div>`;
        contentHtml += `<div id="${reportKey}" class="tab-content ${isFirst ? 'active' : ''}">
            <div class="tab-header"><h2>${tabName}</h2>${exportButtons}</div>${tableContent}
        </div>`;
        isFirst = false;
    });
    tabsHtml += '</div>';
    if (contentHtml === '') { return '<p>無資料可顯示。</p>'; }
    return tabsHtml + contentHtml;
}

// ★★★ 核心修改：修正加粗邏輯和欄位偏移問題 ★★★
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
                const isNumericField = val != null && val !== '' && !isNaN(Number(rawVal)) && isFinite(Number(rawVal));
                
                if (isNumericField) {
                    const numVal = Number(rawVal);
                    displayVal = numVal.toLocaleString();
                    if (isSummaryRow) {
                        displayVal = `<strong>${displayVal}</strong>`;
                    }
                    className = 'numeric-data';
                    if (numVal < 0) {
                        className += ' negative-value';
                    }
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

function initSortableTables() { /* Can be implemented if needed */ }

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
    // This function remains unchanged 
}
