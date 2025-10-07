// config.js

export const CONFIG_OPERATING = {
    '收支表': { parser: 'dynamic_normal', sheetKeyword: '收支', keyColumn: '科目', columns: ['科目', '預算數', '原列決算數', '修正數', '決算核定數', '預算數與決算核定數比較增減'] },
    '撥補表': { parser: 'dynamic_normal', sheetKeyword: '撥補', keyColumn: '項目', columns: ['項目', '預算數', '原列決算數', '修正數', '決算核定數', '預算數與決算核定數比較增減'] },
    '現流表': { parser: 'dynamic_normal', sheetKeyword: '現流', keyColumn: '項目', columns: ['項目', '預算數', '決算核定數', '比較增減'] },
    '平衡表': { parser: 'balance_sheet', sheetKeyword: '平衡', keyColumn: '科目', columns: ['科目', '本年度決算核定數', '上年度決算審定數', '比較增減'] }
};

export const CONFIG_GOVERNMENTAL = {
    '餘絀表': { parser: 'fixed_yuchu', sheetKeyword: '餘絀', keyColumn: '項目', columns: ['項目', '預算數', '原列決算數', '修正數', '決算核定數', '預算與決算核定數比較增減'] },
    '收支表': { parser: 'fixed_shouzhi', sheetKeyword: '收支', keyColumn: '科目', columns: ['科目', '原列決算數', '修正數', '決算核定數'] },
    '現流表': { parser: 'fixed_xianliu', sheetKeyword: '現流', keyColumn: '項目', columns: ['項目', '決算核定數'] },
    '平衡表': { parser: 'balance_sheet', sheetKeyword: '平衡', keyColumn: '科目', columns: ['科目', '本年度決算核定數', '上年度決算審定數', '比較增減'] }
};

export const CONFIG_BUSINESS = {
    '損益表': { parser: 'fixed_business_profitloss', sheetKeyword: '損益', keyColumn: '科目', columns: ['科目', '上年度決算數', '本年度預算數', '原列決算數', '修正數', '決算核定數'] },
    '盈虧撥補表': { parser: 'fixed_business_appropriation', sheetKeyword: '盈虧', keyColumn: '項目', columns: ['項目', '上年度決算數', '本年度預算數', '原列決算數', '修正數', '決算核定數'] },
    '現金流量表': { parser: 'fixed_business_cashflow', sheetKeyword: '現流', keyColumn: '項目', columns: ['項目', '本年度預算數', '原列決算數', '修正數', '決算核定數'] },
    '資產負債表': { parser: 'fixed_business_balancesheet', sheetKeyword: '資負', keyColumn: '科目', columns: ['科目', '上年度決算數', '原列決算數', '修正數', '決算核定數'] }
};

export const FULL_CONFIG = {
    'operating': CONFIG_OPERATING,
    'governmental': CONFIG_GOVERNMENTAL,
    'business': CONFIG_BUSINESS
};

// ★★★ 新增：標準化損益表範本 for "所有基金加總" 模式 ★★★
// type: 'data' (從原始資料抓取), 'calc' (計算得出), 'header' (僅標題)
// aliases: 用於對應不同 Excel 檔案中可能的名稱差異
export const PROFIT_LOSS_SUMMARY_TEMPLATE = [
    { name: '營業收入', indent: 0, type: 'calc' },
    { name: '銷售收入', indent: 1, type: 'data', aliases: ['銷貨收入'] },
    { name: '勞務收入', indent: 1, type: 'data' },
    { name: '金融保險收入', indent: 1, type: 'data' },
    { name: '採用權益法認列之關聯企業及合資利益之份額', indent: 1, type: 'data', special: 'merge_investment_profit' },
    { name: '其他營業收入', indent: 1, type: 'data' },
    { name: '營業成本', indent: 0, type: 'calc' },
    { name: '銷售成本', indent: 1, type: 'data', aliases: ['銷貨成本'] },
    { name: '勞務成本', indent: 1, type: 'data' },
    { name: '金融保險成本', indent: 1, type: 'data' },
    { name: '採用權益法認列之關聯企業及合資損失之份額', indent: 1, type: 'data', special: 'merge_investment_loss' },
    { name: '其他營業成本', indent: 1, type: 'data' },
    { name: '營業毛利（毛損）', indent: 0, type: 'calc', formula: ['營業收入', '-', '營業成本'] },
    { name: '營業費用', indent: 0, type: 'calc' },
    { name: '行銷費用', indent: 1, type: 'data' },
    { name: '業務費用', indent: 1, type: 'data' },
    { name: '管理費用', indent: 1, type: 'data' },
    { name: '其他營業費用', indent: 1, type: 'data' },
    { name: '營業利益（損失）', indent: 0, type: 'calc', formula: ['營業毛利（毛損）', '-', '營業費用'] },
    { name: '營業外收入', indent: 0, type: 'calc' },
    { name: '財務收入', indent: 1, type: 'data' },
    { name: '採用權益法認列之關聯企業及合資利益之份額 (營業外)', indent: 1, type: 'data', source_name: '採用權益法認列之關聯企業及合資利益之份額', note: 'non-operating' },
    { name: '其他營業外收入', indent: 1, type: 'data' },
    { name: '營業外費用', indent: 0, type: 'calc' },
    { name: '財務成本', indent: 1, type: 'data' },
    { name: '採用權益法認列之關聯企業及合資損失之份額 (營業外)', indent: 1, type: 'data', source_name: '採用權益法認列之關聯企業及合資損失之份額', note: 'non-operating' },
    { name: '其他營業外費用', indent: 1, type: 'data' },
    { name: '營業外利益（損失）', indent: 0, type: 'calc', formula: ['營業外收入', '-', '營業外費用'] },
    { name: '稅前淨利（淨損）', indent: 0, type: 'calc', formula: ['營業利益（損失）', '+', '營業外利益（損失）'] },
    { name: '所得稅費用（利益）', indent: 0, type: 'data' },
    { name: '繼續營業單位本期淨利（淨損）', indent: 0, type: 'calc', formula: ['稅前淨利（淨損）', '-', '所得稅費用（利益）'] },
    { name: '停業單位損益', indent: 0, type: 'data' },
    { name: '本期淨利（淨損）', indent: 0, type: 'calc', formula: ['繼續營業單位本期淨利（淨損）', '+', '停業單位損益'] },
    { name: '本期淨利（淨損）歸屬於：', indent: 0, type: 'header' },
    { name: '母公司業主', indent: 1, type: 'calc', formula: ['本期淨利（淨損）', '-', '非控制權益'] },
    { name: '非控制權益', indent: 1, type: 'data' },
];
