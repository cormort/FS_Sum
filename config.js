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
    '損益表': { parser: 'fixed_business_profitloss_stateful', sheetKeyword: '損益', keyColumn: '科目', columns: ['科目', '上年度決算數', '本年度預算數', '原列決算數', '修正數', '決算核定數'] },
    '盈虧撥補表': { parser: 'fixed_business_appropriation', sheetKeyword: '盈虧', keyColumn: '項目', columns: ['項目', '上年度決算數', '本年度預算數', '原列決算數', '修正數', '決算核定數'] },
    '現金流量表': { parser: 'fixed_business_cashflow', sheetKeyword: '現流', keyColumn: '項目', columns: ['項目', '本年度預算數', '原列決算數', '修正數', '決算核定數'] },
    '資產負債表': { parser: 'fixed_business_balancesheet', sheetKeyword: '資負', keyColumn: '科目', columns: ['科目', '上年度決算數', '原列決算數', '修正數', '決算核定數'] }
};

export const FULL_CONFIG = {
    'operating': CONFIG_OPERATING,
    'governmental': CONFIG_GOVERNMENTAL,
    'business': CONFIG_BUSINESS
};

// ★★★ 確保您的檔案中有這一段 export ★★★
export const PROFIT_LOSS_ACCOUNT_ORDER = [
    '營業收入',
    '銷售收入',
    '勞務收入',
    '金融保險收入',
    '採用權益法認列之關聯企業及合資利益之份額',
    '其他營業收入',
    '營業成本',
    '銷售成本',
    '勞務成本',
    '金融保險成本',
    '採用權益法認列之關聯企業及合資損失之份額',
    '其他營業成本',
    '營業毛利（毛損）',
    '營業費用',
    '行銷費用',
    '業務費用',
    '管理費用',
    '其他營業費用',
    '營業利益（損失）',
    '營業外收入',
    '財務收入',
    '採用權益法認列之關聯企業及合資利益之份額 (營業外)',
    '其他營業外收入',
    '營業外費用',
    '財務成本',
    '採用權益法認列之關聯企業及合資損失之份額 (營業外)',
    '其他營業外費用',
    '營業外利益（損失）',
    '稅前淨利（淨損）',
    '所得稅費用（利益）',
    '繼續營業單位本期淨利（淨損）',
    '停業單位損益',
    '本期淨利（淨損）',
    '本期淨利（淨損）歸屬於：',
    '母公司業主',
    '非控制權益'
];
