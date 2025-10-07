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
