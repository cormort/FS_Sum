// config.js: 存放所有靜態設定資料

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

export const BUSINESS_PROFIT_LOSS_ACCOUNTS = [
    { "code": "4", "name": "收入", "type": "calc" }, 
    { "code": "41", "name": "營業收入", "type": "calc" }, 
    { "code": "4101", "name": "銷售收入" },
    { "code": "4102", "name": "勞務收入" }, 
    { "code": "4103", "name": "金融保險收入" },
    { "code": "410601", "name": "採用權益法認列之關聯企業及合資利益之份額" },
    { "code": "4107", "name": "採用權益法認列之子公司、關聯企業及合資利益之份額", "merge_target": "採用權益法認列之關聯企業及合資利益之份額" },
    { "code": "4198", "name": "其他營業收入" }, 
    { "code": "4199", "name": "內部損益" }, 
    { "code": "49", "name": "營業外收入", "type": "calc" },
    { "code": "4901", "name": "財務收入" }, 
    { "code": "4904", "name": "採用權益法認列之關聯企業及合資利益之份額" },
    { "code": "4905", "name": "採用權益法認列之子公司、關聯企業及合資利益之份額", "merge_target": "採用權益法認列之關聯企業及合資利益之份額" }, 
    { "code": "4998", "name": "其他營業外收入" },
    { "code": "5", "name": "支出", "type": "calc" }, 
    { "code": "51", "name": "營業成本", "type": "calc" }, 
    { "code": "5101", "name": "銷售成本" },
    { "code": "5102", "name": "勞務成本" }, 
    { "code": "5103", "name": "金融保險成本" },
    { "code": "5106", "name": "採用權益法認列之關聯企業及合資損失之份額" },
    { "code": "5107", "name": "採用權益法認列之子公司、關聯企業及合資損失之份額", "merge_target": "採用權益法認列之關聯企業及合資損失之份額" },
    { "code": "5198", "name": "其他營業成本" }, 
    { "code": "5199", "name": "內部損益" }, 
    { "code": "52", "name": "營業費用", "type": "calc" },
    { "code": "5201", "name": "行銷費用" }, 
    { "code": "5202", "name": "業務費用" }, 
    { "code": "5203", "name": "管理費用" },
    { "code": "5298", "name": "其他營業費用" }, 
    { "code": "59", "name": "營業外費用", "type": "calc" }, 
    { "code": "5901", "name": "財務成本" },
    { "code": "5904", "name": "採用權益法認列之關聯企業及合資損失之份額" },
    { "code": "5905", "name": "採用權益法認列之子公司、關聯企業及合資損失之份額", "merge_target": "採用權益法認列之關聯企業及合資損失之份額" },
    { "code": "5998", "name": "其他營業外費用" }, 
    { "code": "61", "name": "營業毛利(毛損)", "type": "calc" },
    { "code": "62", "name": "營業利益(損失)", "type": "calc" },
    { "code": "63", "name": "營業外利益(損失)", "type": "calc" },
    { "code": "64", "name": "稅前淨利(淨損)", "type": "calc" },
    { "code": "65", "name": "所得稅費用(利益)" },
    { "code": "66", "name": "繼續營業單位本期淨利(淨損)", "type": "calc" },
    { "code": "67", "name": "停業單位損益" },
    { "code": "68", "name": "本期淨利(淨損)", "type": "calc" },
    { "code": "69", "name": "本期淨利(淨損)歸屬於:", "type": "calc" },
    { "code": "6901", "name": "母公司業主" },
    { "code": "6902", "name": "非控制權益" }
];