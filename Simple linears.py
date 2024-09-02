import pandas as pd
import statsmodels.api as sm

#---------------------- 表格处理 ----------------------
macro_data = pd.read_excel('res/data.xlsx', sheet_name='Macro')
excel_file = 'res/data.xlsx'
sheet_names = pd.ExcelFile(excel_file).sheet_names

significance_level = 0.05

#---------------------- 数据处理 ----------------------
for sheet in sheet_names:
    if sheet != 'Macro':
        print(f"\n--- 同{sheet}的线性回归结果是 ---")
        market_data = pd.read_excel(excel_file, sheet_name=sheet)
        merged_data = pd.merge(macro_data, market_data, on='Date')
        Y = merged_data['Rt']  # 因变量
        X_vars = merged_data.columns.difference(['Date', 'Rt']) # 自变量
        results = []

        # 分别简单线性回归
        for column in X_vars:
            X = merged_data[[column]]
            X = sm.add_constant(X) #添加常数项
            model = sm.OLS(Y, X).fit()
            significant = '显著' if model.pvalues[column] < significance_level else '不显著'
            results.append({
                'Variable': column,
                'Coef': model.params[column],
                'P-value': model.pvalues[column],
                'R-squared': model.rsquared,
                'Significant': significant
            })

        results_df = pd.DataFrame(results)
        print(results_df)