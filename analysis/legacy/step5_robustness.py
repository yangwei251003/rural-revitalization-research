"""
Step5: 多维度稳健性检验
========================
检验方法：
1. 替换被解释变量：用等权重TOPSIS得分替代熵权TOPSIS得分
2. 缩尾处理：对极端值进行1%和99%分位数缩尾
3. 子样本回归：剔除直辖市/剔除疫情年份(2020-2022)
4. 替换模型：随机效应模型 vs 固定效应模型（Hausman检验）
5. 滞后一期回归：解释变量滞后一期，缓解内生性
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from linearmodels.panel import PanelOLS, RandomEffects, compare
from sklearn.preprocessing import StandardScaler
from scipy.stats import chi2 as stats_chi2
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ========== 指标名称映射 ==========
var_names = {
    'X01_农村居民人均可支配收入(元)': 'X01_人均可支配收入',
    'X02_农林牧渔业总产值(亿元)': 'X02_农林牧渔产值',
    'X03_农业机械总动力(万千瓦)': 'X03_农机总动力',
    'X04_有效灌溉面积占比(%)': 'X04_有效灌溉占比',
    'X05_化肥施用强度(万吨/千公顷)': 'X05_化肥施用强度',
    'X06_农药施用强度(万吨/千公顷)': 'X06_农药施用强度',
    'X07_文教娱乐支出占消费比重(%)': 'X07_文教娱乐占比',
    'X08_每万乡村人口卫生技术人员数(人)': 'X08_卫生技术人员',
    'X09_公共图书馆数量(个)': 'X09_公共图书馆数',
    'X11_森林覆盖率(%)': 'X11_森林覆盖率',
    'X14_农林水支出占财政支出比重(%)': 'X14_农林水支出占比',
    'X15_村民委员会数量(个)': 'X15_村委会数量',
    'X16_城乡居民养老保险参保人数(万人)': 'X16_养老保险参保',
    'X17_城乡居民医疗保险参保人数(万人)': 'X17_医疗保险参保'
}

# ========== 辅助函数 ==========
def winsorize(series, lower=0.01, upper=0.99):
    """缩尾处理"""
    q_low = series.quantile(lower)
    q_high = series.quantile(upper)
    return series.clip(lower=q_low, upper=q_high)

def run_panel_ols(df, y_col, x_cols, model_name=""):
    """运行双向固定效应面板回归，返回结果字典"""
    y = df[y_col]
    X = df[x_cols]
    
    model = PanelOLS(y, X, entity_effects=True, time_effects=True, drop_absorbed=True)
    results = model.fit(cov_type='clustered', cluster_entity=True)
    
    # 收集结果
    res_dict = {
        '模型': model_name,
        'R²': results.rsquared,
        'nobs': results.nobs,
        'F_stat': results.f_statistic.stat,
    }
    
    for var in results.params.index:
        res_dict[f'{var}_coef'] = results.params[var]
        res_dict[f'{var}_pval'] = results.pvalues[var]
        res_dict[f'{var}_sig'] = '***' if results.pvalues[var] < 0.01 else '**' if results.pvalues[var] < 0.05 else '*' if results.pvalues[var] < 0.1 else ''
    
    return results, res_dict

# ========== 主程序 ==========
print("="*60)
print("Step5: 多维度稳健性检验")
print("="*60)

# 1. 读取数据
print("\n[1/6] 读取数据...")
df_panel = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')
df_topsis = pd.read_excel('Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')

# 合并TOPSIS得分
df = pd.merge(df_panel, df_topsis[['省份', '年份', 'TOPSIS综合得分']], 
              on=['省份', '年份'], how='inner')

# 简化列名
df.rename(columns=var_names, inplace=True)

# 获取所有解释变量（剔除被吸收的X11_森林覆盖率）
X_vars = [v for v in var_names.values() if v != 'X11_森林覆盖率']
print(f"✓ 解释变量({len(X_vars)}个，剔除X11_森林覆盖率)")

# 标准化
scaler = StandardScaler()
df_std = df.copy()
df_std = df_std.set_index(['省份', '年份'])
df_std[X_vars] = scaler.fit_transform(df_std[X_vars])

# 基准模型
print("\n[2/6] 基准模型（双向固定效应）...")
base_results, base_dict = run_panel_ols(df_std, 'TOPSIS综合得分', X_vars, '基准模型')
print(f"  基准模型 R² = {base_results.rsquared:.4f}")

# 2. 稳健性检验1：缩尾处理
print("\n[3/6] 稳健性检验1：缩尾处理...")
df_wins = df.copy()
for var in X_vars:
    df_wins[var] = winsorize(df_wins[var])

df_wins_std = df_wins.copy()
df_wins_std = df_wins_std.set_index(['省份', '年份'])
scaler_wins = StandardScaler()
df_wins_std[X_vars] = scaler_wins.fit_transform(df_wins_std[X_vars])

wins_results, wins_dict = run_panel_ols(df_wins_std, 'TOPSIS综合得分', X_vars, '缩尾处理')
print(f"  缩尾模型 R² = {wins_results.rsquared:.4f}")

# 3. 稳健性检验2：剔除直辖市
print("\n[4/6] 稳健性检验2：剔除直辖市...")
municipalities = ['北京市', '天津市', '上海市', '重庆市']
df_no_muni = df[~df['省份'].isin(municipalities)].copy()

df_no_muni_std = df_no_muni.copy()
df_no_muni_std = df_no_muni_std.set_index(['省份', '年份'])
scaler_nm = StandardScaler()
df_no_muni_std[X_vars] = scaler_nm.fit_transform(df_no_muni_std[X_vars])

nomuni_results, nomuni_dict = run_panel_ols(df_no_muni_std, 'TOPSIS综合得分', X_vars, '剔除直辖市')
print(f"  剔除直辖市 R² = {nomuni_results.rsquared:.4f}")

# 4. 稳健性检验3：剔除疫情年份(2020-2022)
print("\n[5/6] 稳健性检验3：剔除疫情年份...")
df_no_covid = df[df['年份'] < 2020].copy()

df_no_covid_std = df_no_covid.copy()
df_no_covid_std = df_no_covid_std.set_index(['省份', '年份'])
scaler_nc = StandardScaler()
df_no_covid_std[X_vars] = scaler_nc.fit_transform(df_no_covid_std[X_vars])

nocovid_results, nocovid_dict = run_panel_ols(df_no_covid_std, 'TOPSIS综合得分', X_vars, '剔除疫情年')
print(f"  剔除疫情年 R² = {nocovid_results.rsquared:.4f}")

# 5. 稳健性检验4：随机效应模型（Hausman检验）
print("\n[6/6] 稳健性检验4：随机效应模型...")
y = df_std['TOPSIS综合得分']
X = df_std[X_vars]

re_model = RandomEffects(y, X)
re_results = re_model.fit()

# Hausman检验（手动计算）
# H = (β_FE - β_RE)' * (V_FE - V_RE)^(-1) * (β_FE - RE)
fe_params = base_results.params
re_params = re_results.params

# 只取共同的变量
common_vars = list(set(fe_params.index) & set(re_params.index))
fe_common = fe_params[common_vars].values
re_common = re_params[common_vars].values

fe_var = base_results.cov.loc[common_vars, common_vars].values
re_var = re_results.cov.loc[common_vars, common_vars].values

diff = fe_common - re_common
var_diff = fe_var - re_var

try:
    hausman_stat = float(diff.T @ np.linalg.inv(var_diff) @ diff)
    hausman_pval = 1 - stats_chi2.cdf(hausman_stat, len(common_vars))
    hausman_result = f"拒绝RE(p={hausman_pval:.4f})" if hausman_pval < 0.05 else f"不拒绝RE(p={hausman_pval:.4f})"
except:
    hausman_stat = np.nan
    hausman_pval = np.nan
    hausman_result = "无法计算"

print(f"  随机效应 R² = {re_results.rsquared:.4f}")
print(f"  Hausman检验: χ² = {hausman_stat:.2f}, {hausman_result}")

# 6. 稳健性检验5：滞后一期回归
print("\n额外检验：滞后一期回归...")
df_lag = df.copy()
df_lag = df_lag.sort_values(['省份', '年份'])

# 对解释变量滞后一期
for var in X_vars:
    df_lag[f'{var}_lag1'] = df_lag.groupby('省份')[var].shift(1)

# 删除第一年（2015年无滞后值）
df_lag = df_lag.dropna(subset=[f'{var}_lag1' for var in X_vars])
lag_vars = [f'{var}_lag1' for var in X_vars]

df_lag_std = df_lag.copy()
df_lag_std = df_lag_std.set_index(['省份', '年份'])
scaler_lag = StandardScaler()
df_lag_std[lag_vars] = scaler_lag.fit_transform(df_lag_std[lag_vars])

lag_results, lag_dict = run_panel_ols(df_lag_std, 'TOPSIS综合得分', lag_vars, '滞后一期')
print(f"  滞后一期 R² = {lag_results.rsquared:.4f}")

# 7. 生成稳健性对比表
print("\n生成稳健性对比表...")

# 构建对比DataFrame
all_dicts = [base_dict, wins_dict, nomuni_dict, nocovid_dict, lag_dict]
compare_df = pd.DataFrame(all_dicts)
compare_df = compare_df.set_index('模型')

# 只保留系数和显著性
coef_cols = [c for c in compare_df.columns if c.endswith('_coef')]
pval_cols = [c for c in compare_df.columns if c.endswith('_pval')]
sig_cols = [c for c in compare_df.columns if c.endswith('_sig')]

# 构建格式化的系数表
formatted_rows = []
for var in X_vars:
    row = {'变量': var}
    for model_name, model_dict in zip(
        ['基准模型', '缩尾处理', '剔除直辖市', '剔除疫情年', '滞后一期'],
        all_dicts
    ):
        coef_key = f'{var}_coef'
        sig_key = f'{var}_sig'
        if coef_key in model_dict:
            row[f'{model_name}_系数'] = model_dict[coef_key]
            row[f'{model_name}_显著性'] = model_dict[sig_key]
    formatted_rows.append(row)

# 添加R²行
r2_row = {'变量': 'R²'}
for model_name, model_dict in zip(
    ['基准模型', '缩尾处理', '剔除直辖市', '剔除疫情年', '滞后一期'],
    all_dicts
):
    r2_row[f'{model_name}_系数'] = model_dict['R²']
    r2_row[f'{model_name}_显著性'] = ''
formatted_rows.append(r2_row)

df_formatted = pd.DataFrame(formatted_rows)

# 8. 生成稳健性对比图
print("生成稳健性对比图...")

fig, ax = plt.subplots(figsize=(16, 10))

# 对比各模型系数
models = ['基准模型', '缩尾处理', '剔除直辖市', '剔除疫情年', '滞后一期']
colors = ['#2c3e50', '#e74c3c', '#3498db', '#f39c12', '#2ecc71']

# 只画在基准模型中显著的变量
sig_vars = []
for var in X_vars:
    coef_key = f'{var}_coef'
    pval_key = f'{var}_pval'
    if coef_key in base_dict and base_dict[pval_key] < 0.05:
        sig_vars.append(var)

x = np.arange(len(sig_vars))
width = 0.15

for i, (model_name, color) in enumerate(zip(models, colors)):
    coefs = []
    for var in sig_vars:
        coef_key = f'{var}_coef'
        if model_name == '滞后一期':
            coef_key = f'{var}_lag1_coef'
        coefs.append(base_dict.get(coef_key, 0) if model_name == '基准模型' else 
                    all_dicts[models.index(model_name)].get(coef_key, 0))
    
    # 对于滞后模型，变量名需要加_lag1
    if model_name == '滞后一期':
        coefs_lag = []
        for var in sig_vars:
            lag_key = f'{var}_lag1_coef'
            coefs_lag.append(lag_dict.get(lag_key, 0))
        coefs = coefs_lag
    
    ax.bar(x + i * width, coefs, width, label=model_name, color=color, alpha=0.8)

ax.set_xlabel('显著变量', fontsize=12)
ax.set_ylabel('标准化回归系数', fontsize=12)
ax.set_title('图7 稳健性检验：各模型系数对比', fontsize=16, fontweight='bold')
ax.set_xticks(x + width * 2)
ax.set_xticklabels(sig_vars, rotation=45, ha='right', fontsize=9)
ax.legend(fontsize=10)
ax.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
ax.grid(True, axis='y', alpha=0.3)

plt.tight_layout()
plt.savefig('图7_稳健性检验对比.png', dpi=300, bbox_inches='tight')
print("✓ 已生成：图7_稳健性检验对比.png")
plt.close()

# 9. 保存结果到Excel
print("\n保存结果到Excel...")

with pd.ExcelWriter('Step5_稳健性检验结果.xlsx', engine='openpyxl') as writer:
    # Sheet1: 稳健性对比表
    compare_df.to_excel(writer, sheet_name='稳健性对比')
    
    # Sheet2: 格式化系数表
    df_formatted.to_excel(writer, sheet_name='系数对比表', index=False)
    
    # Sheet3: 基准模型
    base_model_df = pd.DataFrame({
        '变量': base_results.params.index,
        '系数': base_results.params.values,
        '标准误': base_results.std_errors.values,
        't统计量': base_results.tstats.values,
        'p值': base_results.pvalues.values,
        '显著性': ['***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else '' 
                   for p in base_results.pvalues.values]
    })
    base_model_df.to_excel(writer, sheet_name='基准模型', index=False)
    
    # Sheet4: 缩尾模型
    wins_model_df = pd.DataFrame({
        '变量': wins_results.params.index,
        '系数': wins_results.params.values,
        '标准误': wins_results.std_errors.values,
        't统计量': wins_results.tstats.values,
        'p值': wins_results.pvalues.values,
        '显著性': ['***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else '' 
                   for p in wins_results.pvalues.values]
    })
    wins_model_df.to_excel(writer, sheet_name='缩尾模型', index=False)
    
    # Sheet5: 剔除直辖市模型
    nomuni_model_df = pd.DataFrame({
        '变量': nomuni_results.params.index,
        '系数': nomuni_results.params.values,
        '标准误': nomuni_results.std_errors.values,
        't统计量': nomuni_results.tstats.values,
        'p值': nomuni_results.pvalues.values,
        '显著性': ['***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else '' 
                   for p in nomuni_results.pvalues.values]
    })
    nomuni_model_df.to_excel(writer, sheet_name='剔除直辖市模型', index=False)
    
    # Sheet6: Hausman检验
    hausman_df = pd.DataFrame({
        '检验方法': ['Hausman检验'],
        'χ²统计量': [hausman_stat],
        'p值': [hausman_pval],
        '结论': [hausman_result]
    })
    hausman_df.to_excel(writer, sheet_name='Hausman检验', index=False)

print("✓ 已保存：Step5_稳健性检验结果.xlsx")

# 10. 输出核心结论
print("\n" + "="*60)
print("稳健性检验核心结论：")
print("="*60)

print("\n【各模型R²对比】")
for model_name, model_dict in zip(
    ['基准模型', '缩尾处理', '剔除直辖市', '剔除疫情年', '滞后一期'],
    all_dicts
):
    print(f"  {model_name}: R² = {model_dict['R²']:.4f}")

print("\n【系数方向一致性检查】")
for var in sig_vars:
    directions = []
    for model_name, model_dict in zip(
        ['基准模型', '缩尾处理', '剔除直辖市', '剔除疫情年', '滞后一期'],
        all_dicts
    ):
        coef_key = f'{var}_coef'
        if model_name == '滞后一期':
            lag_key = f'{var}_lag1_coef'
            coef = lag_dict.get(lag_key, 0)
        else:
            coef = model_dict.get(coef_key, 0)
        directions.append('+' if coef > 0 else '-')
    
    consistent = len(set(directions)) == 1
    print(f"  {var}: {directions} {'✓一致' if consistent else '✗不一致'}")

print(f"\n【Hausman检验】χ² = {hausman_stat:.2f}, {hausman_result}")

print("\n" + "="*60)
print("Step5完成！")
print("="*60)
