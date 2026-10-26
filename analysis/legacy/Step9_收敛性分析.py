"""
Step9: 收敛性分析
1. σ收敛 (标准差/变异系数的时序变化)
2. 绝对β收敛 (OLS: ΔY = α + β·Y₀ + ε)
3. 条件β收敛 (面板FE: ΔY = α + β·Y_{t-1} + γX + μ_i + ε)
4. 俱乐部收敛 (四大区域分组β收敛)
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

print("=" * 60)
print("Step9: 收敛性分析")
print("=" * 60)

# ============ 1. 读取数据 ============
score_df = pd.read_excel('02_分析结果/Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')
provinces = score_df['省份'].unique()
years = sorted(score_df['年份'].unique())

# 区域划分
region_map = {
    '东部': ['北京市', '天津市', '河北省', '上海市', '江苏省', '浙江省', '福建省', '山东省', '广东省', '海南省'],
    '中部': ['山西省', '安徽省', '江西省', '河南省', '湖北省', '湖南省'],
    '西部': ['内蒙古自治区', '广西壮族自治区', '重庆市', '四川省', '贵州省', '云南省', '西藏自治区', '陕西省', '甘肃省', '青海省', '宁夏回族自治区', '新疆维吾尔自治区'],
    '东北': ['辽宁省', '吉林省', '黑龙江省']
}

prov_region = {}
for reg, provs in region_map.items():
    for p in provs:
        prov_region[p] = reg

score_df['区域'] = score_df['省份'].map(prov_region)
print(f"数据: {score_df.shape}, 省份: {len(provinces)}, 年份: {years}")

# ============ 2. σ收敛 ============
print("\n--- σ收敛分析 ---")

# 全国和各区域的σ(标准差)和变异系数
sigma_results = []
for year in years:
    yr_data = score_df[score_df['年份'] == year]

    # 全国
    nat_scores = yr_data['TOPSIS综合得分']
    sigma_results.append({
        '区域': '全国', '年份': year,
        '均值': round(nat_scores.mean(), 4),
        '标准差(σ)': round(nat_scores.std(), 4),
        '变异系数(CV)': round(nat_scores.std() / nat_scores.mean(), 4)
    })

    # 四大区域
    for reg in ['东部', '中部', '西部', '东北']:
        reg_scores = yr_data[yr_data['区域'] == reg]['TOPSIS综合得分']
        if len(reg_scores) > 1:
            sigma_results.append({
                '区域': reg, '年份': year,
                '均值': round(reg_scores.mean(), 4),
                '标准差(σ)': round(reg_scores.std(), 4),
                '变异系数(CV)': round(reg_scores.std() / reg_scores.mean(), 4) if reg_scores.mean() > 0 else np.nan
            })

sigma_df = pd.DataFrame(sigma_results)
print(sigma_df.to_string(index=False))

# 判断σ收敛
nat_sigma = sigma_df[sigma_df['区域'] == '全国'].sort_values('年份')
sigma_start = nat_sigma['标准差(σ)'].iloc[0]
sigma_end = nat_sigma['标准差(σ)'].iloc[-1]
cv_start = nat_sigma['变异系数(CV)'].iloc[0]
cv_end = nat_sigma['变异系数(CV)'].iloc[-1]
print(f"\n全国σ: {sigma_start:.4f} → {sigma_end:.4f}, 变化: {'收敛↓' if sigma_end < sigma_start else '发散↑'}")
print(f"全国CV: {cv_start:.4f} → {cv_end:.4f}, 变化: {'收敛↓' if cv_end < cv_start else '发散↑'}")

# ============ 3. 绝对β收敛 ============
print("\n--- 绝对β收敛分析 ---")

from scipy import stats

# 绝对β收敛: (ln Y_T - ln Y_0) / T = α + β · ln Y_0 + ε
# 或者逐年: Δln Y_t = α + β · ln Y_{t-1} + ε
# β < 0 表示收敛

# 方法1: 长期收敛(2015→2022)
Y_2015 = score_df[score_df['年份'] == 2015].set_index('省份').loc[provinces, 'TOPSIS综合得分']
Y_2022 = score_df[score_df['年份'] == 2022].set_index('省份').loc[provinces, 'TOPSIS综合得分']

T = 2022 - 2015
growth_rate = (np.log(Y_2022) - np.log(Y_2015)) / T
ln_Y0 = np.log(Y_2015)

slope, intercept, r_value, p_value, std_err = stats.linregress(ln_Y0, growth_rate)
print(f"长期绝对β收敛 (2015→2022):")
print(f"  β = {slope:.6f}, t = {slope/std_err:.4f}, p = {p_value:.6f}")
print(f"  {'✅ 收敛' if slope < 0 and p_value < 0.05 else '❌ 不显著或发散'}")
print(f"  半衰期 = {np.log(2)/abs(slope):.2f} 年" if slope < 0 else "")

# 方法2: 逐年面板收敛
print(f"\n逐年绝对β收敛 (面板OLS):")
beta_panel = []
for year in years[1:]:
    Y_t = score_df[score_df['年份'] == year].set_index('省份').loc[provinces, 'TOPSIS综合得分']
    Y_t1 = score_df[score_df['年份'] == year-1].set_index('省份').loc[provinces, 'TOPSIS综合得分']

    dY = np.log(Y_t) - np.log(Y_t1)
    lnY_prev = np.log(Y_t1)

    slope_y, intercept_y, r_y, p_y, se_y = stats.linregress(lnY_prev, dY)
    sig = ''
    if p_y < 0.01: sig = '***'
    elif p_y < 0.05: sig = '**'
    elif p_y < 0.1: sig = '*'
    beta_panel.append({
        '年份区间': f'{year-1}→{year}',
        'β系数': round(slope_y, 6),
        '标准误': round(se_y, 6),
        't值': round(slope_y/se_y, 4),
        'p值': round(p_y, 4),
        '显著性': sig,
        '收敛性': '收敛' if slope_y < 0 else '发散'
    })
    print(f"  {year-1}→{year}: β={slope_y:.6f}, p={p_y:.4f} {sig} {'收敛' if slope_y < 0 else '发散'}")

beta_panel_df = pd.DataFrame(beta_panel)

# ============ 4. 条件β收敛 (面板固定效应) ============
print("\n--- 条件β收敛分析 (面板双向固定效应) ---")

# 模型: Δln Y_it = α + β · ln Y_{i,t-1} + γX_it + μ_i + λ_t + ε_it
# 简化: 使用省份均值去势 + 年份哑变量

panel_data = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')

# 构建回归数据
reg_data = score_df[['省份', '年份', 'TOPSIS综合得分']].copy()
reg_data['区域'] = reg_data['省份'].map(prov_region)

# 滞后一期的得分
reg_data = reg_data.sort_values(['省份', '年份'])
reg_data['lnY'] = np.log(reg_data['TOPSIS综合得分'])
reg_data['lnY_lag1'] = reg_data.groupby('省份')['lnY'].shift(1)

# 增长率
reg_data['dlnY'] = reg_data.groupby('省份')['lnY'].diff()

# 加入控制变量
reg_data = reg_data.merge(
    panel_data[['省份', '年份',
                'X01_农村居民人均可支配收入(元)',
                'X14_农林水支出占财政支出比重(%)',
                'X07_文教娱乐支出占消费比重(%)',
                'X08_每万乡村人口卫生技术人员数(人)']],
    on=['省份', '年份'], how='left'
)
reg_data.rename(columns={
    'X01_农村居民人均可支配收入(元)': 'Econ',
    'X14_农林水支出占财政支出比重(%)': 'Fiscal',
    'X07_文教娱乐支出占消费比重(%)': 'Human',
    'X08_每万乡村人口卫生技术人员数(人)': 'Digital'
}, inplace=True)

reg_data['lnEcon'] = np.log(reg_data['Econ'])
reg_data['lnDigital'] = np.log(reg_data['Digital'])

# 去掉第一年(无滞后)
reg_data = reg_data.dropna(subset=['dlnY', 'lnY_lag1'])

# 组内去均值 (固定效应)
reg_data = reg_data.sort_values(['省份', '年份'])
for var in ['dlnY', 'lnY_lag1', 'lnEcon', 'Fiscal', 'Human', 'lnDigital']:
    prov_mean = reg_data.groupby('省份')[var].transform('mean')
    reg_data[f'{var}_dm'] = reg_data[var] - prov_mean

# 年份固定效应 (哑变量)
for y in years[1:]:
    reg_data[f'year_{y}'] = (reg_data['年份'] == y).astype(int)
    # 去均值
    reg_data[f'year_{y}_dm'] = reg_data[f'year_{y}'] - reg_data.groupby('省份')[f'year_{y}'].transform('mean')

year_dummies = [f'year_{y}_dm' for y in years[2:]]  # 去掉第一个和最后一个避免共线

# 条件β收敛回归
X_vars_cond = ['lnY_lag1_dm', 'lnEcon_dm', 'Fiscal_dm', 'Human_dm', 'lnDigital_dm'] + year_dummies
y_var = 'dlnY_dm'

y_vec = reg_data[y_var].values
X_mat = reg_data[X_vars_cond].values
X_with_const = np.column_stack([np.ones(len(y_vec)), X_mat])

# OLS (组内估计量)
beta_cond = np.linalg.lstsq(X_with_const, y_vec, rcond=None)[0]
y_hat = X_with_const @ beta_cond
resid = y_vec - y_hat
n_obs = len(y_vec)
k_vars = X_with_const.shape[1]
sigma2 = np.sum(resid**2) / (n_obs - k_vars)
var_beta = sigma2 * np.linalg.inv(X_with_const.T @ X_with_const)
se_beta = np.sqrt(np.diag(var_beta))
t_stats = beta_cond / se_beta
p_vals = 2 * (1 - stats.t.cdf(np.abs(t_stats), n_obs - k_vars))

var_names_cond = ['截距', 'lnY_lag1', 'lnEcon', 'Fiscal', 'Human', 'lnDigital'] + [f'year_{y}' for y in years[2:]]

cond_results = []
for i, name in enumerate(var_names_cond):
    sig = ''
    if p_vals[i] < 0.01: sig = '***'
    elif p_vals[i] < 0.05: sig = '**'
    elif p_vals[i] < 0.1: sig = '*'
    cond_results.append({
        '变量': name,
        '系数': round(beta_cond[i], 6),
        '标准误': round(se_beta[i], 6),
        't值': round(t_stats[i], 4),
        'p值': round(p_vals[i], 4),
        '显著性': sig
    })

cond_df = pd.DataFrame(cond_results)
print("\n条件β收敛回归结果:")
print(cond_df[cond_df['变量'].isin(['lnY_lag1', 'lnEcon', 'Fiscal', 'Human', 'lnDigital'])].to_string(index=False))

beta_cond_val = beta_cond[var_names_cond.index('lnY_lag1')]
beta_cond_p = p_vals[var_names_cond.index('lnY_lag1')]
print(f"\n  条件β = {beta_cond_val:.6f}, p = {beta_cond_p:.4f}")
print(f"  {'✅ 条件收敛' if beta_cond_val < 0 and beta_cond_p < 0.05 else '❌ 不显著或发散'}")
if beta_cond_val < 0:
    print(f"  收敛速度 λ = {abs(beta_cond_val):.4f}, 半衰期 = {np.log(2)/abs(beta_cond_val):.2f} 年")

# R²
ss_res = np.sum(resid**2)
ss_tot = np.sum((y_vec - y_vec.mean())**2)
r2 = 1 - ss_res/ss_tot
print(f"  R² = {r2:.4f}")

# ============ 5. 俱乐部收敛 (四大区域分组) ============
print("\n--- 俱乐部收敛分析 (四大区域分组) ---")

club_results = []
for reg in ['东部', '中部', '西部', '东北']:
    reg_data_sub = reg_data[reg_data['区域'] == reg]
    if len(reg_data_sub) < 10:
        continue

    # 绝对β收敛(分组OLS)
    Y_2015_reg = reg_data_sub[reg_data_sub['年份'] == 2016].set_index('省份')  # 2016因为lag
    Y_2022_reg = reg_data_sub[reg_data_sub['年份'] == 2022].set_index('省份')

    # 逐年面板
    reg_betas = []
    for year in years[2:]:
        sub_t = reg_data_sub[reg_data_sub['年份'] == year]
        sub_t1 = reg_data_sub[reg_data_sub['年份'] == year-1]
        common_provs = list(set(sub_t['省份']) & set(sub_t1['省份']))
        if len(common_provs) < 3:
            continue

        y_t = sub_t.set_index('省份').loc[common_provs, 'TOPSIS综合得分']
        y_t1 = sub_t1.set_index('省份').loc[common_provs, 'TOPSIS综合得分']

        dY = np.log(y_t) - np.log(y_t1)
        lnY_prev = np.log(y_t1)

        if len(dY) >= 3:
            slope_r, _, _, p_r, se_r = stats.linregress(lnY_prev, dY)
            reg_betas.append({'β': slope_r, 'p': p_r})

    if reg_betas:
        avg_beta = np.mean([b['β'] for b in reg_betas])
        avg_p = np.mean([b['p'] for b in reg_betas])
        sig = ''
        if avg_p < 0.01: sig = '***'
        elif avg_p < 0.05: sig = '**'
        elif avg_p < 0.1: sig = '*'
        club_results.append({
            '区域': reg,
            '省份': ', '.join(region_map[reg]),
            '平均β': round(avg_beta, 6),
            '平均p值': round(avg_p, 4),
            '显著性': sig,
            '收敛性': '俱乐部收敛' if avg_beta < 0 else '发散'
        })
        print(f"  {reg}: β={avg_beta:.6f}, p={avg_p:.4f} {sig} → {'俱乐部收敛' if avg_beta < 0 else '发散'}")

    # 条件β收敛(分组)
    if len(reg_data_sub) > 20:
        y_sub = reg_data_sub['dlnY_dm'].values
        X_sub = reg_data_sub[['lnY_lag1_dm']].values
        X_sub_c = np.column_stack([np.ones(len(y_sub)), X_sub])
        beta_sub = np.linalg.lstsq(X_sub_c, y_sub, rcond=None)[0]
        resid_sub = y_sub - X_sub_c @ beta_sub
        n_sub = len(y_sub)
        sigma2_sub = np.sum(resid_sub**2) / (n_sub - 2)
        var_sub = sigma2_sub * np.linalg.inv(X_sub_c.T @ X_sub_c)
        se_sub = np.sqrt(np.diag(var_sub))
        t_sub = beta_sub / se_sub
        p_sub = 2 * (1 - stats.t.cdf(np.abs(t_sub), n_sub - 2))

        print(f"    条件β收敛: β={beta_sub[1]:.6f}, p={p_sub[1]:.4f} {'收敛' if beta_sub[1] < 0 and p_sub[1] < 0.1 else '不显著'}")

club_df = pd.DataFrame(club_results)

# ============ 6. 保存结果 ============
print("\n--- 保存结果 ---")

output_path = '02_分析结果/Step9_收敛性分析结果.xlsx'
with pd.ExcelWriter(output_path) as writer:
    sigma_df.to_excel(writer, sheet_name='σ收敛', index=False)
    beta_panel_df.to_excel(writer, sheet_name='绝对β收敛', index=False)
    cond_df.to_excel(writer, sheet_name='条件β收敛', index=False)
    if not club_df.empty:
        club_df.to_excel(writer, sheet_name='俱乐部收敛', index=False)

print(f"✅ 结果已保存: {output_path}")

# ============ 7. 绘图 ============
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False

# 图9: σ收敛趋势图
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

# 标准差
for reg in ['全国', '东部', '中部', '西部', '东北']:
    data = sigma_df[sigma_df['区域'] == reg].sort_values('年份')
    style = '-' if reg == '全国' else '--'
    width = 2.5 if reg == '全国' else 1.5
    ax1.plot(data['年份'], data['标准差(σ)'], style, linewidth=width, label=reg, marker='o', markersize=4)
ax1.set_xlabel('年份', fontsize=12)
ax1.set_ylabel('标准差(σ)', fontsize=12)
ax1.set_title('σ收敛：省域乡村振兴得分标准差', fontsize=13)
ax1.legend(fontsize=10)
ax1.grid(True, alpha=0.3)

# 变异系数
for reg in ['全国', '东部', '中部', '西部', '东北']:
    data = sigma_df[sigma_df['区域'] == reg].sort_values('年份')
    style = '-' if reg == '全国' else '--'
    width = 2.5 if reg == '全国' else 1.5
    ax2.plot(data['年份'], data['变异系数(CV)'], style, linewidth=width, label=reg, marker='o', markersize=4)
ax2.set_xlabel('年份', fontsize=12)
ax2.set_ylabel('变异系数(CV)', fontsize=12)
ax2.set_title('σ收敛：省域乡村振兴得分变异系数', fontsize=13)
ax2.legend(fontsize=10)
ax2.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig('03_图表文件/图9_σ收敛趋势.png', dpi=200, bbox_inches='tight')
plt.close()
print("✅ 图9: σ收敛趋势图已保存")

# 图10: β收敛散点图
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

# 绝对β收敛散点
Y_2015_all = score_df[score_df['年份'] == 2015].set_index('省份').loc[provinces, 'TOPSIS综合得分']
Y_2022_all = score_df[score_df['年份'] == 2022].set_index('省份').loc[provinces, 'TOPSIS综合得分']
growth = (np.log(Y_2022_all) - np.log(Y_2015_all)) / 7
lnY0 = np.log(Y_2015_all)

colors = []
for p in provinces:
    r = prov_region.get(p, '全国')
    colors.append({'东部': '#e74c3c', '中部': '#2ecc71', '西部': '#3498db', '东北': '#f39c12'}[r])

ax1.scatter(lnY0, growth, c=colors, s=60, alpha=0.8, edgecolors='white', linewidth=0.5)
# 拟合线
slope_abs, intercept_abs, _, _, _ = stats.linregress(lnY0, growth)
x_fit = np.linspace(lnY0.min(), lnY0.max(), 100)
ax1.plot(x_fit, intercept_abs + slope_abs * x_fit, 'k--', linewidth=1.5, label=f'β={slope_abs:.4f}')
ax1.set_xlabel('ln(2015年得分)', fontsize=12)
ax1.set_ylabel('年均增长率', fontsize=12)
ax1.set_title('绝对β收敛散点图 (2015→2022)', fontsize=13)
ax1.legend(fontsize=10)
ax1.grid(True, alpha=0.3)

# 自定义图例
from matplotlib.patches import Patch
legend_elements = [Patch(facecolor='#e74c3c', label='东部'),
                   Patch(facecolor='#2ecc71', label='中部'),
                   Patch(facecolor='#3498db', label='西部'),
                   Patch(facecolor='#f39c12', label='东北')]
ax1.legend(handles=legend_elements + [plt.Line2D([0], [0], color='k', linestyle='--', label=f'β={slope_abs:.4f}')],
           fontsize=9)

# 条件β收敛系数图
cond_vars = cond_df[cond_df['变量'].isin(['lnY_lag1', 'lnEcon', 'Fiscal', 'Human', 'lnDigital'])]
bar_colors = ['#e74c3c' if c < 0 else '#2ecc71' for c in cond_vars['系数']]
bars = ax2.barh(cond_vars['变量'], cond_vars['系数'], color=bar_colors, alpha=0.8, height=0.5)
ax2.axvline(x=0, color='black', linewidth=0.8)
for i, (_, row) in enumerate(cond_vars.iterrows()):
    ax2.text(row['系数'], i, f" {row['系数']:.4f}{row['显著性']}", va='center', fontsize=9)
ax2.set_xlabel('系数', fontsize=12)
ax2.set_title('条件β收敛回归系数', fontsize=13)
ax2.grid(True, alpha=0.3, axis='x')

plt.tight_layout()
plt.savefig('03_图表文件/图10_β收敛分析.png', dpi=200, bbox_inches='tight')
plt.close()
print("✅ 图10: β收敛分析图已保存")

print("\n" + "=" * 60)
print("Step9 收敛性分析完成!")
print("=" * 60)
