"""生成所有需要更新的图表（新指标体系数据）+ 完整格式化DOCX"""

from pathlib import Path as _WorkspacePath
import os as _workspace_os
RESEARCH_WORKSPACE = _WorkspacePath(_workspace_os.environ.get("RESEARCH_WORKSPACE", _workspace_os.getcwd())).resolve()

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import pandas as pd
import os
from scipy.stats import gaussian_kde

# 设置中文字体
font_paths = [
    'C:/Windows/Fonts/simsun.ttc',
    'C:/Windows/Fonts/SimHei.ttf',
    'C:/Windows/Fonts/STKAITI.TTF',
    'C:/Windows/Fonts/msyh.ttc',
]
for fp in font_paths:
    if os.path.exists(fp):
        fm.fontManager.addfont(fp)

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['figure.dpi'] = 150
plt.rcParams['savefig.dpi'] = 150
plt.rcParams['savefig.bbox'] = 'tight'

BASE = str(RESEARCH_WORKSPACE)
CHART_DIR = f'{BASE}/03_图表文件'
os.makedirs(CHART_DIR, exist_ok=True)

# ============================================================
# 加载新数据
# ============================================================
df = pd.read_excel(f'{BASE}/乡村振兴面板数据_相对量_新版.xlsx', sheet_name='面板数据')

region_map = {
    '北京市':'东部','天津市':'东部','河北省':'东部','上海市':'东部','江苏省':'东部',
    '浙江省':'东部','福建省':'东部','山东省':'东部','广东省':'东部','海南省':'东部',
    '山西省':'中部','安徽省':'中部','江西省':'中部','河南省':'中部','湖北省':'中部','湖南省':'中部',
    '辽宁省':'东北','吉林省':'东北','黑龙江省':'东北',
    '内蒙古自治区':'西部','广西壮族自治区':'西部','重庆市':'西部','四川省':'西部',
    '贵州省':'西部','云南省':'西部','西藏自治区':'西部','陕西省':'西部',
    '甘肃省':'西部','青海省':'西部','宁夏回族自治区':'西部','新疆维吾尔自治区':'西部'
}
df['区域'] = df['省份'].map(region_map)

# 重算 TOPSIS
indicator_names = ['X01_人均可支配收入','X02_农林牧渔产值','X03_农机总动力',
                   'X04_有效灌溉占比','X05_化肥强度','X06_农药强度',
                   'X07_卫生技术员密度','X08_文教娱乐支出占比','X09_每百万人图书馆',
                   'X10_森林覆盖率','X11_农林水支出占比','X12_每万人村委会',
                   'X13_每万人养老参保','X14_每万人医疗参保']
reverse_indicators = ['X05_化肥强度', 'X06_农药强度']
dimensions = {
    '产业振兴': indicator_names[:6],
    '人才振兴': [indicator_names[6]],
    '文化振兴': indicator_names[7:9],
    '生态振兴': [indicator_names[9]],
    '组织建设': indicator_names[10:14]
}

n = len(df)
df_std = df.copy()
for col in indicator_names:
    x = df[col].values
    xmin, xmax = x.min(), x.max()
    if col in reverse_indicators:
        df_std[col] = (xmax - x) / (xmax - xmin)
    else:
        df_std[col] = (x - xmin) / (xmax - xmin)

weights = {}
for col in indicator_names:
    p = df_std[col] / df_std[col].sum()
    p = p.replace(0, 1e-10)
    e = -np.sum(p * np.log(p)) / np.log(n)
    weights[col] = (1 - e)
total_d = sum(weights.values())
for col in weights: weights[col] /= total_d

weighted = df_std.copy()
for col in indicator_names:
    weighted[col] = df_std[col] * weights[col]
best = weighted[indicator_names].max()
worst = weighted[indicator_names].min()
d_best = np.sqrt(((weighted[indicator_names] - best)**2).sum(axis=1))
d_worst = np.sqrt(((weighted[indicator_names] - worst)**2).sum(axis=1))
df['TOPSIS'] = d_worst / (d_best + d_worst)

for dim_name, dim_cols in dimensions.items():
    w_sub = {c: weights[c] for c in dim_cols}
    w_total = sum(w_sub.values())
    w_sub_norm = {c: w/w_total for c, w in w_sub.items()}
    w_dim = df_std.copy()
    for c in dim_cols:
        w_dim[c] = df_std[c] * w_sub_norm[c]
    best_d = w_dim[dim_cols].max()
    worst_d = w_dim[dim_cols].min()
    df[f'{dim_name}_得分'] = np.sqrt(((w_dim[dim_cols] - worst_d)**2).sum(axis=1)) /\
                              (np.sqrt(((w_dim[dim_cols] - best_d)**2).sum(axis=1)) + 
                               np.sqrt(((w_dim[dim_cols] - worst_d)**2).sum(axis=1)))

# ============================================================
# 图2: 全国时序趋势
# ============================================================
print("生成图2: 全国时序趋势...")
yearly = df.groupby('年份').agg(
    综合=('TOPSIS','mean'),
    **{d: (f'{d}_得分', 'mean') for d in dimensions}
)

fig, ax = plt.subplots(figsize=(10, 5))
years = yearly.index
ax.plot(years, yearly['综合'], 'ko-', linewidth=2.5, markersize=8, label='综合得分')
colors = ['#E74C3C', '#3498DB', '#2ECC71', '#F39C12', '#9B59B6']
for i, d in enumerate(dimensions):
    ax.plot(years, yearly[d], 's--', color=colors[i], linewidth=1.5, markersize=5, label=d)
ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel('TOPSIS得分', fontsize=12)
ax.set_title('图2  2015—2022年全国乡村振兴综合得分及五大维度得分时序趋势', fontsize=13)
ax.legend(loc='upper left', fontsize=9, ncol=2)
ax.grid(True, alpha=0.3)
ax.set_xticks(years)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图2_全国时序趋势_NEW.png')
plt.close()

# ============================================================
# 图3: 四大区域对比趋势  
# ============================================================
print("生成图3: 四大区域对比趋势...")
region_data = df.groupby(['区域','年份'])['TOPSIS'].mean().unstack(level=0)
fig, ax = plt.subplots(figsize=(10, 5))
colors_r = {'东部': '#E74C3C', '中部': '#3498DB', '西部': '#2ECC71', '东北': '#F39C12'}
for reg in ['东部','中部','西部','东北']:
    ax.plot(region_data.index, region_data[reg], 'o-', color=colors_r[reg], linewidth=2, markersize=6, label=reg)
ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel('综合得分均值', fontsize=12)
ax.set_title('图3  四大区域乡村振兴综合得分对比趋势', fontsize=13)
ax.legend(fontsize=10)
ax.grid(True, alpha=0.3)
ax.set_xticks(years)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图3_四大区域趋势_NEW.png')
plt.close()

# ============================================================
# 图5: 核密度估计
# ============================================================
print("生成图5: 核密度估计演变...")
fig, ax = plt.subplots(figsize=(10, 5))
for yr in [2015, 2018, 2022]:
    data = df[df['年份']==yr]['TOPSIS'].dropna()
    kde = gaussian_kde(data)
    x_grid = np.linspace(data.min()-0.02, data.max()+0.02, 200)
    ax.plot(x_grid, kde(x_grid), linewidth=2, label=f'{yr}年')
ax.set_xlabel('TOPSIS综合得分', fontsize=12)
ax.set_ylabel('密度', fontsize=12)
ax.set_title('图5  省域乡村振兴综合得分核密度估计演变', fontsize=13)
ax.legend(fontsize=11)
ax.grid(True, alpha=0.3)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图5_核密度_NEW.png')
plt.close()

# ============================================================
# 图6: 全局Moran's I时序趋势
# ============================================================
print("生成图6: Moran's I时序趋势...")
moran_data = {
    2015: 0.930, 2016: 0.921, 2017: 0.920, 2018: 0.895,
    2019: 0.909, 2020: 0.914, 2021: 0.918, 2022: 0.923
}
fig, ax = plt.subplots(figsize=(8, 4.5))
years_m = list(moran_data.keys())
vals_m = list(moran_data.values())
ax.plot(years_m, vals_m, 'bo-', linewidth=2, markersize=8)
ax.fill_between(years_m, [v-0.01 for v in vals_m], [v+0.01 for v in vals_m], alpha=0.15, color='blue')
ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel("Moran's I", fontsize=12)
ax.set_title("图6  全局Moran's I时序趋势", fontsize=13)
ax.grid(True, alpha=0.3)
ax.set_ylim(0.88, 0.95)
for yr, v in moran_data.items():
    ax.annotate(f'{v:.3f}', (yr, v), textcoords="offset points", xytext=(0,10), ha='center', fontsize=8)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图6_MoranI_NEW.png')
plt.close()

# ============================================================
# 图7: Dagum基尼系数分解
# ============================================================
print("生成图7: Dagum基尼系数分解...")
dagum_data = {
    2015: (0.124, 29.6, 44.0, 26.4),
    2016: (0.116, 29.4, 41.4, 29.2),
    2017: (0.116, 29.3, 43.9, 26.7),
    2018: (0.099, 30.1, 39.6, 30.3),
    2019: (0.108, 30.7, 36.3, 33.0),
    2020: (0.105, 30.4, 36.0, 33.6),
    2021: (0.106, 30.6, 36.5, 32.9),
    2022: (0.106, 30.9, 35.8, 33.3),
}

fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5))

years_d = list(dagum_data.keys())
G_vals = [d[0] for d in dagum_data.values()]
ax1.plot(years_d, G_vals, 'ro-', linewidth=2, markersize=8)
ax1.set_xlabel('年份', fontsize=12)
ax1.set_ylabel('总体基尼系数', fontsize=12)
ax1.set_title('总体基尼系数演变', fontsize=13)
ax1.grid(True, alpha=0.3)

gw = [d[1] for d in dagum_data.values()]
gnb = [d[2] for d in dagum_data.values()]
gt = [d[3] for d in dagum_data.values()]
ax2.stackplot(years_d, gw, gnb, gt, labels=['区域内', '区域间', '超变密度'],
              colors=['#3498DB', '#E74C3C', '#2ECC71'], alpha=0.7)
ax2.set_xlabel('年份', fontsize=12)
ax2.set_ylabel('贡献率 (%)', fontsize=12)
ax2.set_title('差异来源分解贡献率', fontsize=13)
ax2.legend(loc='upper right', fontsize=9)
ax2.grid(True, alpha=0.3)

fig.suptitle('图7  Dagum基尼系数及分解项贡献度时序趋势', fontsize=14)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图7_Dagum_NEW.png')
plt.close()

# ============================================================
# 图8: 基准回归系数森林图
# ============================================================
print("生成图8: 基准回归系数森林图...")
from sklearn.linear_model import LinearRegression
df_fe = df.copy()
for col in ['TOPSIS','lnEcon','Fiscal','Human','lnDigital']:
    province_mean = df_fe.groupby('省份')[col].transform('mean')
    year_mean = df_fe.groupby('年份')[col].transform('mean')
    grand_mean = df_fe[col].mean()
    df_fe[f'{col}_fe'] = df_fe[col] - province_mean - year_mean + grand_mean

X_fe = df_fe[['lnEcon_fe','Fiscal_fe','Human_fe','lnDigital_fe']].values
y_fe = df_fe['TOPSIS_fe'].values
model = LinearRegression().fit(X_fe, y_fe)
resid = y_fe - model.predict(X_fe)
sigma2 = np.sum(resid**2)/(len(y_fe)-4)
X_inv = np.linalg.inv(X_fe.T @ X_fe)
se = np.sqrt(np.diag(sigma2 * X_inv))

var_names = ['lnEcon','Fiscal','Human','lnDigital']
var_labels = ['经济发展水平\n(lnEcon)', '财政支农力度\n(Fiscal)', '人力资本水平\n(Human)', '数字基础设施\n(lnDigital)']
coefs = model.coef_

fig, ax = plt.subplots(figsize=(9, 5))
y_pos = np.arange(len(var_names))
colors_f = ['#E74C3C' if c < 0 else '#2ECC71' for c in coefs]
ax.barh(y_pos, coefs, xerr=se*1.96, color=colors_f, alpha=0.8, capsize=4, height=0.5)
ax.set_yticks(y_pos)
ax.set_yticklabels(var_labels, fontsize=10)
ax.set_xlabel('回归系数', fontsize=12)
ax.set_title('图8  基准面板模型回归系数（双向固定效应）', fontsize=13)
ax.axvline(x=0, color='black', linestyle='-', linewidth=0.8)
ax.grid(True, alpha=0.3, axis='x')
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图8_回归系数_NEW.png')
plt.close()

# ============================================================
# 图9: σ收敛趋势
# ============================================================
print("生成图9: σ收敛趋势...")
sigma_data = {
    2015: (0.247, 0.240, 0.170, 0.255, 0.250),
    2016: (0.241, 0.231, 0.168, 0.249, 0.245),
    2017: (0.241, 0.235, 0.161, 0.248, 0.250),
    2018: (0.229, 0.226, 0.155, 0.233, 0.248),
    2019: (0.241, 0.249, 0.153, 0.237, 0.252),
    2020: (0.240, 0.246, 0.155, 0.240, 0.244),
    2021: (0.244, 0.247, 0.161, 0.240, 0.252),
    2022: (0.248, 0.252, 0.164, 0.240, 0.256),
}

fig, ax = plt.subplots(figsize=(10, 5))
years_s = list(sigma_data.keys())
labels_s = ['全国','东部','中部','西部','东北']
colors_s = ['black','#E74C3C','#3498DB','#2ECC71','#F39C12']
styles = ['-','--','--','--','--']
widths = [2.5, 1.5, 1.5, 1.5, 1.5]

for i in range(5):
    vals = [sigma_data[yr][i] for yr in years_s]
    ax.plot(years_s, vals, color=colors_s[i], linestyle=styles[i], 
            linewidth=widths[i], marker='o', markersize=5, label=labels_s[i])
ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel('变异系数', fontsize=12)
ax.set_title('图9  全国及四大区域乡村振兴变异系数时序趋势', fontsize=13)
ax.legend(fontsize=10, ncol=5, loc='upper center', bbox_to_anchor=(0.5, -0.12))
ax.grid(True, alpha=0.3)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图9_σ收敛_NEW.png')
plt.close()

# ============================================================
# 图10: β收敛分析
# ============================================================
print("生成图10: β收敛分析...")
df_sorted = df.sort_values(['省份','年份'])
df_sorted['Y_lag'] = df_sorted.groupby('省份')['TOPSIS'].shift(1)
df_sorted['lnY_lag'] = np.log(df_sorted['Y_lag'])
df_sorted['lnY'] = np.log(df_sorted['TOPSIS'])
df_sorted['growth'] = df_sorted['lnY'] - df_sorted['lnY_lag']
df_f = df_sorted.dropna(subset=['Y_lag'])

fig, ax = plt.subplots(figsize=(9, 6))
regions_p = ['东部','中部','西部','东北']
colors_p = {'东部':'#E74C3C','中部':'#3498DB','西部':'#2ECC71','东北':'#F39C12'}
markers_p = {'东部':'o','中部':'s','西部':'^','东北':'D'}

for reg in regions_p:
    df_r = df_f[df_f['区域']==reg]
    m = LinearRegression().fit(df_r[['lnY_lag']].values, df_r['growth'].values)
    x_pred = np.linspace(df_r['lnY_lag'].min(), df_r['lnY_lag'].max(), 50)
    y_pred = m.intercept_ + m.coef_[0] * x_pred
    ax.plot(x_pred, y_pred, color=colors_p[reg], linewidth=2, 
            label=f'{reg} (β={m.coef_[0]:.3f})')
    ax.scatter(df_r['lnY_lag'], df_r['growth'], color=colors_p[reg], 
              marker=markers_p[reg], alpha=0.5, s=30)

ax.set_xlabel('ln(滞后一期得分)', fontsize=12)
ax.set_ylabel('增长率 Δln(得分)', fontsize=12)
ax.set_title('图10  俱乐部β收敛拟合关系', fontsize=13)
ax.axhline(y=0, color='gray', linestyle='--', alpha=0.5)
ax.legend(fontsize=10)
ax.grid(True, alpha=0.3)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图10_β收敛_NEW.png')
plt.close()

# ============================================================
# 图11: 区域障碍度雷达图
# ============================================================
print("生成图11: 区域障碍度雷达图...")
obs_dims_data = {
    '东部': [32.48, 10.78, 25.63, 2.94, 28.17],
    '中部': [33.18, 14.23, 25.07, 4.37, 23.15],
    '西部': [38.94, 14.00, 21.14, 4.69, 21.23],
    '东北': [34.87, 12.91, 23.37, 2.93, 25.92],
}
dim_labels = ['产业振兴','人才振兴','文化振兴','生态振兴','组织建设']

angles = np.linspace(0, 2*np.pi, len(dim_labels), endpoint=False).tolist()
angles += angles[:1]

fig, ax = plt.subplots(figsize=(8, 8), subplot_kw={'projection': 'polar'})
colors_r2 = {'东部':'#E74C3C','中部':'#3498DB','西部':'#2ECC71','东北':'#F39C12'}
for reg, vals in obs_dims_data.items():
    vals_r = vals + vals[:1]
    ax.plot(angles, vals_r, 'o-', color=colors_r2[reg], linewidth=2, markersize=6, label=reg)
    ax.fill(angles, vals_r, alpha=0.1, color=colors_r2[reg])

ax.set_xticks(angles[:-1])
ax.set_xticklabels(dim_labels, fontsize=11)
ax.set_title('图11  四大区域维度障碍度雷达图', fontsize=14, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1), fontsize=10)
ax.set_ylim(0, 45)
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图11_障碍度雷达_NEW.png')
plt.close()

# ============================================================
# 图12: 全国障碍指标排名
# ============================================================
print("生成图12: 全国障碍指标排名...")
weights_array = np.array([weights[c] for c in indicator_names])
obs_summary = {}
for idx, row in df.iterrows():
    std_vals = np.array([df_std.loc[idx, c] for c in indicator_names])
    gap = 1 - std_vals
    wg = gap * weights_array
    total = wg.sum()
    for i, ind in enumerate(indicator_names):
        obs_summary.setdefault(ind, []).append(wg[i]/total*100)

ind_obs_mean = {ind: np.mean(vals) for ind, vals in obs_summary.items()}
ind_sorted = sorted(ind_obs_mean.items(), key=lambda x: -x[1])[:10]

labels_obs = [f"X{ind[1:3]}" for ind, _ in ind_sorted]
vals_obs = [v for _, v in ind_sorted]

fig, ax = plt.subplots(figsize=(10, 5))
colors_obs = ['#E74C3C']*3 + ['#F39C12']*3 + ['#3498DB']*4
bars = ax.barh(range(len(labels_obs)), vals_obs, color=colors_obs, alpha=0.8, height=0.6)
ax.set_yticks(range(len(labels_obs)))
ax.set_yticklabels(labels_obs, fontsize=10)
ax.set_xlabel('障碍度 (%)', fontsize=12)
ax.set_title('图12  全国乡村振兴指标障碍度排名（前10）', fontsize=13)
ax.invert_yaxis()
for bar, val in zip(bars, vals_obs):
    ax.text(bar.get_width()+0.3, bar.get_y()+bar.get_height()/2, f'{val:.1f}%', va='center', fontsize=9)
ax.grid(True, alpha=0.3, axis='x')
fig.tight_layout()
fig.savefig(f'{CHART_DIR}/图12_障碍指标_NEW.png')
plt.close()

print("\n所有图表生成完成！")
print(f"保存至: {CHART_DIR}/")
