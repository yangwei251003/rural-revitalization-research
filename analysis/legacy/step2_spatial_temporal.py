"""
第二步：时空特征统计分析 + 可视化
论文图2：2015—2022年全国乡村振兴综合得分及五大维度得分时序趋势图
论文图3：四大区域乡村振兴综合得分对比趋势图
论文图5：省域乡村振兴得分核密度估计演变图

数据来源：Step1_熵权TOPSIS结果.xlsx（保证数据一致性）
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import gaussian_kde

# 中文字体配置
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'SimSun', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# ============================================================
# 0. 从Step1结果读取综合得分和维度得分
# ============================================================
score_long = pd.read_excel('Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')
score_wide = pd.read_excel('Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_宽格式')

# 合并区域信息
region_map = {
    '东部': ['北京市','天津市','河北省','上海市','江苏省','浙江省','福建省','山东省','广东省','海南省'],
    '中部': ['山西省','安徽省','江西省','河南省','湖北省','湖南省'],
    '西部': ['内蒙古自治区','广西壮族自治区','重庆市','四川省','贵州省','云南省','西藏自治区','陕西省','甘肃省','青海省','宁夏回族自治区','新疆维吾尔自治区'],
    '东北': ['辽宁省','吉林省','黑龙江省'],
}
score_long['区域'] = score_long['省份'].map({p: r for r, ps in region_map.items() for p in ps})

print("数据加载完成，从Step1结果读取。")
print(f"综合得分范围：{score_long['TOPSIS综合得分'].min():.6f} — {score_long['TOPSIS综合得分'].max():.6f}")
print(f"省份数：{score_long['省份'].nunique()}，年份：{score_long['年份'].min()}—{score_long['年份'].max()}")

# 维度得分列
dim_score_cols = ['产业振兴得分','人才振兴得分','文化繁荣得分','生态振兴得分','组织建设得分']
has_dim_scores = all(c in score_long.columns for c in dim_score_cols)
if has_dim_scores:
    print("五大维度得分已包含在Step1结果中。")
else:
    print("警告：Step1结果中未找到维度得分，将仅分析综合得分。")

# ============================================================
# 1. 全国时序趋势（论文图2）
# ============================================================
year_mean = score_long.groupby('年份')['TOPSIS综合得分'].mean()

# 五大维度时序均值
dim_year_mean = {}
if has_dim_scores:
    for dim in dim_score_cols:
        dim_year_mean[dim] = score_long.groupby('年份')[dim].mean()

plt.figure(figsize=(11, 6.5))
years = list(range(2015, 2023))
plt.plot(years, year_mean.values, marker='o', linewidth=2.8, color='#1f77b4', label='综合得分', markersize=7)

if has_dim_scores:
    colors = {
        '产业振兴得分': '#ff7f0e',
        '人才振兴得分': '#2ca02c',
        '文化繁荣得分': '#d62728',
        '生态振兴得分': '#9467bd',
        '组织建设得分': '#8c564b',
    }
    for dim, color in colors.items():
        label = dim.replace('得分','')
        plt.plot(years, dim_year_mean[dim].values, marker='s', linewidth=2.0,
                 color=color, label=label+'得分', alpha=0.75, markersize=5)

plt.xlabel('年份', fontsize=13)
plt.ylabel('得分', fontsize=13)
plt.title('图2  2015—2022年全国乡村振兴综合得分及五大维度得分时序趋势', fontsize=14, fontweight='bold')
plt.legend(loc='lower right', fontsize=10, framealpha=0.9)
plt.grid(True, alpha=0.3, linestyle='--')
plt.tight_layout()
plt.savefig('图2_全国时序趋势.png', dpi=300, bbox_inches='tight')
plt.close()
print("✅ 图2已保存")

# ============================================================
# 2. 四大区域对比趋势（论文图3）
# ============================================================
region_year = score_long.groupby(['区域','年份'])['TOPSIS综合得分'].mean().unstack()

plt.figure(figsize=(11, 6.5))
colors_region = {'东部': '#d73027', '中部': '#fc8d59', '西部': '#91bfdb', '东北': '#4575b4'}
for region in ['东部','中部','西部','东北']:
    plt.plot(years, region_year.loc[region].values, marker='o', linewidth=2.5,
              color=colors_region[region], label=region, markersize=6, alpha=0.85)
plt.xlabel('年份', fontsize=13)
plt.ylabel('综合得分均值', fontsize=13)
plt.title('图3  四大区域乡村振兴综合得分对比趋势（2015—2022）', fontsize=14, fontweight='bold')
plt.legend(fontsize=11, framealpha=0.9)
plt.grid(True, alpha=0.3, linestyle='--')
plt.tight_layout()
plt.savefig('图3_四大区域对比趋势.png', dpi=300, bbox_inches='tight')
plt.close()
print("✅ 图3已保存")

# ============================================================
# 3. 核密度估计演变图（论文图5）
# ============================================================
plt.figure(figsize=(11, 6.5))
colors_kde = plt.cm.viridis(np.linspace(0.15, 0.95, 8))
for i, y in enumerate(years):
    data = score_long[score_long['年份']==y]['TOPSIS综合得分'].values
    kde = gaussian_kde(data)
    x_grid = np.linspace(data.min()-0.03, data.max()+0.03, 300)
    plt.plot(x_grid, kde(x_grid), color=colors_kde[i], linewidth=2.2, label=str(y), alpha=0.82)
plt.xlabel('乡村振兴综合得分', fontsize=13)
plt.ylabel('核密度', fontsize=13)
plt.title('图5  省域乡村振兴综合得分核密度估计演变图（2015—2022）', fontsize=14, fontweight='bold')
plt.legend(title='年份', fontsize=9.5, title_fontsize=9.5, ncol=2, framealpha=0.9)
plt.grid(True, alpha=0.25, linestyle='--')
plt.tight_layout()
plt.savefig('图5_核密度估计演变.png', dpi=300, bbox_inches='tight')
plt.close()
print("✅ 图5已保存")

# ============================================================
# 4. 省域排名变化表（论文表4部分）
# ============================================================
rank_df = score_long.copy()
rank_df['年度排名'] = rank_df.groupby('年份')['TOPSIS综合得分'].rank(method='first', ascending=False).astype(int)

rank_2015 = rank_df[rank_df['年份']==2015][['省份','TOPSIS综合得分','年度排名']].sort_values('年度排名').reset_index(drop=True)
rank_2022 = rank_df[rank_df['年份']==2022][['省份','TOPSIS综合得分','年度排名']].sort_values('年度排名').reset_index(drop=True)
rank_2015.columns = ['省份','2015年得分','2015年排名']
rank_2022.columns = ['省份','2022年得分','2022年排名']
rank_table = pd.merge(rank_2015, rank_2022, on='省份')
rank_table['得分变化'] = (rank_table['2022年得分'] - rank_table['2015年得分']).round(6)
rank_table['排名变化'] = rank_table['2015年排名'] - rank_table['2022年排名']  # 正数=排名上升

# ============================================================
# 5. 写入Excel结果
# ============================================================
output = 'Step2_时空特征分析结果.xlsx'
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    # Sheet1：时序统计描述
    desc = score_long.groupby('年份')['TOPSIS综合得分'].agg(['mean','std','min','max']).round(6)
    desc.columns = ['均值','标准差','最小值','最大值']
    desc.to_excel(writer, sheet_name='时序统计描述')
    # Sheet2：四大区域各年均值
    region_year.round(6).to_excel(writer, sheet_name='区域对比')
    # Sheet3：省域排名变化（2015 vs 2022）
    rank_table.to_excel(writer, sheet_name='省域排名变化', index=False)
    # Sheet4：五大维度时序得分
    if has_dim_scores:
        dim_desc = score_long.groupby('年份')[dim_score_cols].mean().round(6)
        dim_desc.to_excel(writer, sheet_name='维度时序得分')
    # Sheet5：完整得分数据（含区域、排名）
    out_cols = ['省份','区域','年份','TOPSIS综合得分','年度排名']
    if has_dim_scores:
        out_cols += dim_score_cols
    rank_df = score_long.copy()
    rank_df['年度排名'] = rank_df.groupby('年份')['TOPSIS综合得分'].rank(method='first', ascending=False).astype(int)
    out_cols2 = ['省份','区域','年份','TOPSIS综合得分','年度排名']
    if has_dim_scores:
        out_cols2 += dim_score_cols
    rank_df[out_cols2].sort_values(['年份','年度排名']).to_excel(writer, sheet_name='完整得分数据', index=False)

print(f"\n✅ 结果已保存：{output}")
print(f"   包含Sheet：时序统计描述 / 区域对比 / 省域排名变化 / 维度时序得分 / 完整得分数据")
print(f"\n=== 全国时序趋势（2015—2022）===")
for y, v in year_mean.items():
    print(f"   {y}年：{v:.6f}")
print(f"\n=== 2022年四大区域对比 ===")
for r in ['东部','中部','西部','东北']:
    print(f"   {r}：{region_year.loc[r,2022]:.6f}")
print(f"\n=== 数据一致性确认 ===")
print(f"   Step1综合得分均值（2015）：{score_wide[2015].mean():.10f}")
print(f"   第二步综合得分均值（2015）：{year_mean[2015]:.10f}")
print(f"   {'✅ 数据完全一致' if abs(score_wide[2015].mean()-year_mean[2015])<1e-10 else '❌ 数据不一致！'}")
print(f"\n第二步（时空特征分析）完成！")
print(f"下一步：第三步（Dagum基尼系数区域差异分解）")
