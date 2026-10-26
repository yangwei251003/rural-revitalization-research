"""
Step10: 障碍度诊断模型
1. 各指标障碍度计算
2. 各维度障碍度计算
3. 分省份障碍度诊断
4. 分区域障碍度对比
5. 时序障碍度演变
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

print("=" * 60)
print("Step10: 障碍度诊断模型")
print("=" * 60)

# ============ 1. 读取数据 ============
score_df = pd.read_excel('02_分析结果/Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')
weight_df = pd.read_excel('02_分析结果/Step1_熵权TOPSIS结果.xlsx', sheet_name='指标权重')

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

print(f"数据: {score_df.shape}, 省份: {len(provinces)}, 年份: {years}")

# ============ 2. 准备指标数据与权重 ============
print("\n--- 准备指标数据与权重 ---")

# 读取面板数据(标准化前的原始数据)
panel_data = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')

# 指标信息
indicators = {
    'X01_农村居民人均可支配收入(元)': {'维度': '产业振兴', '方向': '正向'},
    'X02_农林牧渔业总产值(亿元)': {'维度': '产业振兴', '方向': '正向'},
    'X03_农业机械总动力(万千瓦)': {'维度': '产业振兴', '方向': '正向'},
    'X04_有效灌溉面积占比(%)': {'维度': '产业振兴', '方向': '正向'},
    'X05_化肥施用强度(万吨/千公顷)': {'维度': '产业振兴', '方向': '逆向'},
    'X06_农药施用强度(万吨/千公顷)': {'维度': '产业振兴', '方向': '逆向'},
    'X07_文教娱乐支出占消费比重(%)': {'维度': '人才振兴', '方向': '正向'},
    'X08_每万乡村人口卫生技术人员数(人)': {'维度': '人才振兴', '方向': '正向'},
    'X09_公共图书馆数量(个)': {'维度': '文化繁荣', '方向': '正向'},
    'X11_森林覆盖率(%)': {'维度': '生态振兴', '方向': '正向'},
    'X14_农林水支出占财政支出比重(%)': {'维度': '组织建设', '方向': '正向'},
    'X15_村民委员会数量(个)': {'维度': '组织建设', '方向': '正向'},
    'X16_城乡居民养老保险参保人数(万人)': {'维度': '组织建设', '方向': '正向'},
    'X17_城乡居民医疗保险参保人数(万人)': {'维度': '组织建设', '方向': '正向'},
}

# 获取权重
weight_dict = {}
for _, row in weight_df.iterrows():
    code = row['指标代码']
    weight_dict[code] = row['熵权权重']

print(f"指标权重: {len(weight_dict)}个")
for k, v in weight_dict.items():
    print(f"  {k}: {v:.4f}")

# ============ 3. 障碍度模型计算 ============
print("\n--- 障碍度模型计算 ---")

# 障碍度模型:
# 偏离度 R_i = 1 - x_i_normalized (正向指标) 或 x_i_normalized (逆向指标)
# 指标障碍度 O_i = (R_i × W_i) / Σ(R_j × W_j) × 100%
# 维度障碍度 = 该维度下所有指标障碍度之和

# 先对数据做极差标准化
def min_max_normalize(df, col, direction):
    """极差标准化"""
    x = df[col].values.astype(float)
    x_min, x_max = x.min(), x.max()
    if x_max == x_min:
        return np.full_like(x, 0.5)
    if direction == '正向':
        return (x - x_min) / (x_max - x_min)
    else:  # 逆向
        return (x_max - x) / (x_max - x_min)

# 对每年每省计算障碍度
all_obstacle = []

for year in years:
    yr_panel = panel_data[panel_data['年份'] == year].set_index('省份')

    for prov in provinces:
        if prov not in yr_panel.index:
            continue
        row_data = yr_panel.loc[prov]

        obstacle_indicators = []
        for ind_code, info in indicators.items():
            if ind_code not in row_data.index:
                continue

            # 获取当年该指标的全范围用于标准化
            col_data = yr_panel[ind_code].values.astype(float)
            col_min, col_max = col_data.min(), col_data.max()

            if col_max == col_min:
                normalized = 0.5
            else:
                val = row_data[ind_code]
                if info['方向'] == '正向':
                    normalized = (val - col_min) / (col_max - col_min)
                else:
                    normalized = (col_max - val) / (col_max - col_min)

            # 偏离度 (越高说明障碍越大)
            deviation = 1 - normalized

            # 权重
            w = weight_dict.get(ind_code, 0)

            obstacle_indicators.append({
                '省份': prov, '年份': year,
                '指标': ind_code, '维度': info['维度'],
                '方向': info['方向'],
                '原始值': row_data[ind_code],
                '标准化值': round(normalized, 4),
                '偏离度': round(deviation, 4),
                '权重': round(w, 6),
                '偏离度×权重': round(deviation * w, 6)
            })

        if not obstacle_indicators:
            continue

        # 计算障碍度
        sum_rw = sum(item['偏离度×权重'] for item in obstacle_indicators)
        if sum_rw > 0:
            for item in obstacle_indicators:
                item['指标障碍度(%)'] = round(item['偏离度×权重'] / sum_rw * 100, 2)

        all_obstacle.extend(obstacle_indicators)

obstacle_df = pd.DataFrame(all_obstacle)
print(f"障碍度数据: {obstacle_df.shape}")

# ============ 4. 维度障碍度汇总 ============
print("\n--- 维度障碍度汇总 ---")

dim_obstacle = obstacle_df.groupby(['省份', '年份', '维度'])['指标障碍度(%)'].sum().reset_index()
dim_obstacle.rename(columns={'指标障碍度(%)': '维度障碍度(%)'}, inplace=True)

# 2022年各省维度障碍度
dim_2022 = dim_obstacle[dim_obstacle['年份'] == 2022].pivot(
    index='省份', columns='维度', values='维度障碍度(%)'
).reset_index()
dim_2022['区域'] = dim_2022['省份'].map(prov_region)
print("2022年各省维度障碍度(前10):")
print(dim_2022.head(10).to_string(index=False))

# ============ 5. 各省Top障碍因子 ============
print("\n--- 各省Top3障碍因子 (2022年) ---")

obs_2022 = obstacle_df[obstacle_df['年份'] == 2022]
top_obstacles = []
for prov in provinces:
    prov_data = obs_2022[obs_2022['省份'] == prov].nlargest(3, '指标障碍度(%)')
    for rank, (_, row) in enumerate(prov_data.iterrows(), 1):
        top_obstacles.append({
            '省份': prov,
            '区域': prov_region.get(prov, ''),
            '排名': rank,
            '障碍指标': row['指标'],
            '维度': row['维度'],
            '障碍度(%)': row['指标障碍度(%)'],
            '偏离度': row['偏离度']
        })

top_obs_df = pd.DataFrame(top_obstacles)
print(top_obs_df.head(15).to_string(index=False))

# ============ 6. 区域障碍度对比 ============
print("\n--- 区域障碍度对比 (2022年) ---")

dim_2022_reg = dim_obstacle[dim_obstacle['年份'] == 2022].copy()
dim_2022_reg['区域'] = dim_2022_reg['省份'].map(prov_region)

reg_dim = dim_2022_reg.groupby(['区域', '维度'])['维度障碍度(%)'].mean().reset_index()
reg_dim_pivot = reg_dim.pivot(index='区域', columns='维度', values='维度障碍度(%)').reset_index()
# 确保列顺序
dim_order = ['产业振兴', '人才振兴', '文化繁荣', '生态振兴', '组织建设']
for d in dim_order:
    if d not in reg_dim_pivot.columns:
        reg_dim_pivot[d] = 0
reg_dim_pivot = reg_dim_pivot[['区域'] + dim_order]
print(reg_dim_pivot.to_string(index=False))

# ============ 7. 时序障碍度演变 ============
print("\n--- 时序障碍度演变 ---")

time_dim = dim_obstacle.groupby(['年份', '维度'])['维度障碍度(%)'].mean().reset_index()
time_dim_pivot = time_dim.pivot(index='年份', columns='维度', values='维度障碍度(%)').reset_index()
for d in dim_order:
    if d not in time_dim_pivot.columns:
        time_dim_pivot[d] = 0
time_dim_pivot = time_dim_pivot[['年份'] + dim_order]
print(time_dim_pivot.to_string(index=False))

# ============ 8. 全国Top5障碍指标(各年) ============
print("\n--- 全国Top5障碍指标(各年均值) ---")

national_obs = obstacle_df.groupby(['指标', '维度'])['指标障碍度(%)'].mean().reset_index()
national_obs = national_obs.sort_values('指标障碍度(%)', ascending=False)
print(national_obs.head(10).to_string(index=False))

# ============ 9. 保存结果 ============
print("\n--- 保存结果 ---")

output_path = '02_分析结果/Step10_障碍度诊断结果.xlsx'
with pd.ExcelWriter(output_path) as writer:
    obstacle_df.to_excel(writer, sheet_name='指标障碍度明细', index=False)
    dim_obstacle.to_excel(writer, sheet_name='维度障碍度汇总', index=False)
    top_obs_df.to_excel(writer, sheet_name='各省Top3障碍因子', index=False)
    reg_dim_pivot.to_excel(writer, sheet_name='区域障碍度对比', index=False)
    time_dim_pivot.to_excel(writer, sheet_name='时序障碍度演变', index=False)
    national_obs.to_excel(writer, sheet_name='全国障碍指标排名', index=False)

print(f"✅ 结果已保存: {output_path}")

# ============ 10. 绘图 ============
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False

# 图11: 区域维度障碍度雷达图
fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
categories = dim_order
N_cat = len(categories)
angles = [n / float(N_cat) * 2 * np.pi for n in range(N_cat)]
angles += angles[:1]

colors_reg = {'东部': '#e74c3c', '中部': '#2ecc71', '西部': '#3498db', '东北': '#f39c12'}
for reg in ['东部', '中部', '西部', '东北']:
    reg_row = reg_dim_pivot[reg_dim_pivot['区域'] == reg]
    if len(reg_row) > 0:
        values = [reg_row[d].values[0] for d in dim_order]
        values += values[:1]
        ax.plot(angles, values, 'o-', linewidth=2, label=reg, color=colors_reg[reg])
        ax.fill(angles, values, alpha=0.1, color=colors_reg[reg])

ax.set_xticks(angles[:-1])
ax.set_xticklabels(categories, fontsize=12)
ax.set_title('四大区域乡村振兴维度障碍度对比', fontsize=14, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1), fontsize=10)
plt.tight_layout()
plt.savefig('03_图表文件/图11_区域障碍度雷达图.png', dpi=200, bbox_inches='tight')
plt.close()
print("✅ 图11: 区域障碍度雷达图已保存")

# 图12: 全国Top障碍指标排名(柱状图)
fig, ax = plt.subplots(figsize=(10, 6))
top_n = national_obs.head(8)
# 简化指标名
short_names = []
for name in top_n['指标']:
    short = name.replace('X01_', '').replace('X02_', '').replace('X03_', '')\
                .replace('X04_', '').replace('X05_', '').replace('X06_', '')\
                .replace('X07_', '').replace('X08_', '').replace('X09_', '')\
                .replace('X11_', '').replace('X14_', '').replace('X15_', '')\
                .replace('X16_', '').replace('X17_', '')
    if len(short) > 15:
        short = short[:15] + '...'
    short_names.append(short)

dim_colors = {'产业振兴': '#e74c3c', '人才振兴': '#3498db', '文化繁荣': '#9b59b6',
              '生态振兴': '#2ecc71', '组织建设': '#f39c12'}
bar_colors = [dim_colors.get(d, '#95a5a6') for d in top_n['维度']]

bars = ax.barh(range(len(top_n)), top_n['指标障碍度(%)'], color=bar_colors, alpha=0.85, height=0.6)
ax.set_yticks(range(len(top_n)))
ax.set_yticklabels(short_names, fontsize=10)
ax.set_xlabel('平均障碍度(%)', fontsize=12)
ax.set_title('全国乡村振兴核心障碍指标排名', fontsize=14)
ax.invert_yaxis()

# 标注数值
for i, (val, dim) in enumerate(zip(top_n['指标障碍度(%)'], top_n['维度'])):
    ax.text(val + 0.2, i, f'{val:.2f}%', va='center', fontsize=10)

# 图例
from matplotlib.patches import Patch
legend_elements = [Patch(facecolor=c, label=d) for d, c in dim_colors.items()]
ax.legend(handles=legend_elements, loc='lower right', fontsize=9)
ax.grid(True, alpha=0.3, axis='x')
plt.tight_layout()
plt.savefig('03_图表文件/图12_全国障碍指标排名.png', dpi=200, bbox_inches='tight')
plt.close()
print("✅ 图12: 全国障碍指标排名图已保存")

# 图13: 时序障碍度堆叠面积图
fig, ax = plt.subplots(figsize=(10, 6))
time_data = time_dim_pivot.set_index('年份')[dim_order]
time_data.plot.area(ax=ax, alpha=0.7, color=[dim_colors.get(d, '#95a5a6') for d in dim_order])
ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel('平均障碍度(%)', fontsize=12)
ax.set_title('全国乡村振兴维度障碍度时序演变', fontsize=14)
ax.legend(fontsize=10, loc='upper right')
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('03_图表文件/图13_时序障碍度演变.png', dpi=200, bbox_inches='tight')
plt.close()
print("✅ 图13: 时序障碍度演变图已保存")

print("\n" + "=" * 60)
print("Step10 障碍度诊断完成!")
print("=" * 60)
