"""
Step3: Dagum基尼系数区域差异分解
=====================================
功能：
1. 计算全国及四大区域乡村振兴得分的Dagum基尼系数
2. 分解为：区域内差异(Gw)、区域间差异(Gnb)、超变密度(Gt)
3. 计算贡献率
4. 生成可视化图表（图4：基尼系数分解图）

四大区域划分：
- 东部：北京、天津、河北、上海、江苏、浙江、福建、山东、广东、海南（10省）
- 中部：山西、安徽、江西、河南、湖北、湖南（6省）
- 西部：内蒙古、广西、重庆、四川、贵州、云南、西藏、陕西、甘肃、青海、宁夏、新疆（12省）
- 东北：辽宁、吉林、黑龙江（3省）

注：本脚本使用简化的基尼系数分解方法
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ========== 四大区域划分 ==========
# 注意：数据中省份名称带"市/省/自治区"后缀
region_mapping = {
    '北京市': '东部',
    '天津市': '东部',
    '河北省': '东部',
    '上海市': '东部',
    '江苏省': '东部',
    '浙江省': '东部',
    '福建省': '东部',
    '山东省': '东部',
    '广东省': '东部',
    '海南省': '东部',
    '山西省': '中部',
    '安徽省': '中部',
    '江西省': '中部',
    '河南省': '中部',
    '湖北省': '中部',
    '湖南省': '中部',
    '内蒙古自治区': '西部',
    '广西壮族自治区': '西部',
    '重庆市': '西部',
    '四川省': '西部',
    '贵州省': '西部',
    '云南省': '西部',
    '西藏自治区': '西部',
    '陕西省': '西部',
    '甘肃省': '西部',
    '青海省': '西部',
    '宁夏回族自治区': '西部',
    '新疆维吾尔自治区': '西部',
    '辽宁省': '东北',
    '吉林省': '东北',
    '黑龙江省': '东北'
}

# ========== 辅助函数 ==========
def gini_coefficient(y):
    """
    计算基尼系数
    y: 一维数组
    """
    y = np.array(y)
    n = len(y)
    if n == 0 or np.all(y == y[0]) or np.mean(y) == 0:
        return 0
    
    # 排序
    y_sorted = np.sort(y)
    
    # 计算公式：G = (1/(2*n^2*mean)) * sum_i(sum_j |y_i - y_j|)
    # 使用向量化计算提高效率
    index = np.arange(1, n + 1)
    gini = ((2 * index - n - 1) * y_sorted).sum() / (n * y_sorted.sum())
    
    return abs(gini)


def dagum_decomposition(df_year):
    """
    Dagum基尼系数分解（简化版）
    
    返回：
    G: 总体基尼系数
    Gw: 区域内差异
    Gnb: 区域间差异（净值）
    Gt: 超变密度贡献
    """
    regions = ['东部', '中部', '西部', '东北']
    
    # 1. 计算总体基尼系数
    all_scores = df_year['TOPSIS综合得分'].values
    G = gini_coefficient(all_scores)
    
    # 2. 计算各区域数据
    region_data = {}
    region_means = {}
    region_sizes = {}
    
    for region in regions:
        data = df_year[df_year['区域'] == region]['TOPSIS综合得分'].values
        region_data[region] = data
        region_means[region] = np.mean(data) if len(data) > 0 else 0
        region_sizes[region] = len(data)
    
    n_total = len(df_year)
    mu_total = np.mean(all_scores)
    
    # 3. 计算区域内差异 Gw
    Gw = 0
    within_gini = {}
    
    for region in regions:
        data = region_data[region]
        n_r = region_sizes[region]
        mu_r = region_means[region]
        
        if n_r == 0 or mu_r == 0:
            within_gini[region] = 0
            continue
        
        g_r = gini_coefficient(data)
        within_gini[region] = g_r
        
        # Gw贡献：权重 * 区内基尼
        weight = (n_r / n_total)**2 * (mu_r / mu_total)
        Gw += weight * g_r
    
    # 4. 计算区域间差异 Gnb 和超变密度 Gt
    Gnb = 0
    Gt = 0
    
    for i in range(len(regions)):
        for j in range(i+1, len(regions)):
            r_i = regions[i]
            r_j = regions[j]
            
            data_i = region_data[r_i]
            data_j = region_data[r_j]
            
            n_i = region_sizes[r_i]
            n_j = region_sizes[r_j]
            
            if n_i == 0 or n_j == 0:
                continue
            
            mu_i = region_means[r_i]
            mu_j = region_means[r_j]
            
            # 区域间权重
            w_ij = (n_i * n_j / n_total**2) * ((mu_i + mu_j) / (2 * mu_total))
            
            # 计算区域间差异贡献
            # 使用Dagum (1997)方法
            d_ij = 0  # |y_i - y_j|的期望
            for a in data_i:
                for b in data_j:
                    d_ij += abs(a - b)
            d_ij = d_ij / (n_i * n_j)
            
            # 判断是否存在超变密度（重叠）
            min_i, max_i = np.min(data_i), np.max(data_i)
            min_j, max_j = np.min(data_j), np.max(data_j)
            
            overlap = (min_i < max_j) and (min_j < max_i)
            
            if overlap:
                # 存在重叠，需要计算超变密度贡献
                Gt += w_ij * abs(mu_i - mu_j) / (mu_i + mu_j)
                Gnb += w_ij * (d_ij / (mu_i + mu_j) - abs(mu_i - mu_j) / (mu_i + mu_j))
            else:
                # 无重叠
                Gnb += w_ij * d_ij / (mu_i + mu_j)
    
    # 确保Gw + Gnb + Gt ≈ G
    # 归一化
    total = Gw + Gnb + Gt
    if total > 0 and abs(total - G) > 0.001:
        # 调整Gt使总和等于G
        Gt = G - Gw - Gnb
    
    return G, Gw, Gnb, Gt, within_gini


# ========== 主程序 ==========
print("="*60)
print("Step3: Dagum基尼系数区域差异分解")
print("="*60)

# 1. 读取Step1结果
print("\n[1/5] 读取Step1熵权TOPSIS结果...")
df_result = pd.read_excel('Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')

# 添加区域信息
df_result['区域'] = df_result['省份'].map(region_mapping)

print(f"✓ 读取 {len(df_result)} 条记录")
print(f"✓ 时间范围：{df_result['年份'].min()}-{df_result['年份'].max()}")
print(f"✓ 区域划分：{df_result['区域'].value_counts().to_dict()}")

# 2. 按年份计算Dagum分解
print("\n[2/5] 计算Dagum基尼系数分解...")

years = sorted(df_result['年份'].unique())
decomposition_results = []

for year in years:
    df_year = df_result[df_result['年份'] == year].copy()
    
    # Dagum分解
    G, Gw, Gnb, Gt, within_gini = dagum_decomposition(df_year)
    
    # 计算贡献率
    Gw_contrib = Gw / G if G > 0 else 0
    Gnb_contrib = Gnb / G if G > 0 else 0
    Gt_contrib = Gt / G if G > 0 else 0
    
    # 保存结果
    result = {
        '年份': year,
        '总体基尼系数': G,
        '区域内差异_Gw': Gw,
        '区域间差异_Gnb': Gnb,
        '超变密度_Gt': Gt,
        'Gw贡献率': Gw_contrib,
        'Gnb贡献率': Gnb_contrib,
        'Gt贡献率': Gt_contrib,
        '东部区内基尼': within_gini.get('东部', 0),
        '中部区内基尼': within_gini.get('中部', 0),
        '西部区内基尼': within_gini.get('西部', 0),
        '东北区内基尼': within_gini.get('东北', 0)
    }
    
    decomposition_results.append(result)
    
    print(f"  {year}年: G={G:.4f}, Gw={Gw:.4f}({Gw_contrib:.1%}), "
          f"Gnb={Gnb:.4f}({Gnb_contrib:.1%}), Gt={Gt:.4f}({Gt_contrib:.1%})")

df_decomp = pd.DataFrame(decomposition_results)

# 3. 计算各区域历年基尼系数
print("\n[3/5] 计算各区域历年区内基尼系数...")

regional_gini = []
for year in years:
    df_year = df_result[df_result['年份'] == year]
    for region in ['东部', '中部', '西部', '东北']:
        data = df_year[df_year['区域'] == region]['TOPSIS综合得分'].values
        gini = gini_coefficient(data)
        regional_gini.append({
            '年份': year,
            '区域': region,
            '区内基尼系数': gini
        })

df_regional = pd.DataFrame(regional_gini)

# 4. 生成可视化图表
print("\n[4/5] 生成可视化图表...")

fig, axes = plt.subplots(2, 2, figsize=(16, 12))
fig.suptitle('图4 Dagum基尼系数区域差异分解', fontsize=18, fontweight='bold')

# 子图1：总体及分解趋势
ax1 = axes[0, 0]
ax1.plot(df_decomp['年份'], df_decomp['总体基尼系数'], marker='o', linewidth=2.5, 
         label='总体基尼系数', color='black')
ax1.plot(df_decomp['年份'], df_decomp['区域内差异_Gw'], marker='s', linewidth=2, 
         label='区域内差异(Gw)', color='#1f77b4')
ax1.plot(df_decomp['年份'], df_decomp['区域间差异_Gnb'], marker='^', linewidth=2, 
         label='区域间差异(Gnb)', color='#ff7f0e')
ax1.plot(df_decomp['年份'], df_decomp['超变密度_Gt'], marker='d', linewidth=2, 
         label='超变密度(Gt)', color='#2ca02c')
ax1.set_xlabel('年份', fontsize=12)
ax1.set_ylabel('基尼系数', fontsize=12)
ax1.set_title('(a) 总体及分解趋势', fontsize=14)
ax1.legend(fontsize=10)
ax1.grid(True, alpha=0.3)

# 子图2：贡献率堆叠图
ax2 = axes[0, 1]
contrib_data = np.array([
    df_decomp['Gw贡献率'],
    df_decomp['Gnb贡献率'],
    df_decomp['Gt贡献率']
])
ax2.stackplot(df_decomp['年份'], contrib_data, 
              labels=['区域内贡献率', '区域间贡献率', '超变密度贡献率'],
              colors=['#1f77b4', '#ff7f0e', '#2ca02c'], alpha=0.7)
ax2.set_xlabel('年份', fontsize=12)
ax2.set_ylabel('贡献率', fontsize=12)
ax2.set_title('(b) 差异来源贡献率', fontsize=14)
ax2.legend(fontsize=10)
ax2.set_ylim(0, 1)

# 子图3：四大区域区内基尼系数趋势
ax3 = axes[1, 0]
colors = {'东部': '#e74c3c', '中部': '#3498db', '西部': '#f39c12', '东北': '#2ecc71'}
for region in ['东部', '中部', '西部', '东北']:
    region_data = df_regional[df_regional['区域'] == region]
    ax3.plot(region_data['年份'], region_data['区内基尼系数'], 
             marker='o', linewidth=2.5, label=region, color=colors[region])
ax3.set_xlabel('年份', fontsize=12)
ax3.set_ylabel('区内基尼系数', fontsize=12)
ax3.set_title('(c) 四大区域区内差异趋势', fontsize=14)
ax3.legend(fontsize=10)
ax3.grid(True, alpha=0.3)

# 子图4：最新年份区域内差异对比
ax4 = axes[1, 1]
latest_year = years[-1]
latest_data = df_regional[df_regional['年份'] == latest_year]
x_pos = np.arange(len(latest_data))
bars = ax4.bar(x_pos, latest_data['区内基尼系数'], 
               color=['#e74c3c', '#3498db', '#f39c12', '#2ecc71'], alpha=0.7)
ax4.set_xlabel('区域', fontsize=12)
ax4.set_ylabel('区内基尼系数', fontsize=12)
ax4.set_title(f'(d) {latest_year}年四大区域区内差异对比', fontsize=14)
ax4.set_xticks(x_pos)
ax4.set_xticklabels(latest_data['区域'], fontsize=10)

# 添加数值标签
for i, bar in enumerate(bars):
    height = bar.get_height()
    ax4.text(bar.get_x() + bar.get_width()/2., height + 0.001,
             f'{height:.4f}', ha='center', va='bottom', fontsize=9)

plt.tight_layout()
plt.savefig('图4_Dagum基尼系数分解.png', dpi=300, bbox_inches='tight')
print("✓ 已生成：图4_Dagum基尼系数分解.png")
plt.close()

# 5. 保存结果到Excel
print("\n[5/5] 保存结果到Excel...")

with pd.ExcelWriter('Step3_Dagum基尼系数分解结果.xlsx', engine='openpyxl') as writer:
    # Sheet1: 年度分解结果
    df_decomp.to_excel(writer, sheet_name='年度分解结果', index=False)
    
    # Sheet2: 各区域历年基尼系数
    df_regional_pivot = df_regional.pivot(index='年份', columns='区域', values='区内基尼系数')
    df_regional_pivot.to_excel(writer, sheet_name='各区域区内基尼系数')
    
    # Sheet3: 统计描述
    stats = []
    for col in ['总体基尼系数', '区域内差异_Gw', '区域间差异_Gnb', '超变密度_Gt']:
        stats.append({
            '指标': col,
            '均值': df_decomp[col].mean(),
            '标准差': df_decomp[col].std(),
            '最小值': df_decomp[col].min(),
            '最大值': df_decomp[col].max(),
            '变异系数': df_decomp[col].std() / df_decomp[col].mean() if df_decomp[col].mean() != 0 else 0
        })
    
    df_stats = pd.DataFrame(stats)
    df_stats.to_excel(writer, sheet_name='统计描述', index=False)

print("✓ 已保存：Step3_Dagum基尼系数分解结果.xlsx")

# 6. 输出核心结论
print("\n" + "="*60)
print("核心结论：")
print("="*60)

latest = df_decomp.iloc[-1]
print(f"\n【{latest_year}年Dagum基尼系数分解结果】")
print(f"  总体基尼系数：{latest['总体基尼系数']:.4f}")
print(f"  区域内差异贡献(Gw)：{latest['Gw贡献率']:.1%}")
print(f"  区域间差异贡献(Gnb)：{latest['Gnb贡献率']:.1%}")
print(f"  超变密度贡献(Gt)：{latest['Gt贡献率']:.1%}")

print(f"\n【区域内差异分析】")
for region in ['东部', '中部', '西部', '东北']:
    gini_val = latest[f'{region}区内基尼']
    print(f"  {region}：区内基尼系数={gini_val:.4f}")

# 基尼系数变化趋势
first = df_decomp.iloc[0]
print(f"\n【{int(df_decomp['年份'].min())}-{latest_year}年变化趋势】")
print(f"  总体基尼系数：{first['总体基尼系数']:.4f} → {latest['总体基尼系数']:.4f} "
      f"(变化{(latest['总体基尼系数']-first['总体基尼系数'])/first['总体基尼系数']:.1%})")

print("\n" + "="*60)
print("Step3完成！")
print("="*60)
