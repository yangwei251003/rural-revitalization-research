"""
Step4: 面板固定效应回归
========================
功能：
1. 双向固定效应模型（个体+时间固定效应）
2. 分析14个指标对乡村振兴综合得分的影响
3. 逐步回归：全变量 → 逐步筛选
4. 生成回归结果表和图表

被解释变量：乡村振兴TOPSIS综合得分
解释变量：14个原始指标（X01-X17，缺X10/X12/X13）
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from linearmodels.panel import PanelOLS
import warnings
warnings.filterwarnings('ignore')

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ========== 指标名称映射（简化变量名） ==========
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

# 五维度分类
dimension_vars = {
    '产业兴旺': ['X01_人均可支配收入', 'X02_农林牧渔产值', 'X03_农机总动力',
                'X04_有效灌溉占比', 'X05_化肥施用强度', 'X06_农药施用强度'],
    '人才支撑': ['X08_卫生技术人员'],
    '文化繁荣': ['X07_文教娱乐占比', 'X09_公共图书馆数'],
    '生态宜居': ['X11_森林覆盖率'],
    '组织建设': ['X14_农林水支出占比', 'X15_村委会数量', 'X16_养老保险参保', 'X17_医疗保险参保']
}

# ========== 主程序 ==========
print("="*60)
print("Step4: 面板固定效应回归")
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

# 获取所有解释变量
X_vars = list(var_names.values())
print(f"✓ 数据维度：{df.shape}")
print(f"✓ 被解释变量：TOPSIS综合得分")
print(f"✓ 解释变量({len(X_vars)}个)：{X_vars}")

# 2. 描述性统计
print("\n[2/6] 描述性统计...")
desc_stats = df[X_vars + ['TOPSIS综合得分']].describe().T
desc_stats['变异系数'] = desc_stats['std'] / desc_stats['mean']
print(desc_stats[['mean', 'std', '变异系数']].to_string())

# 3. 相关性分析
print("\n[3/6] 相关性分析...")
corr_matrix = df[X_vars + ['TOPSIS综合得分']].corr()

# 生成相关性热力图
fig, ax = plt.subplots(figsize=(14, 11))
sns.heatmap(corr_matrix, annot=True, fmt='.2f', cmap='RdBu_r', center=0,
            ax=ax, square=True, linewidths=0.5,
            annot_kws={'size': 7})
ax.set_title('图5 变量相关性热力图', fontsize=16, fontweight='bold')
plt.tight_layout()
plt.savefig('图5_相关性热力图.png', dpi=300, bbox_inches='tight')
print("✓ 已生成：图5_相关性热力图.png")
plt.close()

# 4. 设置面板数据结构
print("\n[4/6] 构建面板数据结构...")

# 为linearmodels设置多重索引（个体，时间）
df = df.set_index(['省份', '年份'])

# 标准化解释变量（减少多重共线性影响）
from sklearn.preprocessing import StandardScaler
scaler = StandardScaler()
df_standardized = df.copy()
df_standardized[X_vars] = scaler.fit_transform(df[X_vars])

# 5. 面板固定效应回归
print("\n[5/6] 面板固定效应回归...")

# ---- 模型1：全变量双向固定效应 ----
print("\n  模型1：全变量双向固定效应（标准化变量）")
y = df_standardized['TOPSIS综合得分']
X_all = df_standardized[X_vars]

model1 = PanelOLS(y, X_all, entity_effects=True, time_effects=True, drop_absorbed=True)
results1 = model1.fit(cov_type='clustered', cluster_entity=True)

print(results1.summary.tables[1])

# ---- 模型2：逐步回归（基于显著性筛选） ----
print("\n  模型2：逐步回归筛选显著变量")

# 先用全变量模型，逐步剔除不显著变量
significant_vars = X_vars.copy()
max_iter = 10

for iteration in range(max_iter):
    X_step = df_standardized[significant_vars]
    model_step = PanelOLS(y, X_step, entity_effects=True, time_effects=True, drop_absorbed=True)
    results_step = model_step.fit(cov_type='clustered', cluster_entity=True)
    
    # 找最不显著的变量
    pvalues = results_step.pvalues
    max_p = pvalues.max()
    
    if max_p > 0.1:  # 剔除p>0.1的变量
        worst_var = pvalues.idxmax()
        significant_vars.remove(worst_var)
        print(f"    第{iteration+1}轮：剔除 {worst_var} (p={max_p:.4f})")
    else:
        print(f"    第{iteration+1}轮：所有变量p<0.1，停止筛选")
        break

print(f"\n  最终保留变量({len(significant_vars)}个)：{significant_vars}")

# 最终模型
X_final = df_standardized[significant_vars]
model2 = PanelOLS(y, X_final, entity_effects=True, time_effects=True, drop_absorbed=True)
results2 = model2.fit(cov_type='clustered', cluster_entity=True)

print("\n  模型2结果：")
print(results2.summary.tables[1])

# ---- 模型3：分维度回归 ----
print("\n  模型3：分维度回归")

dim_results = {}
for dim, vars_list in dimension_vars.items():
    # 只保留存在于significant_vars中的变量
    use_vars = [v for v in vars_list if v in significant_vars]
    if len(use_vars) == 0:
        print(f"    {dim}：无显著变量，跳过")
        continue
    
    X_dim = df_standardized[use_vars]
    model_dim = PanelOLS(y, X_dim, entity_effects=True, time_effects=True, drop_absorbed=True)
    
    try:
        results_dim = model_dim.fit(cov_type='clustered', cluster_entity=True)
        
        # 检查是否所有变量都被吸收了
        if len(results_dim.params) == 0:
            print(f"    {dim}：所有变量被固定效应吸收，跳过")
            continue
        
        dim_results[dim] = {
            'model': model_dim,
            'results': results_dim,
            'vars': use_vars
        }
        
        print(f"\n    {dim} ({len(use_vars)}个变量)：")
        print(f"      R² = {results_dim.rsquared:.4f}")
    except ValueError as e:
        print(f"    {dim}：模型无法估计（{str(e)[:50]}），跳过")
        continue
    for var in use_vars:
        coef = results_dim.params[var]
        pval = results_dim.pvalues[var]
        sig = '***' if pval < 0.01 else '**' if pval < 0.05 else '*' if pval < 0.1 else ''
        print(f"      {var}: β={coef:.4f} (p={pval:.4f}) {sig}")

# 6. 生成回归系数图
print("\n[6/6] 生成回归系数可视化...")

fig, axes = plt.subplots(1, 2, figsize=(18, 8))
fig.suptitle('图6 面板固定效应回归系数', fontsize=16, fontweight='bold')

# 子图1：全变量模型系数
ax1 = axes[0]
coefs1 = results1.params
pvals1 = results1.pvalues
colors1 = ['#e74c3c' if p < 0.01 else '#f39c12' if p < 0.05 else '#3498db' if p < 0.1 else '#95a5a6' 
           for p in pvals1]
y_pos1 = range(len(coefs1))

ax1.barh(y_pos1, coefs1, color=colors1, alpha=0.8)
ax1.set_yticks(y_pos1)
ax1.set_yticklabels(coefs1.index, fontsize=9)
ax1.axvline(x=0, color='black', linestyle='-', linewidth=0.5)
ax1.set_xlabel('标准化回归系数', fontsize=11)
ax1.set_title('(a) 全变量模型系数', fontsize=13)

# 添加显著性标记
for i, (coef, pval) in enumerate(zip(coefs1, pvals1)):
    sig = '***' if pval < 0.01 else '**' if pval < 0.05 else '*' if pval < 0.1 else ''
    ax1.text(coef + 0.002 if coef >= 0 else coef - 0.002, i, sig, 
             ha='left' if coef >= 0 else 'right', va='center', fontsize=9)

# 子图2：逐步回归模型系数
ax2 = axes[1]
coefs2 = results2.params
pvals2 = results2.pvalues
colors2 = ['#e74c3c' if p < 0.01 else '#f39c12' if p < 0.05 else '#3498db' if p < 0.1 else '#95a5a6' 
           for p in pvals2]
y_pos2 = range(len(coefs2))

ax2.barh(coefs2.index, coefs2, color=colors2, alpha=0.8)
ax2.axvline(x=0, color='black', linestyle='-', linewidth=0.5)
ax2.set_xlabel('标准化回归系数', fontsize=11)
ax2.set_title('(b) 逐步回归模型系数', fontsize=13)

for i, (coef, pval) in enumerate(zip(coefs2, pvals2)):
    sig = '***' if pval < 0.01 else '**' if pval < 0.05 else '*' if pval < 0.1 else ''
    ax2.text(coef + 0.002 if coef >= 0 else coef - 0.002, i, sig, 
             ha='left' if coef >= 0 else 'right', va='center', fontsize=9)

# 添加图例
from matplotlib.patches import Patch
legend_elements = [
    Patch(facecolor='#e74c3c', label='p<0.01'),
    Patch(facecolor='#f39c12', label='p<0.05'),
    Patch(facecolor='#3498db', label='p<0.1'),
    Patch(facecolor='#95a5a6', label='p≥0.1')
]
fig.legend(handles=legend_elements, loc='lower center', ncol=4, fontsize=10, 
           bbox_to_anchor=(0.5, -0.02))

plt.tight_layout()
plt.savefig('图6_面板回归系数.png', dpi=300, bbox_inches='tight')
print("✓ 已生成：图6_面板回归系数.png")
plt.close()

# 7. 保存结果到Excel
print("\n保存结果到Excel...")

with pd.ExcelWriter('Step4_面板固定效应回归结果.xlsx', engine='openpyxl') as writer:
    # Sheet1: 描述性统计
    desc_stats.to_excel(writer, sheet_name='描述性统计')
    
    # Sheet2: 模型1全变量回归结果
    model1_df = pd.DataFrame({
        '变量': results1.params.index,
        '系数': results1.params.values,
        '标准误': results1.std_errors.values,
        't统计量': results1.tstats.values,
        'p值': results1.pvalues.values,
        '显著性': ['***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else '' 
                   for p in results1.pvalues.values]
    })
    model1_df.to_excel(writer, sheet_name='全变量回归', index=False)
    
    # Sheet3: 模型2逐步回归结果
    model2_df = pd.DataFrame({
        '变量': results2.params.index,
        '系数': results2.params.values,
        '标准误': results2.std_errors.values,
        't统计量': results2.tstats.values,
        'p值': results2.pvalues.values,
        '显著性': ['***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else '' 
                   for p in results2.pvalues.values]
    })
    model2_df.to_excel(writer, sheet_name='逐步回归', index=False)
    
    # Sheet4: 分维度回归结果
    all_dim_results = []
    for dim, info in dim_results.items():
        res = info['results']
        for var in info['vars']:
            all_dim_results.append({
                '维度': dim,
                '变量': var,
                '系数': res.params[var],
                '标准误': res.std_errors[var],
                't统计量': res.tstats[var],
                'p值': res.pvalues[var],
                '显著性': '***' if res.pvalues[var] < 0.01 else '**' if res.pvalues[var] < 0.05 else '*' if res.pvalues[var] < 0.1 else ''
            })
    if all_dim_results:
        pd.DataFrame(all_dim_results).to_excel(writer, sheet_name='分维度回归', index=False)
    
    # Sheet5: 模型拟合统计
    adj_r2_1 = getattr(results1, 'rsquared_adj', results1.rsquared_between if hasattr(results1, 'rsquared_between') else None)
    adj_r2_2 = getattr(results2, 'rsquared_adj', results2.rsquared_between if hasattr(results2, 'rsquared_between') else None)
    fit_stats = pd.DataFrame({
        '统计量': ['R²', 'Adj.R²', 'F统计量', '观测数', '组数'],
        '模型1_全变量': [results1.rsquared, adj_r2_1, 
                      results1.f_statistic.stat, results1.nobs, results1.entity_info.total],
        '模型2_逐步回归': [results2.rsquared, adj_r2_2, 
                       results2.f_statistic.stat, results2.nobs, results2.entity_info.total]
    })
    fit_stats.to_excel(writer, sheet_name='模型拟合统计', index=False)
    
    # Sheet6: 相关性矩阵
    corr_matrix.to_excel(writer, sheet_name='相关性矩阵')

print("✓ 已保存：Step4_面板固定效应回归结果.xlsx")

# 8. 输出核心结论
print("\n" + "="*60)
print("核心结论：")
print("="*60)

print(f"\n【模型1：全变量双向固定效应】")
print(f"  R² = {results1.rsquared:.4f}")
print(f"  显著变量（p<0.05）：")
for var in results1.params.index:
    p = results1.pvalues[var]
    if p < 0.05:
        coef = results1.params[var]
        print(f"    {var}: β={coef:.4f} (p={p:.4f}) {'正向↑' if coef > 0 else '负向↓'}")

print(f"\n【模型2：逐步回归（{len(significant_vars)}个变量）】")
print(f"  R² = {results2.rsquared:.4f}")
print(f"  保留变量：")
for var in results2.params.index:  # 只遍历模型中实际保留的变量
    coef = results2.params[var]
    p = results2.pvalues[var]
    sig = '***' if p < 0.01 else '**' if p < 0.05 else '*'
    print(f"    {var}: β={coef:.4f} (p={p:.4f}) {sig} {'正向↑' if coef > 0 else '负向↓'}")

# 检查被吸收的变量
absorbed_vars = [v for v in significant_vars if v not in results2.params.index]
if absorbed_vars:
    print(f"  被固定效应吸收的变量：{absorbed_vars}")

print("\n" + "="*60)
print("Step4完成！")
print("="*60)
