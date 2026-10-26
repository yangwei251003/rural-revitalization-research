"""
Step7.1: 替换被解释变量再检验
用等权TOPSIS法、主成分分析法重新计算乡村振兴得分，然后进行面板回归，检验稳健性
"""
import pandas as pd
import numpy as np
from sklearn.decomposition import PCA
from linearmodels.panel import PanelOLS
import warnings
warnings.filterwarnings('ignore')

print("=" * 60)
print("Step7.1: 替换被解释变量再检验")
print("=" * 60)

# 1. 读取面板数据
df = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')
print(f"✅ 读取面板数据: {df.shape[0]}行 × {df.shape[1]}列")

# 2. 获取指标列（X01-X17，除了被剔除的）
indicators = [col for col in df.columns if col.startswith('X')]
print(f"   指标数: {len(indicators)}")
print(f"   指标列表: {indicators[:5]}...")

# 3. 等权TOPSIS法计算得分
print(f"\n{'='*60}")
print("  方法1: 等权TOPSIS法")
print(f"{'='*60}")

def topsis_equal_weight(df, indicators):
    """等权TOPSIS法：所有指标权重相等"""
    # 1. 数据标准化（极差标准化）
    df_norm = df[indicators].copy()
    for col in indicators:
        if '逆向' in col or '强度' in col:  # 逆向指标
            df_norm[col] = (df[col].max() - df[col]) / (df[col].max() - df[col].min())
        else:  # 正向指标
            df_norm[col] = (df[col] - df[col].min()) / (df[col].max() - df[col].min())

    # 2. 等权加权（所有指标权重=1/len(indicators)）
    weights = np.ones(len(indicators)) / len(indicators)
    for i, col in enumerate(indicators):
        df_norm[col] = df_norm[col] * weights[i]

    # 3. 确定正理想解和负理想解
    positive_ideal = df_norm.max()
    negative_ideal = df_norm.min()

    # 4. 计算距离
    d_pos = np.sqrt(((df_norm - positive_ideal) ** 2).sum(axis=1))
    d_neg = np.sqrt(((df_norm - negative_ideal) ** 2).sum(axis=1))

    # 5. 计算贴近度
    score = d_neg / (d_pos + d_neg)
    return score

df['等权TOPSIS得分'] = topsis_equal_weight(df, indicators)
print(f"  ✅ 等权TOPSIS得分计算完成")
print(f"     得分范围: [{df['等权TOPSIS得分'].min():.4f}, {df['等权TOPSIS得分'].max():.4f}]")

# 4. 主成分分析法计算得分
print(f"\n{'='*60}")
print("  方法2: 主成分分析法(PCA)")
print(f"{'='*60}")

# 标准化数据
df_pca = df[indicators].copy()
for col in indicators:
    if '逆向' in col or '强度' in col:  # 逆向指标，取负数使其正向化
        df_pca[col] = - (df[col] - df[col].min()) / (df[col].max() - df[col].min())
    else:  # 正向指标
        df_pca[col] = (df[col] - df[col].min()) / (df[col].max() - df[col].min())

# PCA
pca = PCA()
pca_result = pca.fit_transform(df_pca.fillna(0))
# 用第一主成分作为综合得分（或者按方差贡献率加权）
explained_var = pca.explained_variance_ratio_
print(f"  第一主成分方差贡献率: {explained_var[0]:.4f}")
print(f"  前三个主成分累计方差贡献率: {sum(explained_var[:3]):.4f}")

# 用第一主成分作为综合得分
df['PCA综合得分'] = pca_result[:, 0]
# 标准化到[0,1]区间，方便比较
df['PCA综合得分'] = (df['PCA综合得分'] - df['PCA综合得分'].min()) / (df['PCA综合得分'].max() - df['PCA综合得分'].min())
print(f"  ✅ PCA综合得分计算完成")
print(f"     得分范围: [{df['PCA综合得分'].min():.4f}, {df['PCA综合得分'].max():.4f}]")

# 5. 读取Step1的基准得分（熵权TOPSIS）
df_score = pd.read_excel('Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')
df_score = df_score.rename(columns={'TOPSIS综合得分': '熵权TOPSIS得分'})
print(f"\n✅ 读取基准得分（熵权TOPSIS）: {df_score.shape[0]}行")

# 6. 合并所有得分
df_all = df.merge(df_score[['省份', '年份', '熵权TOPSIS得分']], on=['省份', '年份'], how='left')

# 7. 构造解释变量（与Step4/Step6相同）
df_all['经济发展水平'] = np.log(df_all['X01_农村居民人均可支配收入(元)'] + 1)
df_all['财政支农力度'] = df_all['X14_农林水支出占财政支出比重(%)']
df_all['人力资本水平'] = df_all['X07_文教娱乐支出占消费比重(%)']
df_all['数字基础设施'] = np.log(df_all['X09_公共图书馆数量(个)'] + 1)
df_all['产业结构'] = df_all['X02_农林牧渔业总产值(亿元)'] / df_all['X02_农林牧渔业总产值(亿元)'].max()
df_all['农业现代化'] = df_all['X03_农业机械总动力(万千瓦)'] / df_all['X03_农业机械总动力(万千瓦)'].max()
df_all['交通基础设施'] = df_all['X09_公共图书馆数量(个)'] / df_all['X09_公共图书馆数量(个)'].max()

# 8. 分别用三种得分进行面板回归
print(f"\n{'='*60}")
print("  面板回归（双向固定效应）")
print(f"{'='*60}")

dependent_vars = ['熵权TOPSIS得分', '等权TOPSIS得分', 'PCA综合得分']
results = {}

for dep_var in dependent_vars:
    print(f"\n 被解释变量: {dep_var}")

    # 清理数据
    df_clean = df_all.dropna(subset=[dep_var, '经济发展水平', '财政支农力度',
                                    '人力资本水平', '数字基础设施'])
    df_clean = df_clean.set_index(['省份', '年份'])

    # 解释变量
    X = pd.DataFrame({
        '经济发展水平': df_clean['经济发展水平'],
        '财政支农力度': df_clean['财政支农力度'],
        '人力资本水平': df_clean['人力资本水平'],
        '数字基础设施': df_clean['数字基础设施'],
        '产业结构': df_clean['产业结构'],
        '农业现代化': df_clean['农业现代化'],
        '交通基础设施': df_clean['交通基础设施'],
    })

    y = df_clean[dep_var]

    # 回归
    model = PanelOLS(y, X, entity_effects=True, time_effects=True, drop_absorbed=True)
    result = model.fit(cov_type='clustered', cluster_entity=True)

    results[dep_var] = {
        'result': result,
        'nobs': result.nobs,
        'rsquared': result.rsquared,
        'params': result.params,
        'pvalues': result.pvalues,
        'std_errors': result.std_errors
    }

    print(f"    R² = {result.rsquared:.4f}")
    print(f"    观测值数: {result.nobs}")
    print(f"    显著变量数: {sum(result.pvalues < 0.1)}/7")
    print(f"    系数符号一致性检查:")
    base_params = results[dependent_vars[0]]['params'] if dep_var != dependent_vars[0] else None
    for var in result.params.index:
        coef = result.params[var]
        p = result.pvalues[var]
        stars = '***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else ''
        print(f"      {var}: {coef:8.4f}{stars}", end='')
        if base_params is not None:
            direction = '✓' if (coef * base_params[var]) > 0 else '✗'
            print(f"  {direction}", end='')
        print()

# 9. 输出结果到Excel
print(f"\n{'='*60}")
print("  输出结果到Excel...")
print(f"{'='*60}")

with pd.ExcelWriter('Step7_1_替换被解释变量检验.xlsx', engine='openpyxl') as writer:

    # Sheet1: 三种方法结果对比
    comparison_data = []
    base_result = results[dependent_vars[0]]

    for var in base_result['params'].index:
        row = {'变量': var}
        for dep_var in dependent_vars:
            r = results[dep_var]
            row[f'{dep_var}_系数'] = round(r['params'][var], 4)
            row[f'{dep_var}_p值'] = round(r['pvalues'][var], 4)
            stars = '***' if r['pvalues'][var] < 0.01 else '**' if r['pvalues'][var] < 0.05 else '*' if r['pvalues'][var] < 0.1 else ''
            row[f'{dep_var}_显著性'] = stars
        comparison_data.append(row)

    df_comp = pd.DataFrame(comparison_data)
    df_comp.to_excel(writer, sheet_name='三种方法结果对比', index=False)
    print(f"  ✅ Sheet1: 三种方法结果对比")

    # Sheet2-4: 各方法详细结果
    for dep_var in dependent_vars:
        r = results[dep_var]
        detail = pd.DataFrame({
            '变量': r['params'].index,
            '系数': r['params'].values,
            '标准误': r['std_errors'].values,
            't值': r['params'].values / r['std_errors'].values,
            'p值': r['pvalues'].values,
            '显著性': ['***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else '' for p in r['pvalues'].values]
        })
        # 添加R²和观测值数
        detail['R2'] = r['rsquared']
        detail['观测值数'] = r['nobs']
        detail.to_excel(writer, sheet_name=dep_var[:20], index=False)  # Excel sheet名最多31字符
        print(f"  ✅ Sheet: {dep_var[:20]}")

print("\n✅ Step7.1 替换被解释变量再检验完成！")
print("   结果已保存到: Step7_1_替换被解释变量检验.xlsx")
print(f"\n{'='*60}")
print("  稳健性检验结论：")
print("  1. 等权TOPSIS法与熵权TOPSIS法结果对比：")
print("     - 系数符号一致性：检查上方输出中的✓/✗")
print("     - R²对比：熵权法通常更高（因为考虑了指标区分度）")
print("  2. PCA与主成分分析：")
print("     - 第一主成分捕获最大方差，作为综合得分")
print("     - 与TOPSIS结果应该方向一致")
print(f"{'='*60}")
