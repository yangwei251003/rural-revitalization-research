"""
第一步：熵权TOPSIS —— 数据标准化 + 熵权客观赋权 + TOPSIS综合得分
输出：指标权重表、各省各年综合得分、五大维度分维度得分
"""
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# ============================================================
# 0. 读取面板数据
# ============================================================
panel = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')
ind_cols = panel.columns[2:].tolist()  # 14个指标
provinces = panel['省份'].unique().tolist()
years = sorted(panel['年份'].unique())

print(f"面板数据：{panel.shape[0]}行 × {len(ind_cols)}指标")
print(f"省份数：{len(provinces)}，年份：{years[0]}—{years[-1]}")
print(f"缺失值检查：{panel[ind_cols].isna().sum().sum()}个\n")

# ============================================================
# 1. 指标方向定义
# ============================================================
# 正向：越大越好；逆向：越小越好
positive = [
    'X01_农村居民人均可支配收入(元)',
    'X02_农林牧渔业总产值(亿元)',
    'X03_农业机械总动力(万千瓦)',
    'X04_有效灌溉面积占比(%)',
    'X07_文教娱乐支出占消费比重(%)',
    'X08_每万乡村人口卫生技术人员数(人)',
    'X09_公共图书馆数量(个)',
    'X11_森林覆盖率(%)',
    'X14_农林水支出占财政支出比重(%)',
    'X15_村民委员会数量(个)',
    'X16_城乡居民养老保险参保人数(万人)',
    'X17_城乡居民医疗保险参保人数(万人)',
]
negative = [
    'X05_化肥施用强度(万吨/千公顷)',
    'X06_农药施用强度(万吨/千公顷)',
]
assert len(positive) + len(negative) == len(ind_cols)

direction = {c: '正向' for c in positive}
direction.update({c: '逆向' for c in negative})

# ============================================================
# 2. 极差标准化（min-max normalization）
# ============================================================
def normalize(df, direction_map):
    """对面板数据做极差标准化，返回标准化矩阵（DataFrame）"""
    norm = df.copy()
    for col in df.columns:
        x_min = df[col].min()
        x_max = df[col].max()
        if x_max == x_min:
            norm[col] = 1.0
        elif direction_map[col] == '正向':
            norm[col] = (df[col] - x_min) / (x_max - x_min)
        else:  # 逆向
            norm[col] = (x_max - df[col]) / (x_max - x_min)
    # 防止0值，加微小偏移（熵权法要求 > 0）
    norm = norm.replace(0, 1e-10)
    return norm

norm_df = normalize(panel[ind_cols], direction)
print("标准化完成。检验：各指标取值范围")
print(norm_df.describe().loc[['min','max']].round(4).to_string())
print()

# ============================================================
# 3. 熵权法计算客观权重
# ============================================================
def entropy_weight(norm_df):
    """
    熵权法计算指标权重
    步骤：比重矩阵 → 信息熵 → 差异系数 → 权重
    """
    n, m = norm_df.shape
    # 3.1 计算指标比重 p_ij = x_ij / sum_i(x_ij)
    p = norm_df.div(norm_df.sum(axis=0), axis=1)
    # 3.2 计算第j项指标的信息熵 e_j = -k * sum(p_ij * ln(p_ij))
    k = 1 / np.log(n)
    # 处理 p*ln(p)，当 p=0 时极限为0
    p_safe = p.copy()
    p_safe[p_safe == 0] = 1e-10
    ent = -k * (p_safe * np.log(p_safe)).sum(axis=0)
    # 3.3 差异系数 d_j = 1 - e_j
    d = 1 - ent
    # 3.4 权重 w_j = d_j / sum(d_j)
    w = d / d.sum()
    return w, ent, d

weights, entropy_vals, diff_coeff = entropy_weight(norm_df)

# 整理权重结果
weight_df = pd.DataFrame({
    '指标代码': ind_cols,
    '指标方向': [direction[c] for c in ind_cols],
    '信息熵': entropy_vals.round(6),
    '差异系数': diff_coeff.round(6),
    '熵权权重': weights.round(6),
})
weight_df['权重(%)'] = (weights * 100).round(4)

print("=== 熵权法权重结果 ===")
print(weight_df.to_string(index=False))
print(f"\n权重合计：{weights.sum():.6f}")
print()

# 按维度汇总权重
dim_map = {
    'X01': '产业振兴', 'X02': '产业振兴', 'X03': '产业振兴',
    'X04': '产业振兴', 'X05': '产业振兴', 'X06': '产业振兴',
    'X07': '文化繁荣', 'X08': '人才振兴',
    'X09': '文化繁荣',
    'X11': '生态振兴',
    'X14': '组织建设', 'X15': '组织建设', 'X16': '组织建设', 'X17': '组织建设',
}
weight_df['维度'] = weight_df['指标代码'].str[:3].map(dim_map)
dim_weight = weight_df.groupby('维度')['熵权权重'].sum().sort_values(ascending=False)
print("=== 各维度权重合计 ===")
for dim, w in dim_weight.items():
    print(f"  {dim}：{w:.6f}（{w*100:.2f}%）")
print()

# ============================================================
# 4. TOPSIS 计算综合贴近度得分
# ============================================================
def topsis(norm_df, weights):
    """
    TOPSIS算法
    步骤：加权标准化矩阵 → 正理想解/负理想解 → 距离 → 贴近度
    """
    # 4.1 构建加权标准化决策矩阵
    V = norm_df.mul(weights.values, axis=1)
    # 4.2 确定正理想解(A+)和负理想解(A-)
    pos_ideal = {}
    neg_ideal = {}
    for col in norm_df.columns:
        if direction[col] == '正向':
            pos_ideal[col] = V[col].max()
            neg_ideal[col] = V[col].min()
        else:
            pos_ideal[col] = V[col].min()
            neg_ideal[col] = V[col].max()
    # 4.3 计算各评价对象到正/负理想解的距离
    d_pos = np.sqrt(((V - pd.Series(pos_ideal))**2).sum(axis=1))
    d_neg = np.sqrt(((V - pd.Series(neg_ideal))**2).sum(axis=1))
    # 4.4 计算相对贴近度 C_i = d_neg / (d_pos + d_neg)
    score = d_neg / (d_pos + d_neg)
    return score, pos_ideal, neg_ideal, d_pos, d_neg

panel = panel.copy()
score, pos_ideal, neg_ideal, d_pos, d_neg = topsis(norm_df, weights)
panel['TOPSIS综合得分'] = score.values

print("=== TOPSIS综合得分 ===")
print(f"得分范围：{score.min():.6f} — {score.max():.6f}")
print(f"平均得分：{score.mean():.6f}")
print()

# ============================================================
# 5. 五大维度分维度TOPSIS得分
# ============================================================
# 每个维度内部再做一次熵权 + TOPSIS（子维度客观赋权）
dim_cols_map = {
    '产业振兴': ['X01_农村居民人均可支配收入(元)','X02_农林牧渔业总产值(亿元)','X03_农业机械总动力(万千瓦)','X04_有效灌溉面积占比(%)','X05_化肥施用强度(万吨/千公顷)','X06_农药施用强度(万吨/千公顷)'],
    '人才振兴': ['X08_每万乡村人口卫生技术人员数(人)'],
    '文化繁荣': ['X07_文教娱乐支出占消费比重(%)', 'X09_公共图书馆数量(个)'],
    '生态振兴': ['X11_森林覆盖率(%)'],
    '组织建设': ['X14_农林水支出占财政支出比重(%)','X15_村民委员会数量(个)','X16_城乡居民养老保险参保人数(万人)','X17_城乡居民医疗保险参保人数(万人)'],
}

print("=== 分维度TOPSIS得分计算 ===")
dim_score_cols = []

for dim_name, dim_inds in dim_cols_map.items():
    # 该维度的标准化数据
    nd_norm = norm_df[dim_inds]
    # 该维度指标用熵权法赋权（子维度客观权重）
    nd_p = nd_norm.div(nd_norm.sum(axis=0), axis=1)
    nd_p_safe = nd_p.copy().replace(0, 1e-10)
    k_nd = 1 / np.log(len(nd_norm))
    ent_nd = -k_nd * (nd_p_safe * np.log(nd_p_safe)).sum(axis=0)
    d_nd = 1 - ent_nd
    w_nd = (d_nd / d_nd.sum()).values
    # 加权标准化矩阵
    Vd = nd_norm.mul(w_nd, axis=1)
    # 正/负理想解
    pos_sol = {}
    neg_sol = {}
    for i, col in enumerate(dim_inds):
        if direction[col] == '正向':
            pos_sol[i] = Vd.iloc[:, i].max()
            neg_sol[i] = Vd.iloc[:, i].min()
        else:
            pos_sol[i] = Vd.iloc[:, i].min()
            neg_sol[i] = Vd.iloc[:, i].max()
    # 距离
    dP = np.sqrt(sum((Vd.iloc[:, i] - pos_sol[i])**2 for i in range(len(dim_inds))))
    dN = np.sqrt(sum((Vd.iloc[:, i] - neg_sol[i])**2 for i in range(len(dim_inds))))
    c = dN / (dP + dN)
    col_name = f'{dim_name}得分'
    panel[col_name] = c.values
    dim_score_cols.append(col_name)
    print(f"  {dim_name}：得分范围 {c.min():.4f}—{c.max():.4f}，均值 {c.mean():.4f}")
print()

# ============================================================
# 6. 整理各省各年综合得分表（长→宽）
# ============================================================
score_wide = panel.pivot(index='省份', columns='年份', values='TOPSIS综合得分')
# 按年份顺序排列列
score_wide = score_wide[sorted(score_wide.columns)]

print("=== 各省综合得分（宽格式预览，前5省×所有年份）===")
print(score_wide.round(6).head().to_string())
print()

# 按年计算全国均值
year_mean = score_wide.mean(axis=0).round(6)
print("=== 全国年均综合得分（2015—2022）===")
for y, v in year_mean.items():
    print(f"  {y}年：{v:.6f}")
print()

# ============================================================
# 7. 综合结果写入Excel
# ============================================================
output = 'Step1_熵权TOPSIS结果.xlsx'
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    # Sheet1：指标权重表
    weight_df[['指标代码','指标方向','信息熵','差异系数','熵权权重','权重(%)']].to_excel(
        writer, sheet_name='指标权重', index=False)
    # Sheet2：各省各年综合得分（宽格式）
    score_wide.round(6).to_excel(writer, sheet_name='综合得分_宽格式')
    # Sheet3：各省各年综合得分（长格式）
    out_long = panel[['省份','年份','TOPSIS综合得分'] + dim_score_cols].sort_values(['年份','省份'])
    out_long.to_excel(writer, sheet_name='综合得分_长格式', index=False)
    # Sheet4：各年得分统计描述
    desc = pd.DataFrame({
        '年份': list(year_mean.index),
        '全国均值': list(year_mean.values),
    })
    for y in years:
        col_data = score_wide[y]
        desc.loc[desc['年份']==y, '最大值'] = col_data.max()
        desc.loc[desc['年份']==y, '最小值'] = col_data.min()
        desc.loc[desc['年份']==y, '标准差'] = col_data.std()
    desc.to_excel(writer, sheet_name='时序统计描述', index=False)

print(f"=== 结果已保存：{output} ===")
print(f"包含Sheet：指标权重 / 综合得分_宽格式 / 综合得分_长格式 / 时序统计描述")
print("\n第一步（熵权TOPSIS）完成！")
print("下一步：描述性统计 + 时空特征分析（第四步任务）")
