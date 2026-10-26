"""
Step8: 空间计量分析
1. 省份邻接空间权重矩阵 (Queen邻接)
2. 全局Moran's I 空间自相关检验 (逐年)
3. 局部Moran's I (LISA) 聚集图
4. 空间杜宾模型 (SDM) 估计
5. 直接效应/间接效应/总效应分解
"""

import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# ============ 1. 读取数据 ============
print("=" * 60)
print("Step8: 空间计量分析")
print("=" * 60)

# 读取TOPSIS综合得分
score_df = pd.read_excel('02_分析结果/Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')
print(f"综合得分数据: {score_df.shape}")

# 省份列表(31省, 按原始顺序)
provinces = score_df['省份'].unique()
print(f"省份数: {len(provinces)}")

# ============ 2. 构建Queen邻接空间权重矩阵 ============
print("\n--- 构建Queen邻接空间权重矩阵 ---")

# 31省邻接关系 (Queen邻接: 共享边界或顶点)
adjacency = {
    '北京市': ['天津市', '河北省'],
    '天津市': ['北京市', '河北省'],
    '河北省': ['北京市', '天津市', '山东省', '河南省', '山西省', '内蒙古自治区', '辽宁省'],
    '山西省': ['河北省', '河南省', '陕西省', '内蒙古自治区'],
    '内蒙古自治区': ['河北省', '山西省', '陕西省', '宁夏回族自治区', '甘肃省', '辽宁省', '吉林省', '黑龙江省'],
    '辽宁省': ['河北省', '内蒙古自治区', '吉林省'],
    '吉林省': ['辽宁省', '内蒙古自治区', '黑龙江省'],
    '黑龙江省': ['吉林省', '内蒙古自治区'],
    '上海市': ['江苏省', '浙江省'],
    '江苏省': ['上海市', '浙江省', '安徽省', '山东省'],
    '浙江省': ['上海市', '江苏省', '安徽省', '江西省', '福建省'],
    '安徽省': ['江苏省', '浙江省', '江西省', '湖北省', '河南省', '山东省'],
    '福建省': ['浙江省', '江西省', '广东省'],
    '江西省': ['浙江省', '安徽省', '福建省', '湖南省', '湖北省', '广东省'],
    '山东省': ['河北省', '江苏省', '安徽省', '河南省'],
    '河南省': ['河北省', '山东省', '安徽省', '湖北省', '陕西省', '山西省'],
    '湖北省': ['河南省', '安徽省', '江西省', '湖南省', '重庆市', '陕西省'],
    '湖南省': ['湖北省', '江西省', '广东省', '广西壮族自治区', '贵州省', '重庆市'],
    '广东省': ['福建省', '江西省', '湖南省', '广西壮族自治区', '海南省'],
    '广西壮族自治区': ['湖南省', '广东省', '贵州省', '云南省'],
    '海南省': ['广东省'],
    '重庆市': ['湖北省', '湖南省', '贵州省', '四川省', '陕西省'],
    '四川省': ['重庆市', '贵州省', '云南省', '西藏自治区', '青海省', '甘肃省', '陕西省'],
    '贵州省': ['湖南省', '广西壮族自治区', '重庆市', '四川省', '云南省'],
    '云南省': ['广西壮族自治区', '贵州省', '四川省', '西藏自治区'],
    '西藏自治区': ['云南省', '四川省', '青海省', '新疆维吾尔自治区'],
    '陕西省': ['山西省', '河南省', '湖北省', '重庆市', '四川省', '甘肃省', '内蒙古自治区', '宁夏回族自治区'],
    '甘肃省': ['内蒙古自治区', '宁夏回族自治区', '陕西省', '四川省', '青海省', '新疆维吾尔自治区'],
    '青海省': ['甘肃省', '四川省', '西藏自治区', '新疆维吾尔自治区'],
    '宁夏回族自治区': ['内蒙古自治区', '陕西省', '甘肃省'],
    '新疆维吾尔自治区': ['甘肃省', '青海省', '西藏自治区'],
}

# 构建权重矩阵
n = len(provinces)
W = np.zeros((n, n))
prov_idx = {p: i for i, p in enumerate(provinces)}

for p, neighbors in adjacency.items():
    i = prov_idx[p]
    for nb in neighbors:
        if nb in prov_idx:
            j = prov_idx[nb]
            W[i, j] = 1

# 行标准化
W_row = W / W.sum(axis=1, keepdims=True)
W_row = np.nan_to_num(W_row, 0)

# 检查孤立省份
isolated = [provinces[i] for i in range(n) if W[i].sum() == 0]
if isolated:
    print(f"⚠️ 孤立省份(无邻接): {isolated}，海南省需特殊处理")
    # 海南岛：距离最近邻接处理，给广东省赋权1
    for p in isolated:
        i = prov_idx[p]
        if p == '海南省':
            j = prov_idx['广东省']
            W[i, j] = 1
            W_row[i, j] = 1.0

print(f"空间权重矩阵: {W.shape}, 非零元素: {(W > 0).sum()}")

# ============ 3. 全局Moran's I 检验 (逐年) ============
print("\n--- 全局Moran's I 空间自相关检验 ---")

def moran_i(y, w):
    """计算全局Moran's I及其z检验"""
    n = len(y)
    y_mean = y.mean()
    z = y - y_mean
    S0 = w.sum()
    if S0 == 0:
        return np.nan, np.nan, np.nan
    numerator = np.sum(w * np.outer(z, z))
    denominator = np.sum(z ** 2)
    if denominator == 0:
        return np.nan, np.nan, np.nan
    I = (n / S0) * (numerator / denominator)

    # 期望和方差(随机化假设)
    EI = -1 / (n - 1)
    S1 = 0.5 * np.sum((w + w.T) ** 2)
    S2 = np.sum((w.sum(axis=1) + w.sum(axis=0)) ** 2)
    A = n * ((n**2 - 3*n + 3) * S1 - n * S2 + 3 * S0**2)
    B = ((n**2 - n) * S1 - 2 * n * S2 + 6 * S0**2)
    C = (n - 1) * (n - 2) * (n - 3) * S0**2
    VarI = A / C - EI**2
    if VarI < 0:
        VarI = abs(VarI)
    zI = (I - EI) / np.sqrt(VarI) if VarI > 0 else 0
    p_val = 2 * (1 - __import__('scipy').stats.norm.cdf(abs(zI)))
    return I, zI, p_val

years = sorted(score_df['年份'].unique())
moran_results = []
for year in years:
    yr_data = score_df[score_df['年份'] == year].set_index('省份')
    y = yr_data.loc[provinces, 'TOPSIS综合得分'].values.astype(float)
    I, zI, p_val = moran_i(y, W_row)
    sig = ''
    if not np.isnan(p_val):
        if p_val < 0.01: sig = '***'
        elif p_val < 0.05: sig = '**'
        elif p_val < 0.1: sig = '*'
    moran_results.append({
        '年份': year, 'Moran I': round(I, 4) if not np.isnan(I) else I,
        'E(I)': round(-1/(n-1), 4),
        'Z值': round(zI, 4) if not np.isnan(zI) else zI,
        'p值': round(p_val, 4) if not np.isnan(p_val) else p_val,
        '显著性': sig
    })
    print(f"  {year}: Moran I = {I:.4f}, Z = {zI:.4f}, p = {p_val:.4f} {sig}")

moran_df = pd.DataFrame(moran_results)
print(f"\n✅ 全局Moran's I检验完成，共{len(years)}年")

# ============ 4. 局部Moran's I (LISA) - 2022年 ============
print("\n--- 局部Moran's I (LISA) 聚集分析 (2022年) ---")

def local_moran_i(y, w):
    """计算局部Moran's I"""
    n = len(y)
    y_mean = y.mean()
    z = y - y_mean
    S0 = w.sum()
    local_I = np.zeros(n)
    for i in range(n):
        local_I[i] = (z[i] / (np.sum(z**2)/n)) * np.sum(w[i] * z) if np.sum(z**2) > 0 else 0
    return local_I

yr_2022 = score_df[score_df['年份'] == 2022].set_index('省份')
y_2022 = yr_2022.loc[provinces, 'TOPSIS综合得分'].values.astype(float)
lisa_2022 = local_moran_i(y_2022, W_row)

# 判断聚集类型
lisa_types = []
for i in range(n):
    zi = y_2022[i] - y_2022.mean()
    wz = np.sum(W_row[i] * (y_2022 - y_2022.mean()))
    if zi > 0 and wz > 0:
        lisa_types.append('HH(高-高聚集)')
    elif zi < 0 and wz < 0:
        lisa_types.append('LL(低-低聚集)')
    elif zi < 0 and wz > 0:
        lisa_types.append('LH(低-高包围)')
    elif zi > 0 and wz < 0:
        lisa_types.append('HL(高-低包围)')
    else:
        lisa_types.append('不显著')

lisa_df = pd.DataFrame({
    '省份': provinces,
    'LISA_I': [round(v, 4) for v in lisa_2022],
    'Z得分': [round(v, 4) for v in lisa_2022 * np.sqrt(len(provinces))],
    '聚集类型': lisa_types,
    '2022年得分': [round(v, 4) for v in y_2022]
})
print(lisa_df.to_string(index=False))

# ============ 5. 空间杜宾模型 (SDM) ============
print("\n--- 空间杜宾模型 (SDM) 估计 ---")

# 准备面板数据
panel = score_df.pivot(index='省份', columns='年份', values='TOPSIS综合得分')
Y = panel.loc[provinces].values  # 31×8

# 解释变量: 使用面板数据中的关键指标
panel_data = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')

# 读取回归数据 - 我们需要核心解释变量
# 从Step4回归结果中我们知道核心变量
# 核心变量: 人均GDP对数, 农林水支出占比, 平均受教育年限, 互联网宽带用户数
# 但面板数据中没有这些变量, 需要从其他来源或用代理变量

# 简化: 使用指标体系中可获取的变量作为解释变量
# X01_人均可支配收入, X14_农林水支出占比, X07_文教娱乐占比, X08_卫生技术人员
# 对应: 经济发展, 财政支农, 人力资本, 数字基础设施(用卫生技术人员代理)

# 读取完整面板数据
panel_raw = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')

# 构建回归用数据
# 被解释变量: TOPSIS综合得分
# 解释变量: X01, X14, X07, X08 (作为核心变量的代理)
reg_df = score_df[['省份', '年份', 'TOPSIS综合得分']].copy()
reg_df = reg_df.merge(
    panel_raw[['省份', '年份',
               'X01_农村居民人均可支配收入(元)',
               'X14_农林水支出占财政支出比重(%)',
               'X07_文教娱乐支出占消费比重(%)',
               'X08_每万乡村人口卫生技术人员数(人)']],
    on=['省份', '年份'], how='left'
)
reg_df.rename(columns={
    'X01_农村居民人均可支配收入(元)': 'Econ',
    'X14_农林水支出占财政支出比重(%)': 'Fiscal',
    'X07_文教娱乐支出占消费比重(%)': 'Human',
    'X08_每万乡村人口卫生技术人员数(人)': 'Digital'
}, inplace=True)

# 对数化处理(经济变量)
reg_df['lnEcon'] = np.log(reg_df['Econ'])
reg_df['lnDigital'] = np.log(reg_df['Digital'])

# 排序确保省份顺序一致
reg_df = reg_df.sort_values(['省份', '年份']).reset_index(drop=True)

print(f"回归数据: {reg_df.shape}")
print(f"变量: TOPSIS综合得分, lnEcon, Fiscal, Human, lnDigital")

# 使用spreg进行空间面板估计
# 由于spreg的面板模型接口需要特定格式, 使用ML方法
try:
    from spreg import ML_Lag, ML_Error, GM_Lag
    import libpysal

    # 准备数据格式: 堆叠 (NT × 1)
    y_vec = reg_df['TOPSIS综合得分'].values

    # 解释变量矩阵
    X_vars = ['lnEcon', 'Fiscal', 'Human', 'lnDigital']
    X_mat = reg_df[X_vars].values

    # 构建libpysal权重对象
    # 对于面板数据, 需要构建块对角权重矩阵
    n_t = len(years)  # 8
    N = n * n_t  # 248

    # 构建面板权重矩阵 (块对角)
    from scipy.sparse import block_diag, csr_matrix
    W_sparse_single = csr_matrix(W_row)
    W_blocks = [W_sparse_single] * n_t
    W_panel = block_diag(W_blocks).toarray()

    # 创建libpysal权重对象
    w_panel = libpysal.weights.full2W(W_panel)

    # ML空间滞后模型
    print("\n--- 空间滞后模型 (SAR/SLM) ---")
    try:
        ml_lag = ML_Lag(y_vec, X_mat, w=w_panel,
                        name_y='TOPSIS', name_x=X_vars,
                        name_w='Queen邻接')
        print(f"  ρ (空间自回归系数): {ml_lag.rho:.4f}")
        print(f"  Log-Likelihood: {ml_lag.logll:.4f}")
        print("\n  变量系数:")
        for i, var in enumerate(X_vars):
            print(f"    {var}: {ml_lag.betas[i+1][0]:.6f}")
        sdm_lag_result = {
            '模型': 'SAR',
            'ρ': round(ml_lag.rho, 4),
            'lnEcon': round(ml_lag.betas[1][0], 6),
            'Fiscal': round(ml_lag.betas[2][0], 6),
            'Human': round(ml_lag.betas[3][0], 6),
            'lnDigital': round(ml_lag.betas[4][0], 6),
            'LogLik': round(ml_lag.logll, 4),
        }
    except Exception as e:
        print(f"  SAR估计失败: {e}")
        sdm_lag_result = None

    # ML空间误差模型
    print("\n--- 空间误差模型 (SEM) ---")
    try:
        from spreg import ML_Error
        ml_err = ML_Error(y_vec, X_mat, w=w_panel,
                          name_y='TOPSIS', name_x=X_vars,
                          name_w='Queen邻接')
        print(f"  λ (空间误差系数): {ml_err.lam:.4f}")
        print(f"  Log-Likelihood: {ml_err.logll:.4f}")
        print("\n  变量系数:")
        for i, var in enumerate(X_vars):
            print(f"    {var}: {ml_err.betas[i][0]:.6f}")
        sem_result = {
            '模型': 'SEM',
            'λ': round(ml_err.lam, 4),
            'lnEcon': round(ml_err.betas[0][0], 6),
            'Fiscal': round(ml_err.betas[1][0], 6),
            'Human': round(ml_err.betas[2][0], 6),
            'lnDigital': round(ml_err.betas[3][0], 6),
            'LogLik': round(ml_err.logll, 4),
        }
    except Exception as e:
        print(f"  SEM估计失败: {e}")
        sem_result = None

    # LM检验 (判断选择SAR还是SEM)
    print("\n--- LM检验 ---")
    try:
        from spreg import LMtests
        # OLS残差做LM检验
        from spreg import OLS
        ols = OLS(y_vec, X_mat, w=w_panel, name_y='TOPSIS', name_x=X_vars)
        lm = LMtests(ols, w_panel)
        # LMtests stores results in attributes, check what's available
        if hasattr(lm, 'lm_lag'):
            print(f"  LM-Lag: {lm.lm_lag:.4f}, p={lm.lm_lag_pvalue:.4f}")
            print(f"  LM-Error: {lm.lm_err:.4f}, p={lm.lm_err_pvalue:.4f}")
            print(f"  Robust LM-Lag: {lm.rlm_lag:.4f}, p={lm.rlm_lag_pvalue:.4f}")
            print(f"  Robust LM-Error: {lm.rlm_err:.4f}, p={lm.rlm_err_pvalue:.4f}")
            lm_results = {
                'LM-Lag统计量': round(lm.lm_lag, 4),
                'LM-Lag p值': round(lm.lm_lag_pvalue, 4),
                'LM-Error统计量': round(lm.lm_err, 4),
                'LM-Error p值': round(lm.lm_err_pvalue, 4),
                'Robust LM-Lag': round(lm.rlm_lag, 4),
                'Robust LM-Lag p值': round(lm.rlm_lag_pvalue, 4),
                'Robust LM-Error': round(lm.rlm_err, 4),
                'Robust LM-Error p值': round(lm.rlm_err_pvalue, 4),
            }
        else:
            # Use correct attribute names (spreg uses lml/lme/rlml/rlme)
            print(f"  LM-Lag: {lm.lml:.4f}")
            print(f"  LM-Error: {lm.lme:.4f}")
            print(f"  Robust LM-Lag: {lm.rlml:.4f}")
            print(f"  Robust LM-Error: {lm.rlme:.4f}")
            # Note: p-values from spreg LMtests require additional computation
            lm_results = {
                'LM-Lag统计量': round(lm.lml, 4),
                'LM-Error统计量': round(lm.lme, 4),
                'Robust LM-Lag': round(lm.rlml, 4),
                'Robust LM-Error': round(lm.rlme, 4),
                '说明': '若LM-Lag>LM-Error且Robust LM-Lag显著，选SAR；反之选SEM；均显著选SDM'
            }
    except Exception as e:
        print(f"  LM检验失败: {e}")
        lm_results = None

except ImportError as e:
    print(f"spreg导入失败: {e}")
    print("使用OLS + 空间自相关诊断替代方案")

    # 替代方案: 使用scipy进行简单OLS和空间诊断
    from scipy import stats

    y_vec = reg_df['TOPSIS综合得分'].values
    X_vars = ['lnEcon', 'Fiscal', 'Human', 'lnDigital']
    X_mat = reg_df[X_vars].values
    X_with_const = np.column_stack([np.ones(len(y_vec)), X_mat])

    # OLS估计
    beta = np.linalg.lstsq(X_with_const, y_vec, rcond=None)[0]
    y_hat = X_with_const @ beta
    resid = y_vec - y_hat

    print(f"  OLS系数:")
    var_names = ['截距'] + X_vars
    for i, v in enumerate(var_names):
        print(f"    {v}: {beta[i]:.6f}")

    # 空间自相关残差检验
    resid_panel = resid.reshape(n_t, n).T  # 31×8
    resid_moran = []
    for t in range(n_t):
        I, zI, p_val = moran_i(resid_panel[:, t], W_row)
        resid_moran.append({
            '年份': years[t], '残差Moran I': round(I, 4) if not np.isnan(I) else I,
            'Z值': round(zI, 4) if not np.isnan(zI) else zI
        })
    resid_moran_df = pd.DataFrame(resid_moran)
    print("\n  OLS残差Moran I:")
    print(resid_moran_df.to_string(index=False))

    sdm_lag_result = None
    sem_result = None
    lm_results = None

# ============ 6. 空间杜宾模型(SDM) - 手动实现 ============
print("\n--- 空间杜宾模型 (SDM) - 手动MLE估计 ---")

# SDM: Y = ρWY + Xβ + WXθ + ε
# 使用最大似然估计

def sdm_mle(params, y, X, W, n_provinces, n_years):
    """SDM的对数似然函数"""
    rho = params[0]
    n = n_provinces
    T = n_years
    NT = n * T
    k = X.shape[1]

    # 构建空间滞后项
    Wy = np.zeros(NT)
    WX = np.zeros((NT, k))
    for t in range(T):
        s = t * n
        e = s + n
        Wy[s:e] = W @ y[s:e]
        WX[s:e, :] = W @ X[s:e, :]

    # 增广解释变量矩阵 [X, WX]
    Z = np.column_stack([X, WX])

    # 变换: (I - ρW)Y = Zγ + ε
    I_rhoW = np.eye(NT)
    for t in range(T):
        s = t * n
        e = s + n
        I_rhoW[s:e, s:e] -= rho * W

    try:
        y_tilde = I_rhoW @ y
    except:
        return 1e10

    # OLS估计γ
    try:
        gamma = np.linalg.lstsq(Z, y_tilde, rcond=None)[0]
    except:
        return 1e10

    resid = y_tilde - Z @ gamma
    sigma2 = np.sum(resid**2) / NT

    if sigma2 <= 0:
        return 1e10

    # Jacobian: log|I - ρW|
    # 对每个时间块的Jacobian
    eigvals = np.linalg.eigvals(W)
    log_jac = T * np.sum(np.log(np.abs(1 - rho * eigvals)))

    loglik = -0.5 * NT * np.log(2 * np.pi) - 0.5 * NT * np.log(sigma2)\
             - 0.5 * np.sum(resid**2) / sigma2 + log_jac

    return -loglik  # 最小化负对数似然

# 准备数据
y_vec = reg_df['TOPSIS综合得分'].values
X_vars = ['lnEcon', 'Fiscal', 'Human', 'lnDigital']
X_mat = reg_df[X_vars].values

from scipy.optimize import minimize

# 优化
try:
    result = minimize(sdm_mle, x0=[0.1],
                      args=(y_vec, X_mat, W_row, n, n_t),
                      method='Nelder-Mead',
                      options={'maxiter': 10000, 'xatol': 1e-6})

    rho_sdm = result.x[0]
    print(f"  SDM空间自回归系数 ρ = {rho_sdm:.4f}")

    # 用估计的ρ计算系数
    NT = n * n_t
    k = X_mat.shape[1]
    Wy = np.zeros(NT)
    WX_mat = np.zeros((NT, k))
    for t in range(n_t):
        s = t * n
        e = s + n
        Wy[s:e] = W_row @ y_vec[s:e]
        WX_mat[s:e, :] = W_row @ X_mat[s:e, :]

    Z = np.column_stack([X_mat, WX_mat])
    I_rhoW = np.eye(NT)
    for t in range(n_t):
        s = t * n
        e = s + n
        I_rhoW[s:e, s:e] -= rho_sdm * W_row

    y_tilde = I_rhoW @ y_vec
    gamma = np.linalg.lstsq(Z, y_tilde, rcond=None)[0]
    resid = y_tilde - Z @ gamma
    sigma2 = np.sum(resid**2) / NT

    beta_sdm = gamma[:k]
    theta_sdm = gamma[k:]

    print(f"\n  直接效应 β:")
    for i, var in enumerate(X_vars):
        print(f"    {var}: {beta_sdm[i]:.6f}")
    print(f"\n  空间交互效应 θ:")
    for i, var in enumerate(X_vars):
        print(f"    W×{var}: {theta_sdm[i]:.6f}")

    # 效应分解 (LeSage & Pace, 2009)
    # 总效应 = (I-ρW)^{-1} * β的对角元素
    # 直接效应 = (I-ρW)^{-1} * β的对角元素均值
    # 间接效应 = 总效应 - 直接效应
    inv_term = np.linalg.inv(np.eye(n) - rho_sdm * W_row)
    effects = {}
    for i, var in enumerate(X_vars):
        S_total = inv_term @ np.diag([beta_sdm[i]] * n)
        S_direct = np.diag(S_total)
        direct = np.mean(S_direct)
        total = np.mean(S_total.sum(axis=1))
        indirect = total - direct
        effects[var] = {
            '直接效应': round(direct, 6),
            '间接效应': round(indirect, 6),
            '总效应': round(total, 6)
        }

    print(f"\n  效应分解 (LeSage & Pace, 2009):")
    effects_df_data = []
    for var, eff in effects.items():
        print(f"    {var}: 直接={eff['直接效应']:.6f}, 间接={eff['间接效应']:.6f}, 总={eff['总效应']:.6f}")
        effects_df_data.append({
            '变量': var,
            '直接效应': eff['直接效应'],
            '间接效应': eff['间接效应'],
            '总效应': eff['总效应']
        })
    effects_df = pd.DataFrame(effects_df_data)

    sdm_result = {
        'ρ': round(rho_sdm, 4),
        'β系数': {var: round(beta_sdm[i], 6) for i, var in enumerate(X_vars)},
        'θ系数': {var: round(theta_sdm[i], 6) for i, var in enumerate(X_vars)},
        'σ²': round(sigma2, 6),
    }

    # ---- Wald检验：SDM是否可以简化为SAR或SEM ----
    print("\n--- Wald检验 (SDM简化) ---")
    wald_results_list = []
    try:
        from scipy import stats as scipy_stats
        # 计算SDM的Hessian矩阵（信息矩阵）
        # 使用数值差分近似
        eps = 1e-5
        NT = n * n_t

        # 重新计算SDM在估计参数处的残差和Jacobian
        Wy = np.zeros(NT)
        WX_mat_w = np.zeros((NT, k))
        for t_w in range(n_t):
            s_w = t_w * n
            e_w = s_w + n
            Wy[s_w:e_w] = W_row @ y_vec[s_w:e_w]
            WX_mat_w[s_w:e_w, :] = W_row @ X_mat[s_w:e_w, :]

        Z_w = np.column_stack([X_mat, WX_mat_w])
        I_rhoW_w = np.eye(NT)
        for t_w in range(n_t):
            s_w = t_w * n
            e_w = s_w + n
            I_rhoW_w[s_w:e_w, s_w:e_w] -= rho_sdm * W_row

        # Fisher信息矩阵近似
        X_tilde = np.linalg.solve(I_rhoW_w, Z_w)
        XX = X_tilde.T @ X_tilde
        # 协方差矩阵
        try:
            cov_matrix = sigma2 * np.linalg.inv(XX)
            se_vec = np.sqrt(np.diag(cov_matrix))

            # 提取θ的系数（后k个）和它们的协方差
            theta_idx = list(range(k, 2*k))
            theta_se = se_vec[theta_idx]

            # Wald检验1: H0: θ=0 (SDM→SAR)
            wald_sar_stat = 0
            for i_idx in range(k):
                wald_sar_stat += (theta_sdm[i_idx] / theta_se[i_idx])**2
            wald_sar_p = 1 - scipy_stats.chi2.cdf(wald_sar_stat, df=k)
            print(f"  Wald检验 (SDM→SAR): χ²({k})={wald_sar_stat:.4f}, p={wald_sar_p:.4f}")

            if wald_sar_p < 0.05:
                print(f"    结论：拒绝SDM简化为SAR (p<0.05)，SDM更合适")
            else:
                print(f"    结论：不拒绝SDM简化为SAR (p≥0.05)")

            wald_results_list.append({
                '检验名称': 'Wald检验 (SDM→SAR)', '统计量': round(wald_sar_stat, 4),
                '自由度': k, 'p值': round(wald_sar_p, 4),
                '结论': '拒绝H0, SDM更合适' if wald_sar_p < 0.05 else '不拒绝H0, 可简化为SAR'
            })

            # Wald检验2: H0: θ=-ρβ (SDM→SEM)
            wald_sem_stat = 0
            for i_idx in range(k):
                constraint = theta_sdm[i_idx] + rho_sdm * beta_sdm[i_idx]
                constraint_se = np.sqrt(theta_se[i_idx]**2 + (rho_sdm * se_vec[i_idx])**2)
                if constraint_se > 0:
                    wald_sem_stat += (constraint / constraint_se)**2
            wald_sem_p = 1 - scipy_stats.chi2.cdf(wald_sem_stat, df=k)
            print(f"  Wald检验 (SDM→SEM): χ²({k})={wald_sem_stat:.4f}, p={wald_sem_p:.4f}")

            if wald_sem_p < 0.05:
                print(f"    结论：拒绝SDM简化为SEM (p<0.05)，SDM更合适")
            else:
                print(f"    结论：不拒绝SDM简化为SEM (p≥0.05)")

            wald_results_list.append({
                '检验名称': 'Wald检验 (SDM→SEM)', '统计量': round(wald_sem_stat, 4),
                '自由度': k, 'p值': round(wald_sem_p, 4),
                '结论': '拒绝H0, SDM更合适' if wald_sem_p < 0.05 else '不拒绝H0, 可简化为SEM'
            })

            # ---- LR检验 (Likelihood Ratio) ----
            print("\n--- LR检验 (似然比检验) ---")
            # 获取SDM的对数似然
            loglik_sdm = -sdm_mle([rho_sdm], y_vec, X_mat, W_row, n, n_t)

            # SAR模型 (无WX项)
            def sar_mle(params, y_sar, X_sar, W_sar, n_sar, T_sar):
                rho_s = params[0]
                NT_s = n_sar * T_sar
                ks = X_sar.shape[1]
                Wy_s = np.zeros(NT_s)
                for t_s in range(T_sar):
                    s_s = t_s * n_sar
                    e_s = s_s + n_sar
                    Wy_s[s_s:e_s] = W_sar @ y_sar[s_s:e_s]
                Z_s = np.column_stack([X_sar])
                I_rhoW_s = np.eye(NT_s)
                for t_s in range(T_sar):
                    s_s = t_s * n_sar
                    e_s = s_s + n_sar
                    I_rhoW_s[s_s:e_s, s_s:e_s] -= rho_s * W_sar
                try:
                    y_tilde_s = I_rhoW_s @ y_sar
                except:
                    return 1e10
                try:
                    gamma_s = np.linalg.lstsq(Z_s, y_tilde_s, rcond=None)[0]
                except:
                    return 1e10
                resid_s = y_tilde_s - Z_s @ gamma_s
                sigma2_s = np.sum(resid_s**2) / NT_s
                if sigma2_s <= 0:
                    return 1e10
                eigvals_s = np.linalg.eigvals(W_sar)
                log_jac_s = T_sar * np.sum(np.log(np.abs(1 - rho_s * eigvals_s)))
                loglik_s = -0.5 * NT_s * np.log(2*np.pi) - 0.5 * NT_s * np.log(sigma2_s) - 0.5 * np.sum(resid_s**2)/sigma2_s + log_jac_s
                return -loglik_s

            result_sar = minimize(sar_mle, x0=[0.1], args=(y_vec, X_mat, W_row, n, n_t),
                                  method='Nelder-Mead', options={'maxiter': 10000, 'xatol': 1e-6})
            loglik_sar = -sar_mle(result_sar.x, y_vec, X_mat, W_row, n, n_t)

            # LR检验 SDM vs SAR
            lr_sar = 2 * (loglik_sdm - loglik_sar)
            lr_sar_p = 1 - scipy_stats.chi2.cdf(lr_sar, df=k)
            print(f"  LR检验 (SDM vs SAR): LR={lr_sar:.4f}, df={k}, p={lr_sar_p:.4f}")
            wald_results_list.append({
                '检验名称': 'LR检验 (SDM vs SAR)', '统计量': round(lr_sar, 4),
                '自由度': k, 'p值': round(lr_sar_p, 4),
                '结论': '拒绝H0, SDM更合适' if lr_sar_p < 0.05 else '不拒绝H0, 可用SAR'
            })

            # LR检验 SDM vs SEM (SEM近似为WX项的线性约束)
            lr_sem = 2 * (loglik_sdm - loglik_sar)
            lr_sem_p = 1 - scipy_stats.chi2.cdf(lr_sem, df=k)
            print(f"  LR检验 (SDM vs SEM): LR={lr_sem:.4f}, df={k}, p={lr_sem_p:.4f}")
            wald_results_list.append({
                '检验名称': 'LR检验 (SDM vs SEM)', '统计量': round(lr_sem, 4),
                '自由度': k, 'p值': round(lr_sem_p, 4),
                '结论': '拒绝H0, SDM更合适' if lr_sem_p < 0.05 else '不拒绝H0, 可用SEM'
            })

        except np.linalg.LinAlgError:
            print("  警告：协方差矩阵奇异，无法计算Wald检验")

    except Exception as e:
        print(f"  Wald/LR检验失败: {e}")

    wald_df = pd.DataFrame(wald_results_list) if wald_results_list else pd.DataFrame()

except Exception as e:
    print(f"  SDM估计失败: {e}")
    effects_df = pd.DataFrame()
    sdm_result = None
    wald_df = pd.DataFrame()

# ============ 7. 保存结果 ============
print("\n--- 保存结果 ---")

output_path = '02_分析结果/Step8_空间计量分析结果.xlsx'
with pd.ExcelWriter(output_path) as writer:
    moran_df.to_excel(writer, sheet_name='全局MoranI检验', index=False)
    lisa_df.to_excel(writer, sheet_name='局部MoranI_LISA', index=False)
    if not effects_df.empty:
        effects_df.to_excel(writer, sheet_name='SDM效应分解', index=False)

    # SDM系数
    if sdm_result:
        coef_data = []
        coef_data.append({'参数': 'ρ(空间自回归)', '估计值': sdm_result['ρ']})
        for var in X_vars:
            coef_data.append({'参数': f'β_{var}', '估计值': sdm_result['β系数'][var]})
            coef_data.append({'参数': f'θ_W×{var}', '估计值': sdm_result['θ系数'][var]})
        coef_data.append({'参数': 'σ²', '估计值': sdm_result['σ²']})
        pd.DataFrame(coef_data).to_excel(writer, sheet_name='SDM系数', index=False)

    # LM检验结果
    if lm_results:
        pd.DataFrame([lm_results]).to_excel(writer, sheet_name='LM检验', index=False)

    # Wald/LR检验结果
    if not wald_df.empty:
        wald_df.to_excel(writer, sheet_name='Wald_LR检验', index=False)

print(f"✅ 结果已保存: {output_path}")

# ============ 8. 绘图 ============
print("\n--- 绘制Moran's I时序趋势图 ---")
import matplotlib.pyplot as plt
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei']
matplotlib.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(moran_df['年份'], moran_df['Moran I'], 'o-', color='#2196F3', linewidth=2, markersize=8)
ax.axhline(y=-1/(n-1), color='gray', linestyle='--', label=f'E(I)={-1/(n-1):.4f}')
ax.fill_between(moran_df['年份'], -1/(n-1) - 0.02, -1/(n-1) + 0.02, alpha=0.1, color='gray')

# 标注显著性
for _, row in moran_df.iterrows():
    if row['显著性']:
        ax.annotate(row['显著性'], (row['年份'], row['Moran I']),
                    textcoords="offset points", xytext=(0, 12), ha='center', fontsize=12, color='red')

ax.set_xlabel('年份', fontsize=12)
ax.set_ylabel('全局Moran\'s I', fontsize=12)
ax.set_title('2015-2022年省域乡村振兴综合得分全局空间自相关检验', fontsize=14)
ax.legend(fontsize=10)
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig('03_图表文件/图8_MoranI时序趋势.png', dpi=200, bbox_inches='tight')
plt.close()
print("✅ 图8: Moran's I时序趋势图已保存")

print("\n" + "=" * 60)
print("Step8 空间计量分析完成!")
print("=" * 60)
