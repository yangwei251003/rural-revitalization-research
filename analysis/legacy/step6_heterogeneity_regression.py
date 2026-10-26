"""
Step6: 区域异质性分组回归分析
分四大区域（东部/中部/西部/东北）分别进行面板固定效应回归
"""
import pandas as pd
import numpy as np
from linearmodels.panel import PanelOLS
import warnings
warnings.filterwarnings('ignore')

# 区域划分（使用全称匹配数据）
region_mapping = {
    # 东部10省
    '北京市': '东部', '天津市': '东部', '河北省': '东部', '上海市': '东部',
    '江苏省': '东部', '浙江省': '东部', '福建省': '东部', '山东省': '东部',
    '广东省': '东部', '海南省': '东部',
    # 中部6省
    '山西省': '中部', '安徽省': '中部', '江西省': '中部', '河南省': '中部',
    '湖北省': '中部', '湖南省': '中部',
    # 西部12省
    '内蒙古自治区': '西部', '广西壮族自治区': '西部', '重庆市': '西部', '四川省': '西部',
    '贵州省': '西部', '云南省': '西部', '西藏自治区': '西部', '陕西省': '西部',
    '甘肃省': '西部', '青海省': '西部', '宁夏回族自治区': '西部', '新疆维吾尔自治区': '西部',
    # 东北3省
    '辽宁省': '东北', '吉林省': '东北', '黑龙江省': '东北'
}

print("=" * 60)
print("Step6: 区域异质性分组回归分析")
print("=" * 60)

# 1. 读取面板数据
df = pd.read_excel('乡村振兴面板数据_2015-2022.xlsx', sheet_name='面板数据')
print(f"✅ 读取面板数据: {df.shape[0]}行 × {df.shape[1]}列")

# 2. 读取Step1的乡村振兴综合得分（长格式）
df_score = pd.read_excel('Step1_熵权TOPSIS结果.xlsx', sheet_name='综合得分_长格式')
# 重命名列，保持一致性
df_score = df_score.rename(columns={'TOPSIS综合得分': '乡村振兴综合得分'})
print(f"✅ 读取乡村振兴得分: {df_score.shape[0]}行")
print(f"   列名: {df_score.columns.tolist()}")

# 3. 合并得分到面板数据
df = df.merge(df_score[['省份', '年份', '乡村振兴综合得分']], on=['省份', '年份'], how='left')
print(f"✅ 合并后数据: {df.shape[0]}行")

# 4. 构造解释变量
# 4.1 经济发展水平 - 用X01对数衡量
df['经济发展水平'] = np.log(df['X01_农村居民人均可支配收入(元)'] + 1)

# 4.2 财政支农力度 - X14
df['财政支农力度'] = df['X14_农林水支出占财政支出比重(%)']

# 4.3 人力资本水平 - 用X07代理
df['人力资本水平'] = df['X07_文教娱乐支出占消费比重(%)']

# 4.4 数字基础设施 - 用X09代理
df['数字基础设施'] = np.log(df['X09_公共图书馆数量(个)'] + 1)

# 4.5 控制变量
df['产业结构'] = df['X02_农林牧渔业总产值(亿元)'] / df['X02_农林牧渔业总产值(亿元)'].max()
df['农业现代化'] = df['X03_农业机械总动力(万千瓦)'] / df['X03_农业机械总动力(万千瓦)'].max()
df['交通基础设施'] = df['X09_公共图书馆数量(个)'] / df['X09_公共图书馆数量(个)'].max()

# 5. 添加区域列
df['区域'] = df['省份'].map(region_mapping)

# 6. 移除缺失值
df_clean = df.dropna(subset=['乡村振兴综合得分', '经济发展水平', '财政支农力度',
                              '人力资本水平', '数字基础设施'])
print(f"✅ 清洗后数据: {df_clean.shape[0]}行")
print(f"   区域分布: {df_clean['区域'].value_counts().to_dict()}")

# 7. 分区域回归
results = {}
regions = ['东部', '中部', '西部', '东北']

for region in regions:
    df_region = df_clean[df_clean['区域'] == region].copy()
    if len(df_region) < 10:
        print(f"\n⚠️  {region}样本数过少({len(df_region)}行)，跳过回归")
        continue

    print(f"\n{'='*60}")
    print(f"  {region}区域回归 (样本数: {len(df_region)})")
    print(f"{'='*60}")

    # 设置面板索引
    df_region = df_region.set_index(['省份', '年份'])

    # 被解释变量
    y = df_region['乡村振兴综合得分']

    # 解释变量
    X = pd.DataFrame({
        '经济发展水平': df_region['经济发展水平'],
        '财政支农力度': df_region['财政支农力度'],
        '人力资本水平': df_region['人力资本水平'],
        '数字基础设施': df_region['数字基础设施'],
        '产业结构': df_region['产业结构'],
        '农业现代化': df_region['农业现代化'],
        '交通基础设施': df_region['交通基础设施'],
    })

    # 双向固定效应模型
    model = PanelOLS(y, X, entity_effects=True, time_effects=True, drop_absorbed=True)
    result = model.fit(cov_type='clustered', cluster_entity=True)

    results[region] = {
        'result': result,
        'nobs': result.nobs,
        'rsquared': result.rsquared,
        'params': result.params,
        'pvalues': result.pvalues,
        'std_errors': result.std_errors,
        'conf_int': result.conf_int()
    }

    print(f"  R² = {result.rsquared:.4f}")
    print(f"  观测值数: {result.nobs}")
    print(f"\n  系数估计:")
    for var in result.params.index:
        coef = result.params[var]
        se = result.std_errors[var]
        p = result.pvalues[var]
        stars = '***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else ''
        print(f"    {var}: {coef:8.4f}{stars}  (se={se:.4f}, p={p:.4f})")

# 8. 输出结果到Excel
print(f"\n{'='*60}")
print("  输出结果到Excel...")
print(f"{'='*60}")

with pd.ExcelWriter('Step6_区域异质性回归结果.xlsx', engine='openpyxl') as writer:

    # Sheet1: 各区域回归结果汇总
    summary_data = []
    for region in regions:
        if region in results:
            r = results[region]
            row = {'区域': region, '样本数': r['nobs'], 'R2': round(r['rsquared'], 4)}
            for var in r['params'].index:
                row[f'{var}_系数'] = round(r['params'][var], 4)
                row[f'{var}_标准误'] = round(r['std_errors'][var], 4)
                row[f'{var}_p值'] = round(r['pvalues'][var], 4)
                stars = '***' if r['pvalues'][var] < 0.01 else '**' if r['pvalues'][var] < 0.05 else '*' if r['pvalues'][var] < 0.1 else ''
                row[f'{var}_显著性'] = stars
            summary_data.append(row)

    df_summary = pd.DataFrame(summary_data)
    df_summary.to_excel(writer, sheet_name='各区域回归结果汇总', index=False)
    print(f"  ✅ Sheet1: 各区域回归结果汇总")

    # Sheet2-5: 各区域详细结果
    for region in regions:
        if region in results:
            r = results[region]
            detail = pd.DataFrame({
                '变量': r['params'].index,
                '系数': r['params'].values,
                '标准误': r['std_errors'].values,
                't值': (r['params'].values / r['std_errors'].values),
                'p值': r['pvalues'].values,
                '95%CI下限': r['conf_int'].iloc[:, 0].values,
                '95%CI上限': r['conf_int'].iloc[:, 1].values,
                '显著性': ['***' if p < 0.01 else '**' if p < 0.05 else '*' if p < 0.1 else '' for p in r['pvalues'].values]
            })
            detail.to_excel(writer, sheet_name=f'{region}详细结果', index=False)
            print(f"  ✅ Sheet: {region}详细结果")

print("\n✅ Step6 区域异质性分组回归分析完成！")
print("   结果已保存到: Step6_区域异质性回归结果.xlsx")
print(f"\n{'='*60}")
print("  各区域R²对比：")
for region in regions:
    if region in results:
        print(f"  {region}: R² = {results[region]['rsquared']:.4f}")
print(f"{'='*60}")
