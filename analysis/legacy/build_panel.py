import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# 0. 基础设置
# ============================================================
provinces_31 = [
    '北京市','天津市','河北省','山西省','内蒙古自治区',
    '辽宁省','吉林省','黑龙江省','上海市','江苏省',
    '浙江省','安徽省','福建省','江西省','山东省',
    '河南省','湖北省','湖南省','广东省','广西壮族自治区',
    '海南省','重庆市','四川省','贵州省','云南省',
    '西藏自治区','陕西省','甘肃省','青海省','宁夏回族自治区',
    '新疆维吾尔自治区'
]
years = list(range(2015, 2023))

def parse_file(fname):
    df = pd.read_excel(fname, header=None)
    header_row = None
    for i, row in df.iterrows():
        if '地区' in str(row.values):
            header_row = i
            break
    if header_row is None:
        raise ValueError(f"找不到header行: {fname}")
    df2 = df.iloc[header_row:].reset_index(drop=True)
    df2.columns = df2.iloc[0]
    df2 = df2.iloc[1:].reset_index(drop=True)
    df2 = df2[df2['地区'].isin(provinces_31)].copy()
    return df2

def extract_matrix(df):
    result = df.set_index('地区').copy()
    rename = {}
    for c in result.columns:
        s = str(c)
        if '年' in s:
            try:
                rename[c] = int(s.replace('年',''))
            except:
                pass
    result = result.rename(columns=rename)
    avail_years = [y for y in years if y in result.columns]
    result = result[avail_years].copy()
    result = result.reindex(provinces_31)
    result = result.apply(pd.to_numeric, errors='coerce')
    return result

# ============================================================
# 1. 读取各文件
# ============================================================
print("读取原始文件...")
d01  = parse_file('01_农村居民人均可支配收入.xlsx')
d02  = parse_file('02_农林牧渔业总产值.xlsx')
d03  = parse_file('03_农业机械总动力.xlsx')
d04a = parse_file('04a_有效灌溉面积.xlsx')
d05a = parse_file('05a_化肥施用折纯量.xlsx')
d05b = parse_file('05b_农作物播种面积.xlsx')
d06a = parse_file('06a_农药使用量.xlsx')
d07a = parse_file('07a_农村居民文教娱乐支出.xlsx')
d07b = parse_file('07b_农村居民人均消费支出.xlsx')
d08a = parse_file('08a_卫生技术人员数.xlsx')
d08b = parse_file('08b_乡村人口.xlsx')
d09a = parse_file('09a_公共图书馆数量.xlsx')
d11  = parse_file('11_森林覆盖率.xlsx')
d14a = parse_file('14a_农林水支出.xlsx')
d14b = parse_file('14b_一般公共预算支出.xlsx')
d15  = parse_file('15_村民委员会数量.xlsx')
d16a = parse_file('16a_养老保险参保人数.xlsx')
d17a = parse_file('17a_城乡居民基本医疗保险年末参保人数 (万人) .xlsx')
print("文件读取完成。")

# ============================================================
# 2. 提取矩阵
# ============================================================
m01  = extract_matrix(d01)
m02  = extract_matrix(d02)
m03  = extract_matrix(d03)
m04a = extract_matrix(d04a)
m05a = extract_matrix(d05a)
m05b = extract_matrix(d05b)
m06a = extract_matrix(d06a)
m07a = extract_matrix(d07a)
m07b = extract_matrix(d07b)
m08a = extract_matrix(d08a)
m08b = extract_matrix(d08b)
m09a = extract_matrix(d09a)
m11  = extract_matrix(d11)
m14a = extract_matrix(d14a)
m14b = extract_matrix(d14b)
m15  = extract_matrix(d15)
m16a = extract_matrix(d16a)
m17a = extract_matrix(d17a)
print("矩阵提取完成。")

# ============================================================
# 3. 缺失值处理
# ============================================================

# 农药使用量2022年：全部缺失，用2021年填充
m06a_filled = m06a.copy()
m06a_filled[2022] = m06a[2021].values
print("农药使用量2022年：已用2021年数据填充")

# 森林覆盖率2019-2022：用2018年填充
m11_filled = m11.copy()
for y in [2019, 2020, 2021, 2022]:
    m11_filled[y] = m11[2018].values
print("森林覆盖率2019-2022：已用2018年数据填充")

# 村民委员会2019-2022：用2018年填充
m15_filled = m15.copy()
for y in [2019, 2020, 2021, 2022]:
    m15_filled[y] = m15[2018].values
print("村民委员会2019-2022：已用2018年数据填充")

# 文教娱乐支出：用实际比重+2017年各省基准比重填补
m07_ratio = m07a / m07b * 100
ratio_2017 = m07_ratio[2017].copy()
national_ratio_2017 = ratio_2017.dropna().mean()
ratio_2017['新疆维吾尔自治区'] = national_ratio_2017

m07_ratio_filled = pd.DataFrame(index=provinces_31, columns=years, dtype=float)
for prov in provinces_31:
    for y in years:
        if not pd.isna(m07_ratio.loc[prov, y]):
            m07_ratio_filled.loc[prov, y] = m07_ratio.loc[prov, y]
        else:
            base = ratio_2017.get(prov, national_ratio_2017)
            if pd.isna(base):
                base = national_ratio_2017
            m07_ratio_filled.loc[prov, y] = base
print(f"文教比重（全国均值={national_ratio_2017:.2f}%）缺失情况：{m07_ratio_filled.isna().sum().sum()}")

# ============================================================
# 4. 计算衍生指标
# ============================================================
print("\n计算衍生指标...")
# 指标4：有效灌溉面积占比(%) = 有效灌溉面积(千公顷) / 农作物播种面积(千公顷) * 100
m04_irrig_ratio = m04a / m05b * 100
# 指标5：化肥施用强度 = 化肥施用折纯量(万吨) / 农作物播种面积(千公顷)
m05_fert_int = m05a / m05b
# 指标6：农药施用强度 = 农药使用量(万吨) / 农作物播种面积(千公顷)
m06_pest_int = m06a_filled / m05b
# 指标8：每万乡村人口卫生技术人员数
m08_health = m08a / m08b * 10000
# 指标14：农林水支出占财政支出比重
m14_ratio = m14a / m14b * 100
print("衍生指标计算完成。")

# ============================================================
# 5. 组装面板数据（长格式）
# ============================================================
print("\n组装面板数据...")

# 定义最终指标（去掉12/13，共14个指标）
indicator_dict = {
    'X01_农村居民人均可支配收入(元)':   (m01,              '产业振兴', '正向', '直接获取'),
    'X02_农林牧渔业总产值(亿元)':       (m02,              '产业振兴', '正向', '直接获取'),
    'X03_农业机械总动力(万千瓦)':       (m03,              '产业振兴', '正向', '直接获取'),
    'X04_有效灌溉面积占比(%)':           (m04_irrig_ratio,   '产业振兴', '正向', '计算：有效灌溉面积/农作物播种面积×100；替换原"第一产业劳动生产率"指标'),
    'X05_化肥施用强度(万吨/千公顷)':    (m05_fert_int,     '产业振兴', '逆向', '计算：化肥施用折纯量/农作物播种面积'),
    'X06_农药施用强度(万吨/千公顷)':    (m06_pest_int,     '产业振兴', '逆向', '计算：农药使用量/农作物播种面积；2022年用2021年数据替代'),
    'X07_文教娱乐支出占消费比重(%)':    (m07_ratio_filled, '文化繁荣', '正向', '计算：文教娱乐支出/人均消费支出；缺失年份用各省2017年比重填补'),
    'X08_每万乡村人口卫生技术人员数(人)':(m08_health,      '人才振兴', '正向', '计算：卫生技术人员数/乡村人口×10000'),
    'X09_公共图书馆数量(个)':           (m09a,             '文化繁荣', '正向', '直接获取'),
    'X11_森林覆盖率(%)':               (m11_filled,       '生态振兴', '正向', '直接获取；2019-2022年用2018年数据填充'),
    'X14_农林水支出占财政支出比重(%)':  (m14_ratio,        '组织建设', '正向', '计算：农林水支出/一般公共预算支出×100'),
    'X15_村民委员会数量(个)':           (m15_filled,       '组织建设', '正向', '直接获取；2019-2022年用2018年数据填充'),
    'X16_城乡居民养老保险参保人数(万人)':(m16a,            '组织建设', '正向', '直接获取'),
    'X17_城乡居民医疗保险参保人数(万人)':(m17a,            '组织建设', '正向', '直接获取'),
}

# 构建长格式面板数据
rows = []
for prov in provinces_31:
    for y in years:
        row = {'省份': prov, '年份': y}
        for ind_name, (mat, dim, direction, note) in indicator_dict.items():
            row[ind_name] = mat.loc[prov, y] if prov in mat.index else np.nan
        rows.append(row)

panel_df = pd.DataFrame(rows)
print(f"面板数据形状: {panel_df.shape}")
print(f"面板数据总缺失: {panel_df.isna().sum().sum()}")

# 检查各指标缺失
for col in panel_df.columns[2:]:
    missing = panel_df[col].isna().sum()
    if missing > 0:
        print(f"  {col}: 缺失 {missing} 格")

# ============================================================
# 6. 宽格式数据（每个指标独立sheet）
# ============================================================

# ============================================================
# 7. 写入Excel
# ============================================================
print("\n写入Excel文件...")

output_path = '乡村振兴面板数据_2015-2022.xlsx'

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

    # ----- Sheet1: 面板数据（长格式）-----
    panel_df.to_excel(writer, sheet_name='面板数据', index=False)

    # ----- Sheet2: 宽格式（每个指标一个子块）-----
    wide_rows = []
    header_row_wide = ['维度', '指标名称', '省份'] + [str(y) for y in years]
    wide_rows.append(header_row_wide)
    for ind_name, (mat, dim, direction, note) in indicator_dict.items():
        for prov in provinces_31:
            r = [dim, ind_name, prov]
            for y in years:
                val = mat.loc[prov, y] if prov in mat.index else np.nan
                r.append(val)
            wide_rows.append(r)
    wide_df = pd.DataFrame(wide_rows[1:], columns=wide_rows[0])
    wide_df.to_excel(writer, sheet_name='宽格式数据', index=False)

    # ----- Sheet3: 指标说明 -----
    meta_rows = []
    dim_map = {
        '产业振兴': '产业兴旺',
        '人才振兴': '人才振兴',
        '文化繁荣': '文化繁荣',
        '生态振兴': '生态宜居',
        '组织建设': '组织有效',
    }
    # 重新排序：按维度顺序
    dim_order = ['产业振兴', '人才振兴', '文化繁荣', '生态振兴', '组织建设']
    meta_rows.append(['序号', '维度', '指标名称', '指标方向', '数据说明', '来源'])
    idx = 1
    for dim in dim_order:
        for ind_name, (mat, d, direction, note) in indicator_dict.items():
            if d == dim:
                meta_rows.append([idx, dim, ind_name, direction, note, '国家统计局分省年度数据'])
                idx += 1
    meta_df = pd.DataFrame(meta_rows[1:], columns=meta_rows[0])
    meta_df.to_excel(writer, sheet_name='指标说明', index=False)

    # ----- Sheet4: 缺失值说明 -----
    missing_notes = [
        ['指标', '缺失情况', '处理方法', '说明'],
        ['X04_有效灌溉面积占比', '原X04"第一产业劳动生产率"因第一产业就业人员数据2015-2019年缺失而替换', '替换为"有效灌溉面积占比(%)"，数据完整', '有效灌溉面积占比反映农业基础设施水平，数据来源于国家统计局分省年度数据'],
        ['X06_农药施用强度', '2022年全部31省缺失', '用2021年数据替代', '国家统计局2022年起停止公布分省农药使用量'],
        ['X07_文教娱乐支出占消费比重', '大多数省份大多数年份缺失（73%）', '将指标改为"文教娱乐支出占人均消费支出比重(%)"；缺失年份用各省2017年比重填补，新疆用全国均值', '国家统计局仅部分省分年份公布该细分项'],
        ['X10_农村居民人均文教娱乐消费支出', '与X07分子重复', '指标已剔除', '原17指标中#10与#7分子相同，保留比重指标更有意义'],
        ['X11_森林覆盖率', '2019-2022年全部31省缺失', '用2018年数据填充', '森林资源清查每5年一次，第九次清查(2014-2021)结果以2018年发布数据为准'],
        ['X12_农村卫生厕所普及率', '无法获取', '指标已剔除', '属住建部专项统计，国家统计局未公布分省数据'],
        ['X13_农村生活垃圾无害化处理率', '无法获取', '指标已剔除', '属住建部专项统计，国家统计局未公布分省数据'],
        ['X15_村民委员会数量', '2019-2022年全部31省缺失', '用2018年数据填充', '国家统计年鉴2019年起停止分省发布该数据'],
        ['X16/X17', '原为"参保率"需乡村人口作分母', '改为"参保人数(万人)"', '参保率计算需乡村人口数据且口径不一，直接用绝对量指标更稳定'],
    ]
    missing_df = pd.DataFrame(missing_notes[1:], columns=missing_notes[0])
    missing_df.to_excel(writer, sheet_name='缺失说明', index=False)

print(f"Excel文件已写入: {output_path}")

# ============================================================
# 8. 美化Excel格式
# ============================================================
print("美化Excel格式...")

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = load_workbook(output_path)

# 颜色定义
COLOR_HEADER = 'FF2E75B6'  # 深蓝
COLOR_HEADER2 = 'FF4BACC6'  # 中蓝
COLOR_DIM = {
    '产业振兴': 'FFFFE599',  # 黄
    '人才振兴': 'FFFF9999',  # 粉红
    '文化繁荣': 'FF99FF99',  # 绿
    '生态振兴': 'FF99CCFF',  # 蓝
    '组织建设': 'FFFFCC99',  # 橙
}
thin = Side(style='thin')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header(ws, row=1, bg='FF2E75B6'):
    for cell in ws[row]:
        if cell.value is not None:
            cell.font = Font(bold=True, color='FFFFFFFF', size=10)
            cell.fill = PatternFill('solid', fgColor=bg)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border

def auto_width(ws, max_w=40):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                cl = len(str(cell.value)) if cell.value else 0
                max_len = max(max_len, cl)
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_len * 1.5 + 2, max_w)

# --- 面板数据 sheet ---
ws1 = wb['面板数据']
style_header(ws1, row=1)
ws1.freeze_panes = 'C2'
auto_width(ws1, max_w=30)

# 为省份列添加交替行颜色
alt_colors = ['FFFFFFFF', 'FFF2F2F2']
for i, row in enumerate(ws1.iter_rows(min_row=2, max_row=ws1.max_row), start=0):
    fill = PatternFill('solid', fgColor=alt_colors[i % 2])
    for cell in row:
        cell.fill = fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')

# --- 宽格式数据 sheet ---
ws2 = wb['宽格式数据']
style_header(ws2, row=1)
ws2.freeze_panes = 'D2'
auto_width(ws2, max_w=35)
for i, row in enumerate(ws2.iter_rows(min_row=2, max_row=ws2.max_row), start=0):
    dim_val = row[0].value
    bg = COLOR_DIM.get(dim_val, 'FFFFFFFF')
    fill = PatternFill('solid', fgColor=bg)
    for cell in row:
        cell.fill = fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center')

# --- 指标说明 sheet ---
ws3 = wb['指标说明']
style_header(ws3, row=1)
auto_width(ws3, max_w=50)
for i, row in enumerate(ws3.iter_rows(min_row=2, max_row=ws3.max_row), start=0):
    dim_val = row[1].value
    bg = COLOR_DIM.get(dim_val, 'FFFFFFFF')
    fill = PatternFill('solid', fgColor=bg)
    for cell in row:
        cell.fill = fill
        cell.border = border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

# --- 缺失说明 sheet ---
ws4 = wb['缺失说明']
style_header(ws4, row=1, bg='FFFF0000')
auto_width(ws4, max_w=60)
for cell in ws4[1]:
    if cell.value:
        cell.font = Font(bold=True, color='FFFFFFFF', size=10)
for i, row in enumerate(ws4.iter_rows(min_row=2, max_row=ws4.max_row), start=0):
    for cell in row:
        cell.border = border
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
ws4.row_dimensions[1].height = 20
for r in range(2, ws4.max_row+1):
    ws4.row_dimensions[r].height = 45

wb.save(output_path)
print(f"\n=== 完成！文件保存至: {output_path} ===")
print(f"共 {len(indicator_dict)} 个指标，{len(provinces_31)} 个省份，{len(years)} 个年份")
print(f"面板数据总行数: {len(provinces_31) * len(years)} 行")
