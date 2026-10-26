"""
在现有的论文终稿DOCX中插入所有数学公式（OMML可编辑格式）
通过Unpack→编辑document.xml→Repack实现
"""

from pathlib import Path as _WorkspacePath
import os as _workspace_os
RESEARCH_WORKSPACE = _WorkspacePath(_workspace_os.environ.get("RESEARCH_WORKSPACE", _workspace_os.getcwd())).resolve()

import zipfile, os, shutil, re, copy
from lxml import etree

BASE = str(RESEARCH_WORKSPACE)
UNPACK = f'{BASE}/unpacked_formulas'
DOCX = str(RESEARCH_WORKSPACE / '论文终稿_乡村振兴_完整排版.docx')
OUT_DOCX = str(RESEARCH_WORKSPACE / '论文终稿_乡村振兴_含公式.docx')

# Namespaces
nsmap = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

# Register namespaces for clean output
for prefix, uri in nsmap.items():
    etree.register_namespace(prefix, uri)

def Mtag(tag):
    return f'{{{M}}}{tag}'
def Wtag(tag):
    return f'{{{W}}}{tag}'

def m_text(text):
    """Create m:r with m:t"""
    mr = etree.Element(Mtag('r'))
    mt = etree.SubElement(mr, Mtag('t'))
    mt.text = text
    mt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return mr

def m_sub(base, sub):
    """Create subscript: base_sub"""
    ssub = etree.Element(Mtag('sSub'))
    ssubPr = etree.SubElement(ssub, Mtag('sSubPr'))
    e = etree.SubElement(ssub, Mtag('e'))
    if isinstance(base, str): e.append(m_text(base))
    else: e.append(base)
    sub_e = etree.SubElement(ssub, Mtag('sub'))
    if isinstance(sub, str): sub_e.append(m_text(sub))
    else: sub_e.append(sub)
    return ssub

def m_sup(base, sup):
    """Create superscript: base^sup"""
    ssup = etree.Element(Mtag('sSup'))
    ssupPr = etree.SubElement(ssup, Mtag('sSupPr'))
    e = etree.SubElement(ssup, Mtag('e'))
    if isinstance(base, str): e.append(m_text(base))
    else: e.append(base)
    sup_e = etree.SubElement(ssup, Mtag('sup'))
    if isinstance(sup, str): sup_e.append(m_text(sup))
    else: sup_e.append(sup)
    return ssup

def m_frac(num, den):
    """Create fraction: num/den"""
    f = etree.Element(Mtag('f'))
    fPr = etree.SubElement(f, Mtag('fPr'))
    type_e = etree.SubElement(fPr, Mtag('type'))
    type_e.set(Mtag('val'), 'bar')
    num_e = etree.SubElement(f, Mtag('num'))
    if isinstance(num, str): num_e.append(m_text(num))
    else: num_e.append(num)
    den_e = etree.SubElement(f, Mtag('den'))
    if isinstance(den, str): den_e.append(m_text(den))
    else: den_e.append(den)
    return f

def m_paren(content):
    """Wrap in parentheses"""
    d = etree.Element(Mtag('d'))
    dPr = etree.SubElement(d, Mtag('dPr'))
    beg = etree.SubElement(dPr, Mtag('begChr'))
    beg.set(Mtag('val'), '(')
    end = etree.SubElement(dPr, Mtag('endChr'))
    end.set(Mtag('val'), ')')
    e = etree.SubElement(d, Mtag('e'))
    if isinstance(content, str): e.append(m_text(content))
    else: e.append(content)
    return d

def m_sum(upper, lower, body):
    """Create summation ∑"""
    nary = etree.Element(Mtag('nary'))
    naryPr = etree.SubElement(nary, Mtag('naryPr'))
    chr_e = etree.SubElement(naryPr, Mtag('chr'))
    chr_e.set(Mtag('val'), '\u2211')  # ∑
    limLoc = etree.SubElement(naryPr, Mtag('limLoc'))
    limLoc.set(Mtag('val'), 'subSup')
    sub_e = etree.SubElement(nary, Mtag('sub'))
    if isinstance(lower, str): sub_e.append(m_text(lower))
    else: sub_e.append(lower)
    sup_e = etree.SubElement(nary, Mtag('sup'))
    if isinstance(upper, str): sup_e.append(m_text(upper))
    else: sup_e.append(upper)
    e = etree.SubElement(nary, Mtag('e'))
    if isinstance(body, str): e.append(m_text(body))
    else: e.append(body)
    return nary

def m_sqrt(content):
    """Create square root"""
    rad = etree.Element(Mtag('rad'))
    radPr = etree.SubElement(rad, Mtag('radPr'))
    deg = etree.SubElement(radPr, Mtag('degHide'))
    deg.set(Mtag('val'), '1')
    e = etree.SubElement(rad, Mtag('e'))
    if isinstance(content, str): e.append(m_text(content))
    else: e.append(content)
    return rad

def m_group(*items):
    """Create a group of math elements"""
    g = etree.Element(Mtag('r'))
    for item in items:
        if isinstance(item, str):
            g.append(m_text(item))
        else:
            g.append(item)
    return g

def make_formula_paragraph(math_elements, formula_number):
    """Create a paragraph containing an OMML math formula"""
    p = etree.Element(Wtag('p'))
    
    # Paragraph properties: centered
    pPr = etree.SubElement(p, Wtag('pPr'))
    jc = etree.SubElement(pPr, Wtag('jc'))
    jc.set(Wtag('val'), 'center')
    spacing = etree.SubElement(pPr, Wtag('spacing'))
    spacing.set(Wtag('line'), '480')  # 24pt = 480 twips
    spacing.set(Wtag('lineRule'), 'exact')
    spacing.set(Wtag('before'), '60')
    spacing.set(Wtag('after'), '60')
    # Font properties for text portions
    rPr = etree.SubElement(pPr, Wtag('rPr'))
    sz = etree.SubElement(rPr, Wtag('sz'))
    sz.set(Wtag('val'), '24')  # 12pt
    rFonts = etree.SubElement(rPr, Wtag('rFonts'))
    rFonts.set(Wtag('eastAsia'), '宋体')
    
    # Math object
    oMathPara = etree.SubElement(p, Mtag('oMathPara'))
    oMath = etree.SubElement(oMathPara, Mtag('oMath'))
    
    for elem in math_elements:
        if isinstance(elem, str):
            # Split by operators and create separate runs
            parts = re.split(r'(\s*=\s*)', elem)
            for part in parts:
                if part.strip():
                    oMath.append(m_text(part.strip()))
        else:
            oMath.append(elem)
    
    # Right-aligned formula number using tab
    run_tab = etree.SubElement(p, Wtag('r'))
    rPr_tab = etree.SubElement(run_tab, Wtag('rPr'))
    tab = etree.SubElement(run_tab, Wtag('tab'))
    
    run_num = etree.SubElement(p, Wtag('r'))
    rPr_num = etree.SubElement(run_num, Wtag('rPr'))
    sz_num = etree.SubElement(rPr_num, Wtag('sz'))
    sz_num.set(Wtag('val'), '24')
    rf_num = etree.SubElement(rPr_num, Wtag('rFonts'))
    rf_num.set(Wtag('eastAsia'), '宋体')
    t_num = etree.SubElement(run_num, Wtag('t'))
    t_num.text = f'公式({formula_number})'
    t_num.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    
    # Add tab stop to paragraph properties
    tabs = etree.SubElement(pPr, Wtag('tabs'))
    tab_stop = etree.SubElement(tabs, Wtag('tab'))
    tab_stop.set(Wtag('val'), 'right')
    tab_stop.set(Wtag('pos'), '9072')  # Right margin position
    
    return p

def add_text_paragraph(text, bold=False):
    """Create a simple text paragraph (indented body text)"""
    p = etree.Element(Wtag('p'))
    pPr = etree.SubElement(p, Wtag('pPr'))
    indent = etree.SubElement(pPr, Wtag('ind'))
    indent.set(Wtag('firstLine'), '480')
    spacing = etree.SubElement(pPr, Wtag('spacing'))
    spacing.set(Wtag('line'), '480')
    spacing.set(Wtag('lineRule'), 'exact')
    r = etree.SubElement(p, Wtag('r'))
    rPr = etree.SubElement(r, Wtag('rPr'))
    sz = etree.SubElement(rPr, Wtag('sz'))
    sz.set(Wtag('val'), '24')
    rf = etree.SubElement(rPr, Wtag('rFonts'))
    rf.set(Wtag('eastAsia'), '宋体')
    rf.set(Wtag('ascii'), '宋体')
    if bold:
        b = etree.SubElement(rPr, Wtag('b'))
    t = etree.SubElement(r, Wtag('t'))
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return p

# ============================================================
# Build all formulas
# ============================================================

formula_blocks = []

# ---- 公式(1)(2): 极差标准化 (插入到2.4节P69位置) ----
# Target: after "对所有连续变量进行1%和99%分位的缩尾处理"
f1_text = "设共有m个评价对象（省份）、n个评价指标，采用极差标准化方法对原始数据进行无量纲化处理："

f1 = make_formula_paragraph([
    m_sub('X', 'ij'), "'", " = ", m_frac(
        f'{m_sub("X","ij")} - min({m_sub("X","j")})',
        f'max({m_sub("X","j")}) - min({m_sub("X","j")})'
    )
], 1)
f2_text = "公式(1)适用于正向指标（取值越大越好）。对于逆向指标（如化肥施用强度、农药施用强度），标准化公式为："
f2 = make_formula_paragraph([
    m_sub('X', 'ij'), "'", " = ", m_frac(
        f'max({m_sub("X","j")}) - {m_sub("X","ij")}',
        f'max({m_sub("X","j")}) - min({m_sub("X","j")})'
    )
], 2)

target1 = "对所有连续变量进行1%和99%分位的缩尾处理"
formula_blocks.append((target1, [f1_text, f1, f2_text, f2]))

# ---- 公式(3)(4)(5): 熵权法 (3.1节) ----
target2 = "熵权法通过计算指标的信息熵确定客观权重"
f3 = make_formula_paragraph([
    m_sub('p', 'ij'), " = ", m_frac(m_sub("X'","ij"), m_sum('m','i=1', m_sub("X'","ij")))
], 3)
f4 = make_formula_paragraph([
    m_sub('e', 'j'), " = ", "-", m_frac('1','ln m'), " ", 
    m_sum('m','i=1', f'{m_sub("p","ij")} ln {m_sub("p","ij")}')
], 4)
f5 = make_formula_paragraph([
    m_sub('w', 'j'), " = ", m_frac(f'1 - {m_sub("e","j")}', m_sum('n','j=1', f'(1 - {m_sub("e","j")})'))
], 5)
f345_text = "其中，p_ij为第i个评价对象在第j个指标上的比重，e_j为第j个指标的信息熵。差异系数d_j = 1 - e_j，指标权重w_j由差异系数归一化得到。"
formula_blocks.append((target2, [f345_text, f3, f4, f5]))

# ---- 公式(6)(7)(8): TOPSIS (3.1节) ----
target3 = "第三，TOPSIS法输出的贴近度天然位于"
f6 = make_formula_paragraph([
    m_sub('Z', 'ij'), " = ", m_sub('w', 'j'), " \u00d7 ", m_sub("X'", 'ij')
], 6)
f7 = make_formula_paragraph([
    m_sub('D', 'i'), m_sup('+', ''), " = ", 
    m_sqrt(m_sum('n','j=1', f'({m_sub("Z","ij")} - {m_sub("Z","j")}{m_sup("+","")}){m_sup("2","")}')),
    ",   ",
    m_sub('D', 'i'), m_sup('-', ''), " = ", 
    m_sqrt(m_sum('n','j=1', f'({m_sub("Z","ij")} - {m_sub("Z","j")}{m_sup("-","")}){m_sup("2","")}'))
], 7)
f8 = make_formula_paragraph([
    m_sub('C', 'i'), " = ", m_frac(
        m_text('D_i^-'),
        m_text('D_i^+ + D_i^-')
    ),
    ",   ", m_sub('C', 'i'), " \u2208 [0, 1]"
], 8)
f678_text = "其中Z_ij为加权标准化矩阵，Z_j^+ = max(Z_ij)为正理想解，Z_j^- = min(Z_ij)为负理想解。D_i^+和D_i^-分别为评价对象i到正、负理想解的欧氏距离，C_i为综合贴近度。"
formula_blocks.append((target3, [f678_text, f6, f7, f8]))

# ---- 公式(9)(10)(11)(12): Dagum (3.2节) ----
target4 = "Dagum基尼系数将总体差异G分解为三部分"
f9 = make_formula_paragraph([
    "G = ", m_frac(
        m_sum('k','h=1', m_sum('k','j=1', m_sum('n_h','i=1', m_sum('n_j','r=1', '|y_hi - y_jr|')))),
        m_text('2n\u00b2\u0233')
    )
], 9)
f10 = make_formula_paragraph([
    m_sub('G', 'jj'), " = ", m_frac(
        m_text('\u03a3_i\u03a3_r |y_ji - y_jr|'),
        m_text('2n_j\u00b2 \u0233_j')
    )
], 10)
f11 = make_formula_paragraph([
    m_sub('G', 'jh'), " = ", m_frac(
        m_text('\u03a3_i\u03a3_r |y_ji - y_hr|'),
        m_text('n_j n_h (\u0233_j + \u0233_h)')
    )
], 11)
f12 = make_formula_paragraph([
    "G = ", m_sub('G', 'w'), " + ", m_sub('G', 'nb'), " + ", m_sub('G', 't')
], 12)
f912_text = "其中k为区域组数，n_h和n_j分别为第h和第j区域的省份数，\u0233为全国均值，\u0233_j为第j区域均值。区域内差异G_w=\u03a3G_jj\u00b7p_j\u00b7s_j，其中p_j=n_j/n，s_j=n_j\u0233_j/n\u0233。G_nb为区域间净差异，G_t为超变密度。"
formula_blocks.append((target4, [f912_text, f9, f10, f11, f12]))

# ---- 公式(13): 面板模型 (3.3节) ----
target5 = "基准面板模型设定包含省份和年份双向固定效应"
f13 = make_formula_paragraph([
    m_sub('Y', 'it'), " = \u03b1 + ", m_sub('\u03b2', '1'), "lnEcon", m_sub('','it'), " + ",
    m_sub('\u03b2', '2'), "Fiscal", m_sub('','it'), " + ",
    m_sub('\u03b2', '3'), "Human", m_sub('','it'), " + ",
    m_sub('\u03b2', '4'), "lnDigital", m_sub('','it'), " + ",
    m_sub('\u03bc', 'i'), " + ", m_sub('\u03bb', 't'), " + ", m_sub('\u03b5', 'it')
], 13)
f13_text = "其中Y_it为第i省第t年的乡村振兴TOPSIS得分，\u03bc_i为省份固定效应（吸收不随时间变化的省际异质性），\u03bb_t为年份固定效应（控制全国共同时间趋势），\u03b5_it为随机误差项。模型采用省份聚类稳健标准误以修正异方差和组内序列相关。"
formula_blocks.append((target5, [f13_text, f13]))

# ---- 公式(14): Moran's I (3.3节) ----
target6 = "空间自相关检验通过全局Moran"
f14 = make_formula_paragraph([
    "I = ", m_frac('n', 'S_0'), " \u00d7 ", 
    m_frac(
        m_text('\u03a3_i\u03a3_j w_ij(y_i - \u0233)(y_j - \u0233)'),
        m_text('\u03a3_i (y_i - \u0233)\u00b2')
    )
], 14)
f14_text = "其中S_0 = \u03a3\u03a3w_ij为空间权重之和，w_ij为Queen邻接权重（两省有共同边界则w_ij=1，否则w_ij=0，含海南省与广东/广西的跨海邻接），\u0233为全国均值。I>0表示正向空间自相关。"
formula_blocks.append((target6, [f14_text, f14]))

# ---- 公式(15)(16): SDM (3.3节) ----
target7 = "空间杜宾模型（SDM）同时容纳了被解释变量和解释变量的空间滞后项"
f15 = make_formula_paragraph([
    "Y = ", "\u03c1", "WY + X\u03b2 + WX\u03b8 + \u03b5"
], 15)
f15_text = "其中W为行标准化Queen邻接矩阵，\u03c1为空间自回归系数，WY为被解释变量的空间滞后项，WX\u03b8为解释变量的空间滞后项。模型选择经过规范检验：Lagrange Multiplier检验在5%水平上拒绝无空间自相关的原假设，Hausman检验支持固定效应模型，Wald检验和LR检验均在1%水平上拒绝SDM可简化为SAR(\u03b8=0)或SEM(\u03b8+\u03c1\u03b2=0)的原假设，表明SDM为最优空间计量模型设定。"

target8 = "采用LeSage-Pace效应分解方法将影响分解为直接效应"
f16 = make_formula_paragraph([
    "Y = (I - \u03c1W)", m_sup('-1',''), "(X\u03b2 + WX\u03b8) + (I - \u03c1W)", m_sup('-1',''), "\u03b5"
], 16)
f16_text = "根据LeSage和Pace(2009)，直接效应为偏微分矩阵\u2202Y_i/\u2202X_ik的对角元均值，反映本省解释变量变动对本省乡村振兴水平的影响；间接效应（空间溢出效应）为非对角元均值，反映本省解释变量变动通过空间传导对邻省乡村振兴水平的影响；总效应=直接效应+间接效应。"
formula_blocks.append((target7, [f15_text, f15, f16_text, f16]))
# We need to also add formulas after the SDM text

# ---- 公式(17)(18)(19)(20): 收敛性 (3.4节) ----
target9 = "\u03c3收敛通过变异系数的时序变化判断"
f17 = make_formula_paragraph([
    m_sub('CV', 't'), " = ", m_frac(m_sub('\u03c3', 't'), m_sub('\u0233', 't')), " = ",
    m_frac(
        m_text('\u221a(\u03a3_i (y_it - \u0233_t)\u00b2 / n)'),
        m_sub('\u0233', 't')
    )
], 17)
f17_text = "若CV_{t+1} < CV_t，则存在\u03c3收敛，表明省域乡村振兴的相对差异在时序上缩小。"

target10 = "\u03b2收敛包括绝对\u03b2收敛和条件\u03b2收敛"
f18 = make_formula_paragraph([
    "\u0394ln", m_sub('Y', 'it'), " = \u03b1 + \u03b2 ln", m_sub('Y', 'i,t-1'), " + ", m_sub('\u03b5', 'it')
], 18)
f18_text = "若\u03b2<0且显著，表明存在绝对\u03b2收敛——初始得分越低的省份具有更快的增长速度，省际差距将自发缩小。"

f19 = make_formula_paragraph([
    "\u0394ln", m_sub('Y', 'it'), " = \u03b1 + \u03b2 ln", m_sub('Y', 'i,t-1'), " + \u03b3", m_sub('X', 'it'), " + ", m_sub('\u03bc', 'i'), " + ", m_sub('\u03bb', 't'), " + ", m_sub('\u03b5', 'it')
], 19)

target11 = "政策干预是缩小差距的必要条件"
f20 = make_formula_paragraph([
    "\u03bb = -", m_frac('ln(1+\u03b2)', 'T'), ",   \u03c4 = ", m_frac('ln 2', '\u03bb')
], 20)
f20_text = "其中\u03bb为收敛速度，\u03c4为半衰期（即落后省份与领先省份的差距缩小一半所需的年数），T为样本时间跨度。"
formula_blocks.append((target9, [f17_text, f17, f18_text, f18, f19]))
formula_blocks.append((target11, [f20_text, f20]))

# ---- 公式(21): 障碍度 (3.4节) ----
target12 = "障碍度诊断模型通过计算各指标标准化值与理想值的加权差距"
f21 = make_formula_paragraph([
    m_sub('O', 'ij'), " = ", m_frac(
        m_text('(1 - X\'_ij) \u00d7 w_j'),
        m_text('\u03a3_j (1 - X\'_ij) \u00d7 w_j')
    ), " \u00d7 100%"
], 21)
f21_text = "其中O_ij为第i个省份第j个指标的障碍度，X'_ij为极差标准化值，w_j为熵权权重。维度障碍度等于该维度下所有指标障碍度之和。障碍度越高，表示该指标或维度对乡村振兴水平提升的制约越强。"
formula_blocks.append((target12, [f21_text, f21]))

print(f"共准备 {len(formula_blocks)} 组公式插入点")

# ============================================================
# 读取并修改 document.xml
# ============================================================
xml_path = f'{UNPACK}/word/document.xml'
tree = etree.parse(xml_path)
root = tree.getroot()
body = root.find(f'{{{W}}}body')

# Collect all paragraphs
paragraphs = list(body.findall(f'{{{W}}}p'))

insertions = []  # list of (index_after, elements_to_insert)

for target_text, formula_elements in formula_blocks:
    found = False
    for i, para in enumerate(paragraphs):
        # Get all text from this paragraph
        texts = []
        for t in para.iter(f'{{{W}}}t'):
            if t.text:
                texts.append(t.text)
        full_text = ''.join(texts)
        
        if target_text in full_text:
            # Found the target paragraph - insert formulas AFTER it
            new_elems = []
            for elem in formula_elements:
                if isinstance(elem, str):
                    # It's a text paragraph
                    new_elems.append(add_text_paragraph(elem))
                else:
                    # It's a formula paragraph (already an lxml element)
                    new_elems.append(elem)
            insertions.append((i, new_elems))
            print(f'  找到目标: "{target_text[:30]}..." -> 插入{len(new_elems)}个元素 (在P{i}之后)')
            found = True
            break
    
    if not found:
        print(f'  ⚠️ 未找到: "{target_text[:40]}..."')

# Sort insertions by paragraph index (descending so we can insert without index shifting)
insertions.sort(key=lambda x: -x[0])

# Perform insertions
for idx, elems in insertions:
    ref_para = paragraphs[idx]
    for elem in reversed(elems):
        # Insert after ref_para
        parent = ref_para.getparent()
        ref_index = list(parent).index(ref_para)
        parent.insert(ref_index + 1, elem)

print(f"\n共插入 {sum(len(e) for _, e in insertions)} 个元素")

# Save modified XML
tree.write(xml_path, xml_declaration=True, encoding='UTF-8', standalone=True)

# ============================================================
# Repack to DOCX
# ============================================================
# Remove old file
if os.path.exists(OUT_DOCX):
    os.remove(OUT_DOCX)

# Create new zip
with zipfile.ZipFile(OUT_DOCX, 'w', zipfile.ZIP_DEFLATED) as zout:
    for root_dir, dirs, files in os.walk(UNPACK):
        for file in files:
            file_path = os.path.join(root_dir, file)
            arcname = os.path.relpath(file_path, UNPACK)
            zout.write(file_path, arcname)

print(f"\n✅ 公式插入完成！")
print(f"输出文件: {OUT_DOCX}")

# Clean up
_cleanup_target = _WorkspacePath(UNPACK).resolve()
if RESEARCH_WORKSPACE not in _cleanup_target.parents or _cleanup_target.name != 'unpacked_formulas':
    raise RuntimeError('Refusing cleanup outside the declared workspace')
shutil.rmtree(_cleanup_target)
