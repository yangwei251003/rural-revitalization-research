"""
生成学术简洁风图表：
图1 - 分析框架图（概念级）
图2 - 建模流程图（方法+变量+检验）
"""

from pathlib import Path as _WorkspacePath
import os as _workspace_os
RESEARCH_WORKSPACE = _WorkspacePath(_workspace_os.environ.get("RESEARCH_WORKSPACE", _workspace_os.getcwd())).resolve()

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

OUTPUT_DIR = str(RESEARCH_WORKSPACE / '03_图表文件')

# ============================================================
# 图1: 分析框架图（概念级）
# ============================================================
fig, ax = plt.subplots(1, 1, figsize=(14, 9))
ax.set_xlim(0, 14)
ax.set_ylim(0, 9)
ax.axis('off')
ax.set_facecolor('white')

# 颜色方案：学术黑白 + 浅灰
BOX_STYLE = dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='black', linewidth=1.5)
ARROW_PROPS = dict(arrowstyle='->', lw=1.5, color='black')
DASHED_BOX = dict(boxstyle='round,pad=0.3', facecolor='#f5f5f5', edgecolor='#555555', linewidth=1, linestyle='--')

# 顶层：国家战略
ax.text(7, 8.3, '乡村振兴战略', ha='center', va='center', fontsize=16, fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.4', facecolor='black', edgecolor='black', color='white'))

# 第二层：五个振兴（横向排列）
five_topics = ['产业兴旺', '生态宜居', '乡风文明', '治理有效', '生活富裕']
five_positions = [2.5, 5, 7.5, 10, 12.5]
for i, (topic, pos) in enumerate(zip(five_topics, five_positions)):
    ax.text(pos, 7.0, topic, ha='center', va='center', fontsize=12,
            bbox=BOX_STYLE)
    ax.annotate('', xy=(pos, 7.35), xytext=(7, 7.95),
                arrowprops=dict(arrowstyle='->', lw=1, color='gray', linestyle='dotted'))

# 第三层：指标体系
ax.text(7, 5.8, '乡村振兴综合评价指标体系（14个指标，5个维度）', ha='center', va='center', fontsize=13,
        bbox=dict(boxstyle='round,pad=0.4', facecolor='#e8e8e8', edgecolor='black', linewidth=1.5))

for pos in five_positions:
    ax.annotate('', xy=(7, 6.15), xytext=(pos, 6.6),
                arrowprops=dict(arrowstyle='->', lw=0.8, color='gray', linestyle='dotted'))

# 第四层：三大核心研究问题
core_qs = [
    '核心问题一\n水平测度与时空特征',
    '核心问题二\n差异来源与空间溢出',
    '核心问题三\n收敛性与障碍诊断'
]
core_positions = [3, 7, 11]

methods = [
    ['熵权TOPSIS法', '时序趋势分析', '核密度估计', '空间分布\n可视化'],
    ['Dagum基尼\n系数分解', '面板固定\n效应回归', '空间杜宾\n模型(SDM)', '效应分解\n(LeSage)'],
    ['σ收敛/\nβ收敛检验', '俱乐部\n收敛分析', '障碍度\n诊断模型', '区域对比\n雷达图']
]

for i, (q, pos) in enumerate(zip(core_qs, core_positions)):
    ax.text(pos, 4.7, q, ha='center', va='center', fontsize=11, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='black', linewidth=2))
    ax.annotate('', xy=(pos, 5.15), xytext=(7, 5.45),
                arrowprops=ARROW_PROPS)

    # 方法标签
    for j, m in enumerate(methods[i]):
        ax.text(pos + (j-1.5)*1.2, 3.5, m, ha='center', va='center', fontsize=8,
                bbox=dict(boxstyle='round,pad=0.15', facecolor='#fafafa', edgecolor='#888', linewidth=0.8))
    ax.annotate('', xy=(pos, 4.05), xytext=(pos, 4.35),
                arrowprops=dict(arrowstyle='->', lw=1, color='gray'))

# 底层：政策输出
ax.text(7, 2.2, '服务国家战略的政策建议', ha='center', va='center', fontsize=14, fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.4', facecolor='#333333', edgecolor='black', color='white'))

policies = ['省际协同\n帮扶机制', '差异化\n条件收敛', '聚焦组织\n与产业', '分区精准\n施策']
for i, p in enumerate(policies):
    ax.text(3.5 + i*3.5, 1.2, p, ha='center', va='center', fontsize=10,
            bbox=BOX_STYLE)
    ax.annotate('', xy=(3.5 + i*3.5, 1.55), xytext=(7, 1.85),
                arrowprops=dict(arrowstyle='->', lw=0.8, color='gray', linestyle='dotted'))

ax.text(7, 0.3, '图1 研究分析框架', ha='center', va='center', fontsize=12, fontweight='bold')

plt.tight_layout(pad=0.5)
plt.savefig(f'{OUTPUT_DIR}/图1_分析框架图.png', dpi=200, bbox_inches='tight', facecolor='white')
plt.close()
print("图1 分析框架图已生成")

# ============================================================
# 图2: 建模流程图（方法+变量+检验）
# ============================================================
fig, ax = plt.subplots(1, 1, figsize=(16, 8))
ax.set_xlim(0, 16)
ax.set_ylim(0, 8)
ax.axis('off')
ax.set_facecolor('white')

steps = [
    ('1. 数据准备\n与预处理', '31省×8年面板\n14个指标\n极差标准化\n线性插值填补', (1.5, 5.5)),
    ('2. 熵权TOPSIS\n综合测度', '客观权重计算\n综合/分维度得分\n时序趋势分析\n核密度估计', (4.7, 5.5)),
    ('3. Dagum基尼\n差异分解', '总体基尼系数G\n区域内Gw(26%)\n区域间Gnb(12%)\n超变密度Gt(62%)', (7.9, 5.5)),
    ('4. 面板回归\n因素识别', '双向固定效应\n14指标→得分\nR²=0.974\n逐步回归筛选', (11.1, 5.5)),
    ('5. 空间计量\n溢出分析', 'Moran I检验\nSDM(ρ=0.229)\n效应分解\nWald/LR检验', (14.3, 5.5)),
]

# 下半部分
steps_bottom = [
    ('6. 稳健性\n检验', '缩尾处理\n剔除直辖市\n替换估计方法\nHausman检验', (1.5, 2.5)),
    ('7. 收敛性\n检验', 'σ收敛\n绝对β不收敛\n条件β=-0.605\n俱乐部收敛', (4.7, 2.5)),
    ('8. 障碍度\n诊断', '指标/维度障碍度\nTop3障碍因子\n区域对比雷达图\n星级评分', (7.9, 2.5)),
    ('9. 政策建议\n输出', '实证→指标→建议\n省际帮扶(60%)\n条件收敛(-0.605)\n聚焦组织+产业(73%)', (11.1, 2.5)),
]

for title, detail, (x, y) in steps + steps_bottom:
    # 主框
    ax.text(x, y, title, ha='center', va='center', fontsize=10, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='black', linewidth=1.5))
    # 细节
    ax.text(x, y-1.05, detail, ha='center', va='center', fontsize=7,
            bbox=dict(boxstyle='round,pad=0.2', facecolor='#f8f8f8', edgecolor='#aaa', linewidth=0.8))

# 绘制连接箭头（上半行）
for i in range(len(steps)-1):
    x1 = steps[i][2][0] + 1.0
    x2 = steps[i+1][2][0] - 1.0
    y1 = steps[i][2][1]
    y2 = steps[i+1][2][1]
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', lw=1.5, color='black'))

# 上半行到下半行的箭头
for i, (_, _, (x, _)) in enumerate(steps):
    if i < len(steps_bottom):
        bx = steps_bottom[i][2][0]
        by = steps_bottom[i][2][1]
        ax.annotate('', xy=(bx, by+0.8), xytext=(x, y-1.4),
                    arrowprops=dict(arrowstyle='->', lw=1.2, color='#555', connectionstyle='arc3,rad=0'))

# 下半行连接箭头
for i in range(len(steps_bottom)-1):
    x1 = steps_bottom[i][2][0] + 1.0
    x2 = steps_bottom[i+1][2][0] - 1.0
    y1 = steps_bottom[i][2][1]
    y2 = steps_bottom[i+1][2][1]
    ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                arrowprops=dict(arrowstyle='->', lw=1.5, color='black'))

# 标题
ax.text(8, 7.5, '图2 统计建模与分析流程图', ha='center', va='center', fontsize=14, fontweight='bold')

# 图例
legend_items = [
    (mpatches.FancyBboxPatch((0,0), 0.1, 0.1, boxstyle='round,pad=0.1', facecolor='white', edgecolor='black', linewidth=1.5), '分析步骤'),
    (mpatches.FancyBboxPatch((0,0), 0.1, 0.1, boxstyle='round,pad=0.1', facecolor='#f8f8f8', edgecolor='#aaa', linewidth=0.8), '关键指标/结果'),
]
leg = ax.legend([l[0] for l in legend_items], [l[1] for l in legend_items],
                loc='lower right', fontsize=9, framealpha=0.9, edgecolor='black')

plt.tight_layout(pad=0.5)
plt.savefig(f'{OUTPUT_DIR}/图2_建模流程图.png', dpi=200, bbox_inches='tight', facecolor='white')
plt.close()
print("图2 建模流程图已生成")

print("\n图表生成完成！")
