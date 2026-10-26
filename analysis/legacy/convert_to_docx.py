"""将修订版论文 Markdown 转换为格式化的 DOCX"""

from pathlib import Path as _WorkspacePath
import os as _workspace_os
RESEARCH_WORKSPACE = _WorkspacePath(_workspace_os.environ.get("RESEARCH_WORKSPACE", _workspace_os.getcwd())).resolve()

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import re, os

# 读取 Markdown
with open(str(RESEARCH_WORKSPACE / '论文修订版_完整.md'), 'r', encoding='utf-8') as f:
    content = f.read()

# 解析 Markdown 为结构化段落
lines = content.strip().split('\n')

doc = Document()

# 设置默认字体
style = doc.styles['Normal']
font = style.font
font.name = '宋体'
font.size = Pt(12)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
style.paragraph_format.line_spacing = 1.5
style.paragraph_format.space_after = Pt(0)
style.paragraph_format.first_line_indent = Cm(0.74)

# 设置页边距
for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)

current_section = None
in_table = False
table_rows = []
in_code_block = False

def add_heading(doc, text, level, bold=True):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    if level == 0:  # 论文标题
        run.font.size = Pt(16)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif level == 1:  # 章标题
        run.font.size = Pt(14)
        run.font.name = '黑体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    elif level == 2:  # 节标题
        run.font.size = Pt(13)
        run.font.name = '黑体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
    elif level == 3:
        run.font.size = Pt(12)
        run.font.name = '楷体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '楷体')
        run.bold = True
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)

def add_para(doc, text, bold=False, indent=True):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    else:
        p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run(text)
    run.font.size = Pt(12)
    run.font.name = '宋体'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    if bold:
        run.bold = True
    return p

def add_table(doc, rows_data, headers=None):
    """添加表格"""
    if not rows_data:
        return
    ncols = len(rows_data[0])
    table = doc.add_table(rows=len(rows_data), cols=ncols)
    table.style = 'Table Grid'
    
    # Set cell content
    for i, row in enumerate(rows_data):
        for j, cell_text in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = cell_text.strip()
            # Set font
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
                    run.font.name = '宋体'
                    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    
    # Add caption before table
    if headers:
        add_para(doc, headers[0], bold=True, indent=False)
    
    doc.add_paragraph()  # space after table

# 解析并生成文档
i = 0
current_table = []
table_caption = None

while i < len(lines):
    line = lines[i]
    
    # 跳过空行
    if not line.strip():
        i += 1
        continue
    
    # 标题
    if line.startswith('# ') and not line.startswith('## '):
        # 论文标题
        text = line[2:].strip()
        if '目录' in text:
            pass  # Skip TOC, use Word auto-TOC
        else:
            add_heading(doc, text, 0)
            if text == '摘要' or 'Abstract' in text:
                add_para(doc, '■', indent=False)  # separator
        i += 1
        if current_table:
            add_table(doc, current_table, table_caption)
            current_table = []
            table_caption = None
        continue
    
    elif line.startswith('## '):
        text = line[3:].strip()
        add_heading(doc, text, 1)
        i += 1
        if current_table:
            add_table(doc, current_table, table_caption)
            current_table = []
            table_caption = None
        continue
    
    # 表格行
    elif '|' in line and not line.startswith('>'):
        # Table separator/meta lines
        if re.match(r'^\|[\s\-:|]+\|$', line):
            if current_table and len(current_table) > 0:
                table_caption_text = None
            i += 1
            continue
        
        # Parse table row
        cells = [c.strip() for c in line.split('|')[1:-1]]
        if cells:
            current_table.append(cells)
        i += 1
        continue
    
    # 注释行
    elif line.startswith('>'):
        text = line[1:].strip()
        add_para(doc, text, indent=False)
        i += 1
        continue
    
    # 空白/分隔
    elif line.startswith('---'):
        # Check if previous content was table caption
        i += 1
        continue
    
    # 带图表标注的行
    elif line.startswith('**图') or line.startswith('**表'):
        add_para(doc, line.strip('*'), bold=True, indent=False)
        i += 1
        continue
    
    # 参考文献
    elif line.startswith('[') and ']' in line[:5]:
        add_para(doc, line, indent=False)
        i += 1
        continue
    
    # 普通段落
    else:
        # Check for bold text
        if line.startswith('**') and line.endswith('**'):
            add_para(doc, line[2:-2], bold=True)
        else:
            # Handle inline formatting
            text = line.strip()
            # Remove bold markers for inline
            text = re.sub(r'\*\*', '', text)
            add_para(doc, text)
        i += 1
        if current_table:
            add_table(doc, current_table, table_caption)
            current_table = []
            table_caption = None
        continue

# Add remaining table
if current_table:
    add_table(doc, current_table, table_caption)

# Save
output_path = str(RESEARCH_WORKSPACE / '论文修订版_完整_新.docx')
doc.save(output_path)
print(f"论文已保存: {output_path}")

# 清理临时文件
# os.remove(r'<local-workspace>\乡村振兴数据\论文修订版_完整.md')
print("完成！")
