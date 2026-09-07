#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用中文简历生成器（也可用于英文简历：把模块标题/内容写成英文即可）。

用法：
    pip install python-docx
    python3 generate_resume.py sample_config.json      # 生成配置中 output 指定的 .docx

配置文件字段（JSON，无注释；字段都可省略，省略即不输出对应部分）：
{
  "output": "输出文件名.docx",
  "info": {
    "name": "姓名（大字号居中标题）",
    "meta": "性别/出生/城市 等一行",
    "contact": "电话/邮箱/微信 等一行",
    "intent": "求职意向一行（加粗高亮）"
  },
  "photo": "可选：照片路径（放入后右上角显示，建议先压缩到几百KB）",
  "photo_width_cm": 3.0,
  "style": {"font": "微软雅黑", "size": 10.5, "heading_color": "#1F4E79"},
  "sections": [
    {"title": "模块标题（如：教育背景）", "entries": [
        {"left": "主标题", "right": "右侧日期/时间", "sub": "副标题/角色",
         "bullets": ["要点1", "要点2"]}
    ]}
  ]
}

"bullets" 为列表；没有内容的字段省略即可（见 sample_config.json）。
生成结果：A4、中文字体（微软雅黑默认）、深色小标题带下划线、条目"左标题右日期"。
"""
import json
import os
import sys

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DEFAULT_STYLE = {"font": "微软雅黑", "size": 10.5, "heading_color": "#1F4E79"}
CONTENT_WIDTH_CM = 17.2  # 21cm - 左右边距 1.9cm*2


def _hex_to_rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_run(run, text=None, size=None, bold=None, color=None, font=None):
    if text is not None:
        run.text = text
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font or '微软雅黑')
    run.font.name = font or '微软雅黑'
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def bottom_border(p, sz='6', color_hex='1F4E79'):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single')
    b.set(qn('w:sz'), sz)
    b.set(qn('w:space'), '2')
    b.set(qn('w:color'), color_hex.lstrip('#'))
    pBdr.append(b)
    pPr.append(pBdr)


def add_header(doc, info, style, photo, photo_width_cm):
    """抬头：左信息 + 右照片（无边框表格）；下加分割线。"""
    name = info.get('name', '')
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.allow_autofit = False
    left_cell, right_cell = table.rows[0].cells
    left_cell.width = Cm(13.0)
    right_cell.width = Cm(CONTENT_WIDTH_CM - 13.0)
    left_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    p = left_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(2)
    set_run(p.add_run(), text=name, size=20, bold=True, font=style['font'])

    for key in ('meta', 'contact', 'intent'):
        val = info.get(key)
        if not val:
            continue
        p = left_cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(2 if key != 'intent' else 0)
        if key == 'intent':
            set_run(p.add_run(), text=val, size=style['size'], bold=True,
                    color=_hex_to_rgb(style['heading_color']), font=style['font'])
        else:
            set_run(p.add_run(), text=val, size=style['size'], font=style['font'])

    if photo and os.path.exists(photo):
        p = right_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(photo, width=Cm(photo_width_cm or 3.0))

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(0)
    bottom_border(p, sz='8', color_hex=style['heading_color'])


def add_section(doc, title, style):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(2)
    set_run(p.add_run(), text=title, size=style['size'] + 1.5, bold=True,
            color=_hex_to_rgb(style['heading_color']), font=style['font'])
    bottom_border(p, sz='6', color_hex=style['heading_color'])


def add_entry(doc, e, style):
    left = e.get('left', '')
    right = e.get('right', '')
    sub = e.get('sub', '')
    bullets = e.get('bullets') or []

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(CONTENT_WIDTH_CM),
                                             WD_TAB_ALIGNMENT.RIGHT)
    if left:
        set_run(p.add_run(), text=left, size=style['size'], bold=True, font=style['font'])
    if right:
        set_run(p.add_run(), text='\t' + right, size=style['size'], font=style['font'])
    if sub:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_after = Pt(1)
        r = p2.add_run(sub)
        set_run(r, size=style['size'] - 0.5, font=style['font'])
        r.font.italic = True
    for b in bullets:
        pb = doc.add_paragraph()
        pb.paragraph_format.left_indent = Cm(0.4)
        pb.paragraph_format.space_after = Pt(1)
        set_run(pb.add_run(), text='• ' + b, size=style['size'], font=style['font'])


def build(cfg):
    style = {**DEFAULT_STYLE, **(cfg.get('style') or {})}
    info = cfg.get('info') or {}
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(1.5)
        s.bottom_margin = Cm(1.5)
        s.left_margin = Cm(1.9)
        s.right_margin = Cm(1.9)
    st = doc.styles['Normal']
    st.font.name = style['font']
    st.font.size = Pt(style['size'])
    rPr = st.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), style['font'])

    add_header(doc, info, style, cfg.get('photo'), cfg.get('photo_width_cm'))
    for sec in cfg.get('sections') or []:
        add_section(doc, sec['title'], style)
        for e in sec.get('entries') or []:
            add_entry(doc, e, style)

    out = cfg.get('output', 'resume.docx')
    doc.save(out)
    print('已生成:', os.path.abspath(out))
    total_bullets = sum(len(e.get('bullets') or [])
                        for s in cfg.get('sections') or [] for e in s.get('entries') or [])
    print('模块数: %d | 条目数: %d | 要点数: %d'
          % (len(cfg.get('sections') or []),
             sum(len(s.get('entries') or []) for s in cfg.get('sections') or []),
             total_bullets))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    cfg_path = sys.argv[1]
    with open(cfg_path, encoding='utf-8') as f:
        cfg = json.load(f)
    build(cfg)


if __name__ == '__main__':
    main()
