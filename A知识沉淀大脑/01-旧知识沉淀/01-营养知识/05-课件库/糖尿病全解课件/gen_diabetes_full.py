# -*- coding: utf-8 -*-
"""
糖尿病全解课件 · 流行病学 / 诊断 / 并发症 / 营养调理
风格：Yellow x Black Editorial（黄底黑字 · 时尚杂志风）
风格来源：F技能库skill/01-PPT技能/notebooklm-slide-prompts/SKILL.md
渲染：python-pptx -> .pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

C = {
    'YEL':   RGBColor(0xFF, 0xCC, 0x00),
    'YEL_L': RGBColor(0xFF, 0xE4, 0x7A),
    'YEL_D': RGBColor(0xE0, 0xAE, 0x00),
    'INK':   RGBColor(0x11, 0x11, 0x11),
    'WHITE': RGBColor(0xFF, 0xFF, 0xFF),
    'RED':   RGBColor(0xE0, 0x2B, 0x2B),
    'GRAY':  RGBColor(0x6B, 0x6B, 0x6B),
    'DGRAY': RGBColor(0x3A, 0x3A, 0x3A),
    'CREAM': RGBColor(0xFF, 0xF7, 0xDD),
}
FONT = 'Arial'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
W, H = prs.slide_width, prs.slide_height


def bg(slide, color=None):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color or C['YEL']


def tb(slide, l, t, w, h, txt, sz=18, clr=None, bold=False,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf.vertical_anchor = anchor
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.text = txt
    p.font.size = Pt(sz)
    p.font.color.rgb = clr or C['INK']
    p.font.bold = bold
    p.font.name = FONT
    p.alignment = align
    try:
        p.line_spacing = spacing
    except Exception:
        pass
    return box


def add_p(tf, txt, sz=16, clr=None, bold=False, align=PP_ALIGN.LEFT, sb=Pt(4), sa=Pt(2), spacing=1.0):
    p = tf.add_paragraph()
    p.text = txt
    p.font.size = Pt(sz)
    p.font.color.rgb = clr or C['INK']
    p.font.bold = bold
    p.font.name = FONT
    p.alignment = align
    p.space_before = sb
    p.space_after = sa
    try:
        p.line_spacing = spacing
    except Exception:
        pass
    return p


def textcard(slide, l, t, w, h, lines, fill=None, border=None):
    """lines: list of (txt, sz, color, bold)"""
    box = rect(slide, l, t, w, h, fill=fill, border=border)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.16)
    tf.margin_bottom = Inches(0.12)
    first = True
    for item in lines:
        txt, sz, clr, bold = item
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.text = txt
        p.font.size = Pt(sz)
        p.font.color.rgb = clr
        p.font.bold = bold
        p.font.name = FONT
        p.space_after = Pt(3)
    return box


def rect(slide, l, t, w, h, fill=None, border=None, bw=Pt(2), shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill if fill is not None else C['WHITE']
    if border is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = border
        s.line.width = bw
    s.shadow.inherit = False
    return s


def tag(slide, l, t, txt, bg_c=None, txt_c=None, w=2.1, h=0.36, sz=11):
    s = rect(slide, l, t, Inches(w), Inches(h), fill=bg_c or C['INK'], border=None)
    tf = s.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.size = Pt(sz)
    p.font.color.rgb = txt_c or C['YEL']
    p.font.bold = True
    p.font.name = FONT
    p.alignment = PP_ALIGN.CENTER
    try:
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    except Exception:
        pass
    return s


def new_slide(color=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, color)
    return s


def head(s, num, title, sub='', tagtxt=None):
    tag(s, Inches(0.7), Inches(0.32), tagtxt or ('SLIDE %s' % num))
    tb(s, Inches(0.7), Inches(0.86), Inches(12.2), Inches(0.72), title, sz=30, bold=True)
    rect(s, Inches(0.72), Inches(1.56), Inches(1.9), Inches(0.08), fill=C['INK'], border=None)
    if sub:
        tb(s, Inches(0.72), Inches(1.72), Inches(12.2), Inches(0.4), sub, sz=12.5, clr=C['GRAY'])


def section(s, num, cn, en):
    bg(s, C['INK'])
    rect(s, Inches(0), Inches(0), Inches(0.35), Inches(7.5), fill=C['YEL'], border=None)
    tb(s, Inches(1.0), Inches(1.05), Inches(7), Inches(0.5), 'PART', sz=20, clr=C['YEL'], bold=True)
    tb(s, Inches(0.95), Inches(1.4), Inches(8), Inches(2.6), num, sz=120, clr=C['WHITE'], bold=True)
    rect(s, Inches(1.05), Inches(4.15), Inches(3.4), Inches(0.13), fill=C['YEL'], border=None)
    tb(s, Inches(1.0), Inches(4.45), Inches(10), Inches(0.9), cn, sz=40, clr=C['YEL'], bold=True)
    tb(s, Inches(1.05), Inches(5.5), Inches(10), Inches(0.5), en, sz=15, clr=C['WHITE'])
    for i in range(4):
        rect(s, Inches(11.4), Inches(2.2 + i * 0.7), Inches(1.1), Inches(0.5),
             fill=C['YEL'] if i % 2 == 0 else C['DGRAY'], border=None)


def barchart(slide, left, top, width, items):
    max_v = max(i[1] for i in items) if items else 100
    bh = Inches(0.4)
    gap = Inches(0.18)
    y = top
    for label, val in items:
        ratio = val / max_v
        tb(slide, left, y, Inches(1.7), bh, label, sz=12, clr=C['DGRAY'], bold=True, anchor=MSO_ANCHOR.MIDDLE)
        rect(slide, left + Inches(1.8), y, width - Inches(1.8), bh, fill=C['WHITE'], border=C['INK'], bw=Pt(1.25))
        bar_w = int((width - Inches(1.8)) * ratio)
        rect(slide, left + Inches(1.8), y, bar_w, bh, fill=C['INK'], border=C['INK'], bw=Pt(1.25))
        tb(slide, left + Inches(1.8) + bar_w + Pt(4), y, Inches(1.1), bh, '%d%%' % val,
           sz=12, bold=True, anchor=MSO_ANCHOR.MIDDLE)
        y += bh + gap


# ==================================================================
# 01 · 封面
# ==================================================================
s = new_slide(C['INK'])
rect(s, Inches(0), Inches(0), Inches(0.4), Inches(7.5), fill=C['YEL'], border=None)
rect(s, Inches(0.4), Inches(6.05), Inches(12.93), Inches(0.14), fill=C['YEL'], border=None)
tb(s, Inches(1.0), Inches(1.35), Inches(11), Inches(1.3), '糖尿病全解', sz=76, clr=C['YEL'], bold=True)
tb(s, Inches(1.05), Inches(2.62), Inches(11), Inches(0.6), 'DIABETES MELLITUS', sz=26, clr=C['WHITE'], bold=True)
rect(s, Inches(1.05), Inches(3.35), Inches(4.6), Inches(0.1), fill=C['RED'], border=None)
tb(s, Inches(1.05), Inches(3.6), Inches(11), Inches(0.6),
   '流行病学 · 诊断 · 并发症 · 营养调理', sz=24, clr=C['YEL'], bold=True)
tb(s, Inches(1.05), Inches(4.35), Inches(11), Inches(0.5),
   'Evidence-Based  ·  Clinical Nutrition  ·  Lifestyle Medicine', sz=14, clr=C['GRAY'])
tb(s, Inches(1.05), Inches(5.05), Inches(11), Inches(0.5),
   '一节课讲透：从数字到诊断，从并发症到一张能落地的餐桌', sz=15, clr=C['WHITE'])
tb(s, Inches(1.05), Inches(6.35), Inches(11), Inches(0.5),
   'AI 营养师 · 张医生   |   2026', sz=16, clr=C['YEL'], bold=True)

# ==================================================================
# 02 · 目录
# ==================================================================
s = new_slide()
head(s, '02', 'NAVIGATION', '今天讲四件事，按这个顺序走')
cols = [
    ('01', '流行病学', 'Epidemiology', '全球与中国有多严峻\n谁在名单上', '糖尿病已经不是罕见病，\n它就在我们身边', C['INK']),
    ('02', '诊断', 'Diagnosis', '三条诊断线\n症状 · 前期 · 胰岛功能', '早一天识别，\n多一分逆转机会', C['INK']),
    ('03', '并发症', 'Complications', '急性 vs 慢性\n大血管 · 微血管 · 糖尿病足', '并发症才是糖尿病\n真正的杀伤力', C['RED']),
    ('04', '营养调理', 'Nutrition', '餐盘法则 · 六大要素\n三阶段方案', '把方案落到\n一日三餐里', C['INK']),
]
for i, (num, cn, en, desc, foot, ac) in enumerate(cols):
    x = Inches(0.7 + i * 3.06)
    card = rect(s, x, Inches(2.35), Inches(2.86), Inches(4.35), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    tb(s, x + Inches(0.25), Inches(2.55), Inches(2.4), Inches(1.0), num, sz=48, clr=ac, bold=True)
    tb(s, x + Inches(0.25), Inches(3.55), Inches(2.4), Inches(0.5), cn, sz=24, clr=C['INK'], bold=True)
    tb(s, x + Inches(0.25), Inches(4.05), Inches(2.4), Inches(0.35), en, sz=11, clr=C['GRAY'])
    rect(s, x + Inches(0.25), Inches(4.4), Inches(2.35), Inches(0.05), fill=ac, border=None)
    tb(s, x + Inches(0.25), Inches(4.55), Inches(2.45), Inches(0.9), desc, sz=13, clr=C['DGRAY'], spacing=1.25)
    tb(s, x + Inches(0.25), Inches(5.75), Inches(2.45), Inches(0.75), foot, sz=12, clr=ac, bold=True, spacing=1.15)

# ==================================================================
# 03 · PART 01
# ==================================================================
s = new_slide()
section(s, '01', '流行病学', 'EPIDEMIOLOGY — 糖尿病离我们有多近')

# ==================================================================
# 04 · 中国与全球数据
# ==================================================================
s = new_slide()
head(s, '04', '糖尿病大流行', '数字不会说谎 —— 这不是别人的病', 'PART 01 · 流行病学')
stats = [('11.9%', '中国成人糖尿病患病率', '每 10 个成年人里就有 1 个'),
         ('1.41 亿', '中国糖尿病患者总数', '全球第一，且仍在增长'),
         ('35.2%', '糖尿病前期检出率', '超过 1/3 的人是"后备军"')]
for i, (big, lab, sub) in enumerate(stats):
    x = Inches(0.7 + i * 4.06)
    rect(s, x, Inches(2.3), Inches(3.86), Inches(2.35), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    tb(s, x + Inches(0.25), Inches(2.5), Inches(3.4), Inches(1.0), big, sz=48, clr=C['INK'], bold=True)
    tb(s, x + Inches(0.25), Inches(3.5), Inches(3.4), Inches(0.5), lab, sz=14, clr=C['INK'], bold=True)
    tb(s, x + Inches(0.25), Inches(3.95), Inches(3.4), Inches(0.5), sub, sz=12, clr=C['GRAY'])
band = rect(s, Inches(0.7), Inches(4.95), Inches(11.93), Inches(1.75), fill=C['INK'], border=None)
tf = band.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]
p.text = '记住这三个数字'
p.font.size = Pt(18)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
add_p(tf, '每 10 个成年人中 1 个是糖尿病，3 个是糖尿病前期 —— 合计接近一半的人血糖已经出问题。', sz=14, clr=C['WHITE'])
add_p(tf, '这意味着：糖尿病不是"少数人的病"，而是一个全民级的代谢危机。', sz=14, clr=C['YEL'], bold=True)
add_p(tf, '数据来源：中国成人糖尿病流行病学调查 · IDF 糖尿病地图', sz=11, clr=C['GRAY'])

# ==================================================================
# 05 · 分型与根源
# ==================================================================
s = new_slide()
head(s, '05', '糖尿病分型与根源', '搞清是哪一型，才知道往哪里治', 'PART 01 · 流行病学')
types = [('1 型', '5%', '根源在胰腺', '胰岛β细胞被自身免疫破坏\n胰岛素绝对缺乏，需外源补充', C['RED']),
         ('2 型', '90%', '根源在肝脏', '胰岛素抵抗 + 分泌相对不足\n与生活方式、内脏脂肪密切相关', C['INK']),
         ('妊娠糖尿病', '—', '孕期特有', '妊娠期首次发现糖代谢异常\n产后需长期随访', C['DGRAY']),
         ('特殊类型', '少', '继发因素', '药物、胰腺疾病、内分泌疾病\n等特定病因所致', C['DGRAY'])]
for i, (name, pct, root, desc, ac) in enumerate(types):
    x = Inches(0.7 + i * 3.06)
    rect(s, x, Inches(2.3), Inches(2.86), Inches(3.05), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x, Inches(2.3), Inches(2.86), Inches(0.62), fill=ac, border=None)
    tb(s, x + Inches(0.22), Inches(2.36), Inches(2.5), Inches(0.5), name + '   ' + pct,
       sz=17, clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True)
    tb(s, x + Inches(0.22), Inches(3.1), Inches(2.5), Inches(0.45), root, sz=16, clr=ac, bold=True)
    tb(s, x + Inches(0.22), Inches(3.65), Inches(2.5), Inches(1.6), desc, sz=12.5, clr=C['DGRAY'], spacing=1.25)
band = rect(s, Inches(0.7), Inches(5.6), Inches(11.93), Inches(1.1), fill=C['INK'], border=None)
tf = band.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
tf.margin_top = Inches(0.14)
p = tf.paragraphs[0]
p.text = '本节课重点：2 型糖尿病 —— 占 90%，可防、可控、部分可逆转'
p.font.size = Pt(17)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
add_p(tf, '2 型不是"胰岛素没了"，而是"胰岛素钥匙插不进锁孔"—— 这正是营养与生活方式能发力之处。', sz=13, clr=C['WHITE'])

# ==================================================================
# 06 · 发病机制四步
# ==================================================================
s = new_slide()
head(s, '06', '发病机制：四步链条', '从一顿饭到一次诊断，身体经历了什么', 'PART 01 · 流行病学')
steps = [('STEP 01', '坏油泛滥 · 细胞膜僵硬', 'Omega-6 与 Omega-3 比例严重失衡\n细胞膜弹性丧失，信号接收变差', C['INK']),
         ('STEP 02', '胰岛素抵抗', '受体"门禁"失灵，胰岛素刷不开门\n糖进不去细胞，血糖开始升高', C['RED']),
         ('STEP 03', '代偿性高胰岛素血症', '胰腺拼命分泌更多胰岛素\nβ 细胞长期超负荷、逐渐累瘫', C['INK']),
         ('STEP 04', '胰岛功能衰竭', '细胞内质网应激、β 细胞凋亡\n从"抵抗"走向"衰竭"，确诊糖尿病', C['RED'])]
for i, (lab, title, desc, ac) in enumerate(steps):
    y = Inches(2.3 + i * 1.22)
    rect(s, Inches(0.7), y, Inches(1.75), Inches(0.95), fill=ac, border=None)
    tb(s, Inches(0.7), y + Inches(0.24), Inches(1.75), Inches(0.5), lab, sz=13,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True, align=PP_ALIGN.CENTER)
    rect(s, Inches(2.65), y, Inches(9.98), Inches(0.95), fill=C['WHITE'], border=C['INK'], bw=Pt(1.5))
    tb(s, Inches(2.9), y + Inches(0.1), Inches(4.2), Inches(0.4), title, sz=16, bold=True)
    tb(s, Inches(2.9), y + Inches(0.5), Inches(9.5), Inches(0.4), desc.replace('\n', '   |   '), sz=11.5, clr=C['GRAY'])
    if i < 3:
        tb(s, Inches(6.3), y + Inches(0.95), Inches(0.6), Inches(0.28), '▼', sz=11, clr=C['INK'], bold=True, align=PP_ALIGN.CENTER)

# ==================================================================
# 07 · 高危人群
# ==================================================================
s = new_slide()
head(s, '07', '谁在名单上？', '高危人群自查 —— 符合越多，风险越高', 'PART 01 · 流行病学')
left = rect(s, Inches(0.7), Inches(2.3), Inches(7.3), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = left.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
tf.margin_top = Inches(0.22)
p = tf.paragraphs[0]
p.text = '高危人群清单'
p.font.size = Pt(19)
p.font.color.rgb = C['RED']
p.font.bold = True
p.font.name = FONT
for t in ['年龄 ≥ 40 岁，且长期久坐、缺乏运动',
          '超重或肥胖：BMI ≥ 24，或男性腰围 ≥ 90cm / 女性 ≥ 85cm',
          '一级亲属中有糖尿病患者（父母、兄弟姐妹）',
          '有高血压、高血脂、脂肪肝、高尿酸等代谢问题',
          '女性有妊娠糖尿病史，或曾分娩巨大儿（≥ 4kg）',
          '长期高糖饮食、熬夜、压力大、睡眠不足',
          '体检发现空腹血糖或糖化血红蛋白已偏高']:
    add_p(tf, '▪  ' + t, sz=14, clr=C['DGRAY'], sb=Pt(9))
right = rect(s, Inches(8.3), Inches(2.3), Inches(4.33), Inches(4.4), fill=C['INK'], border=None)
tf = right.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
tf.margin_right = Inches(0.25)
tf.margin_top = Inches(0.3)
p = tf.paragraphs[0]
p.text = '认知重启'
p.font.size = Pt(20)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
for t, c, b in [('糖尿病 ≠ 终身判决书', C['WHITE'], True),
                ('', C['WHITE'], False),
                ('早期 2 型糖尿病，通过科学的生活方式与营养干预，是可以实现临床缓解的。', C['WHITE'], False),
                ('', C['WHITE'], False),
                ('它不是一列只能向前的单向列车，', C['YEL'], True),
                ('而是一条可以科学折返的通道。', C['YEL'], True),
                ('', C['WHITE'], False),
                ('关键前提：发现得早，行动得早。', C['WHITE'], True)]:
    add_p(tf, t, sz=14, clr=c, bold=b, sb=Pt(6))

# ==================================================================
# 08 · PART 02
# ==================================================================
s = new_slide()
section(s, '02', '诊断', 'DIAGNOSIS — 早一天识别，多一分机会')

# ==================================================================
# 09 · 诊断金标准
# ==================================================================
s = new_slide()
head(s, '09', '诊断金标准', '三条线，任何一条踩线都算数', 'PART 02 · 诊断')
criteria = [
    ('空腹血糖', [('≥ 7.0 mmol/L', '糖尿病', C['RED']), ('6.1 - 7.0 mmol/L', '糖尿病前期', C['YEL_D']), ('3.9 - 6.1 mmol/L', '正常', C['INK'])]),
    ('餐后 2 小时血糖', [('≥ 11.1 mmol/L', '糖尿病', C['RED']), ('7.8 - 11.1 mmol/L', '糖尿病前期', C['YEL_D']), ('< 7.8 mmol/L', '正常', C['INK'])]),
    ('糖化血红蛋白 HbA1c', [('≥ 6.5%', '达到诊断标准', C['RED']), ('6.0 - 6.5%', '偏高·需警惕', C['YEL_D']), ('< 6.0%', '理想范围', C['INK'])]),
]
for i, (title, rows) in enumerate(criteria):
    x = Inches(0.7 + i * 4.06)
    rect(s, x, Inches(2.25), Inches(3.86), Inches(3.9), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x, Inches(2.25), Inches(3.86), Inches(0.6), fill=C['INK'], border=None)
    tb(s, x, Inches(2.32), Inches(3.86), Inches(0.5), title, sz=15, clr=C['YEL'], bold=True, align=PP_ALIGN.CENTER)
    y = 3.05
    for val, lab, ac in rows:
        tb(s, x + Inches(0.25), Inches(y), Inches(3.3), Inches(0.45), val, sz=20, clr=ac, bold=True)
        tb(s, x + Inches(0.25), Inches(y + 0.45), Inches(3.3), Inches(0.35), lab, sz=13, clr=C['GRAY'])
        y += 1.0
band = rect(s, Inches(0.7), Inches(6.35), Inches(11.93), Inches(0.75), fill=C['INK'], border=None)
tf = band.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
p = tf.paragraphs[0]
p.text = '诊断需在非同日重复检测确认；有典型症状者一次异常即可提示。具体以临床医师判断为准。'
p.font.size = Pt(13)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
try:
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
except Exception:
    pass

# ==================================================================
# 10 · 典型症状
# ==================================================================
s = new_slide()
head(s, '10', '典型症状：三多一少', '但请记住 —— 多数早期患者没有任何感觉', 'PART 02 · 诊断')
syms = [('多饮', '口渴难耐，饮水量明显增加', C['INK']),
        ('多尿', '夜尿频繁，尿量增多', C['INK']),
        ('多食', '容易饿，食量比以前大', C['INK']),
        ('体重下降', '吃得多反而变瘦', C['RED'])]
for i, (name, desc, ac) in enumerate(syms):
    x = Inches(0.7 + i * 3.06)
    rect(s, x, Inches(2.3), Inches(2.86), Inches(2.1), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x + Inches(0.95), Inches(2.55), Inches(0.95), Inches(0.95),
         fill=ac, border=None, shape=MSO_SHAPE.OVAL)
    tb(s, x + Inches(0.95), Inches(2.82), Inches(0.95), Inches(0.45), str(i + 1), sz=26,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True, align=PP_ALIGN.CENTER)
    tb(s, x, Inches(3.66), Inches(2.86), Inches(0.5), name, sz=20, bold=True, align=PP_ALIGN.CENTER)
    tb(s, x + Inches(0.15), Inches(4.02), Inches(2.56), Inches(0.4), desc, sz=11.5, clr=C['GRAY'], align=PP_ALIGN.CENTER)
warn = rect(s, Inches(0.7), Inches(4.72), Inches(11.93), Inches(1.95), fill=C['INK'], border=None)
tf = warn.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.35)
tf.margin_top = Inches(0.25)
p = tf.paragraphs[0]
p.text = '⚠ 最重要的提醒'
p.font.size = Pt(19)
p.font.color.rgb = C['RED']
p.font.bold = True
p.font.name = FONT
add_p(tf, '2 型糖尿病起病隐匿，绝大多数早期患者没有"三多一少"，等到出现典型症状时，血糖往往已经高了很久。', sz=14, clr=C['WHITE'], sb=Pt(8))
add_p(tf, '所以：40 岁以上、有家族史或代谢问题的人，请把血糖检测列入每年体检。', sz=14, clr=C['YEL'], bold=True)

# ==================================================================
# 11 · 糖尿病前期
# ==================================================================
s = new_slide()
head(s, '11', '糖尿病前期', '最后的可逆窗口 —— 抓住它，就不必走到糖尿病', 'PART 02 · 诊断')
left = rect(s, Inches(0.7), Inches(2.3), Inches(7.0), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = left.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.28)
p = tf.paragraphs[0]
p.text = '什么是糖尿病前期？'
p.font.size = Pt(19)
p.font.color.rgb = C['INK']
p.font.bold = True
p.font.name = FONT
for t in ['血糖已高于正常，但还没达到糖尿病诊断标准。',
          '',
          '身体此时还在"勉强维持平衡"：胰岛素拼命工作，还能把血糖压住。',
          '',
          '这个阶段，胰岛功能尚存、损伤尚可修复 —— 是干预性价比最高的窗口。',
          '',
          '中国有超过 1/3 的成年人正处于这个阶段，却大多毫不知情。']:
    add_p(tf, t, sz=14, clr=C['DGRAY'] if t else C['WHITE'], sb=Pt(6))
right = rect(s, Inches(8.0), Inches(2.3), Inches(4.63), Inches(4.4), fill=C['INK'], border=None)
tf = right.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.3)
p = tf.paragraphs[0]
p.text = '往前一步 vs 往后一步'
p.font.size = Pt(19)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
add_p(tf, '', sz=8)
add_p(tf, '▶ 现在行动', sz=16, clr=C['YEL'], bold=True, sb=Pt(8))
add_p(tf, '饮食结构 + 运动 + 体重管理', sz=13, clr=C['WHITE'])
add_p(tf, '血糖有机会回到正常范围', sz=13, clr=C['WHITE'])
add_p(tf, '', sz=8)
add_p(tf, '▶ 放任不管', sz=16, clr=C['RED'], bold=True, sb=Pt(8))
add_p(tf, '每年约 5%–10% 会进展为糖尿病', sz=13, clr=C['WHITE'])
add_p(tf, '并伴随心脑血管风险上升', sz=13, clr=C['WHITE'])

# ==================================================================
# 12 · C肽
# ==================================================================
s = new_slide()
head(s, '12', '胰岛功能评估：C 肽', '血糖告诉你"结果"，C 肽告诉你"还有多少底牌"', 'PART 02 · 诊断')
lb = rect(s, Inches(0.7), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = lb.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.28)
p = tf.paragraphs[0]
p.text = '空腹 C 肽（参考范围约 0.10 - 0.43 ng/mL）'
p.font.size = Pt(15)
p.font.bold = True
p.font.name = FONT
for val, desc, ac in [('> 0.43', '胰岛素抵抗、代偿性分泌亢进', C['RED']),
                      ('0.10 - 0.43', '功能尚在，干预窗口好', C['INK']),
                      ('< 0.10', '胰岛功能严重受损，需谨慎评估', C['GRAY'])]:
    add_p(tf, '', sz=6)
    add_p(tf, val, sz=22, clr=ac, bold=True)
    add_p(tf, desc, sz=13, clr=C['DGRAY'])
rb = rect(s, Inches(6.8), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = rb.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.28)
p = tf.paragraphs[0]
p.text = '餐后 2 小时 C 肽（应升至空腹的 3 - 5 倍）'
p.font.size = Pt(15)
p.font.bold = True
p.font.name = FONT
for val, desc, ac in [('< 3 倍', '储备功能下降，典型 2 型表现', C['RED']),
                      ('> 5 倍', '提示较明显的胰岛素抵抗', C['RED']),
                      ('几乎不升高', '胰岛反应近乎衰竭', C['GRAY'])]:
    add_p(tf, '', sz=6)
    add_p(tf, val, sz=22, clr=ac, bold=True)
    add_p(tf, desc, sz=13, clr=C['DGRAY'])
add_p(tf, '', sz=10)
add_p(tf, '临床意义：C 肽反映胰岛自身分泌能力，是判断"能否干预、如何干预"的关键指标。', sz=12.5, clr=C['INK'], bold=True)

# ==================================================================
# 13 · PART 03
# ==================================================================
s = new_slide()
section(s, '03', '并发症', 'COMPLICATIONS — 糖尿病真正的杀伤力')

# ==================================================================
# 14 · 并发症总览
# ==================================================================
s = new_slide()
head(s, '14', '并发症全景', '糖尿病本身不致命，致命的是它带来的一切', 'PART 03 · 并发症')
acute = rect(s, Inches(0.7), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = acute.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.26)
p = tf.paragraphs[0]
p.text = '急性并发症 · 来势凶险'
p.font.size = Pt(19)
p.font.color.rgb = C['RED']
p.font.bold = True
p.font.name = FONT
for t, d in [('糖尿病酮症酸中毒', '高血糖 + 酮体堆积，可致昏迷'),
             ('高渗性高血糖状态', '极度脱水、意识障碍，多见于老年人'),
             ('低血糖', '用药不当或进食不足，严重可致昏迷')]:
    add_p(tf, '▪ ' + t, sz=16, clr=C['INK'], bold=True, sb=Pt(14))
    add_p(tf, '    ' + d, sz=12.5, clr=C['GRAY'], sb=Pt(2))
add_p(tf, '', sz=8)
add_p(tf, '特点：起病急、变化快，需要立即就医。', sz=12.5, clr=C['RED'], bold=True, sb=Pt(10))
chronic = rect(s, Inches(6.8), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['INK'], border=None)
tf = chronic.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.26)
p = tf.paragraphs[0]
p.text = '慢性并发症 · 温水煮蛙'
p.font.size = Pt(19)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
for t, d in [('大血管病变', '冠心病、脑卒中、下肢动脉闭塞'),
             ('微血管病变', '视网膜病变、糖尿病肾病、神经病变'),
             ('糖尿病足', '神经 + 血管 + 感染三重打击')]:
    add_p(tf, '▪ ' + t, sz=16, clr=C['YEL'], bold=True, sb=Pt(14))
    add_p(tf, '    ' + d, sz=12.5, clr=C['WHITE'], sb=Pt(2))
add_p(tf, '', sz=8)
add_p(tf, '特点：长期高糖慢慢"腐蚀"全身血管，早期无声。', sz=12.5, clr=C['YEL'], bold=True, sb=Pt(10))

# ==================================================================
# 15 · 大血管并发症
# ==================================================================
s = new_slide()
head(s, '15', '大血管并发症', '高血糖让血管"生锈"，心脑腿首当其冲', 'PART 03 · 并发症')
ves = [('心脏', '冠心病 · 心肌梗死', '糖尿病患者心梗风险显著升高，\n常被称作"冠心病等危症"', C['RED']),
       ('大脑', '脑卒中 · 血管性痴呆', '血管弹性下降、血栓风险上升，\n中风风险明显增加', C['INK']),
       ('下肢', '下肢动脉闭塞症', '走一段路就腿疼（间歇性跛行），\n严重时伤口不愈、面临截肢', C['INK'])]
for i, (organ, name, desc, ac) in enumerate(ves):
    x = Inches(0.7 + i * 4.06)
    rect(s, x, Inches(2.3), Inches(3.86), Inches(3.0), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x, Inches(2.3), Inches(3.86), Inches(0.62), fill=ac, border=None)
    tb(s, x, Inches(2.37), Inches(3.86), Inches(0.5), organ, sz=18,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True, align=PP_ALIGN.CENTER)
    tb(s, x + Inches(0.25), Inches(3.12), Inches(3.36), Inches(0.5), name, sz=16, clr=C['INK'], bold=True)
    tb(s, x + Inches(0.25), Inches(3.7), Inches(3.36), Inches(1.4), desc, sz=13, clr=C['DGRAY'], spacing=1.25)
band = rect(s, Inches(0.7), Inches(5.55), Inches(11.93), Inches(1.15), fill=C['INK'], border=None)
tf = band.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
tf.margin_top = Inches(0.16)
p = tf.paragraphs[0]
p.text = '核心机制：长期高血糖 → 血管内皮损伤 → 动脉粥样硬化加速'
p.font.size = Pt(16)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
add_p(tf, '控糖的意义，不只是"降一个数字"，而是保护全身每一根血管。', sz=13, clr=C['WHITE'])

# ==================================================================
# 16 · 微血管并发症
# ==================================================================
s = new_slide()
head(s, '16', '微血管并发症', '最细的血管最先坏 —— 眼、肾、神经', 'PART 03 · 并发症')
micro = [('糖尿病视网膜病变', '致盲的首要原因', '眼底出血、渗出、视网膜脱离\n早期可无任何自觉症状', C['RED']),
         ('糖尿病肾病', '终末期肾病的主因', '蛋白尿 → 肾功能下降 → 尿毒症\n早期信号：泡沫尿', C['RED']),
         ('糖尿病周围神经病变', '最常见却最被忽视', '手脚麻木、刺痛、烧灼感\n感觉减退易致烫伤、外伤', C['INK'])]
for i, (name, tagline, desc, ac) in enumerate(micro):
    x = Inches(0.7 + i * 4.06)
    rect(s, x, Inches(2.3), Inches(3.86), Inches(3.6), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x, Inches(2.3), Inches(3.86), Inches(0.13), fill=ac, border=None)
    tb(s, x + Inches(0.25), Inches(2.62), Inches(3.36), Inches(0.85), name, sz=18, clr=ac, bold=True, spacing=1.1)
    rect(s, x + Inches(0.25), Inches(3.45), Inches(2.0), Inches(0.34), fill=ac, border=None)
    tb(s, x + Inches(0.25), Inches(3.49), Inches(2.0), Inches(0.3), tagline, sz=11,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True, align=PP_ALIGN.CENTER)
    tb(s, x + Inches(0.25), Inches(3.98), Inches(3.4), Inches(1.7), desc, sz=12.5, clr=C['DGRAY'], spacing=1.3)
band = rect(s, Inches(0.7), Inches(6.05), Inches(11.93), Inches(0.95), fill=C['INK'], border=None)
tf = band.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
p = tf.paragraphs[0]
p.text = '建议：确诊糖尿病后，每年至少做 1 次眼底检查、尿微量白蛋白检测与足部检查。'
p.font.size = Pt(14)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
try:
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
except Exception:
    pass

# ==================================================================
# 17 · 糖尿病足 + 预警自查
# ==================================================================
s = new_slide()
head(s, '17', '糖尿病足 & 预警信号自查', '三分治疗、七分预防 —— 早发现，能保住的不只是脚', 'PART 03 · 并发症')
left = rect(s, Inches(0.7), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = left.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.26)
p = tf.paragraphs[0]
p.text = '糖尿病足：三重打击'
p.font.size = Pt(19)
p.font.color.rgb = C['RED']
p.font.bold = True
p.font.name = FONT
for t in ['① 神经病变 —— 感觉不到疼痛，小伤口被忽略',
          '② 血管病变 —— 供血不足，伤口愈合能力差',
          '③ 细菌感染 —— 高糖环境是细菌的温床',
          '',
          '三者叠加，一个不起眼的水泡或剪甲伤口，',
          '都可能发展成难以愈合的溃疡，严重时面临截肢。',
          '',
          '日常护足：每天检查双脚、穿合脚的鞋袜、',
          '不赤脚行走、不自行修剪厚茧。']:
    add_p(tf, t, sz=13.5, clr=C['DGRAY'] if t else C['WHITE'], bold=t.startswith('糖尿病足'), sb=Pt(6))
right = rect(s, Inches(6.8), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['INK'], border=None)
tf = right.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.26)
p = tf.paragraphs[0]
p.text = '并发症预警信号 · 自查清单'
p.font.size = Pt(19)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
for t in ['视力模糊、看东西变形',
          '手脚麻木、刺痛、蚁爬感',
          '伤口久不愈合、反复感染',
          '小便泡沫多、夜尿增多',
          '皮肤干燥瘙痒、反复长疖肿',
          '胸闷、心慌、活动后气短',
          '足部发凉、走路腿疼休息后缓解']:
    add_p(tf, '▪  ' + t, sz=13.5, clr=C['WHITE'], sb=Pt(11))
add_p(tf, '', sz=8)
add_p(tf, '出现以上任一情况，都建议尽快就医评估。', sz=12.5, clr=C['YEL'], bold=True, sb=Pt(8))

# ==================================================================
# 18 · 底层逻辑
# ==================================================================
s = new_slide()
head(s, '18', '一句话理解并发症', '为什么医生总说：控糖，是在保护血管', 'PART 03 · 并发症')
big = rect(s, Inches(0.7), Inches(2.4), Inches(11.93), Inches(2.3), fill=C['INK'], border=None)
tf = big.text_frame
tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = '糖尿病，本质是一种"全身血管病"'
p.font.size = Pt(30)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
p.alignment = PP_ALIGN.CENTER
add_p(tf, '', sz=6)
add_p(tf, '持续的高血糖，就像把血管泡在糖水里 —— 慢慢腐蚀、慢慢变硬、慢慢堵住。',
       sz=15, clr=C['WHITE'], align=PP_ALIGN.CENTER)
add_p(tf, '从头到脚，凡是有血管的地方，都可能是并发症的战场。', sz=15, clr=C['YEL'], bold=True, align=PP_ALIGN.CENTER)
chain = [('高血糖', C['RED']), ('血管内皮损伤', C['INK']), ('动脉硬化 / 微循环障碍', C['INK']),
         ('器官缺血缺氧', C['INK']), ('并发症发生', C['RED'])]
for i, (t, ac) in enumerate(chain):
    x = Inches(0.7 + i * 2.42)
    rect(s, x, Inches(5.15), Inches(2.1), Inches(1.0), fill=ac, border=None)
    tb(s, x + Inches(0.08), Inches(5.22), Inches(1.94), Inches(0.85), t, sz=12.5,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True, align=PP_ALIGN.CENTER,
       anchor=MSO_ANCHOR.MIDDLE, spacing=1.1)
    if i < 4:
        tb(s, x + Inches(2.1), Inches(5.4), Inches(0.32), Inches(0.5), '▶', sz=13, clr=C['INK'], bold=True, align=PP_ALIGN.CENTER)

# ==================================================================
# 19 · PART 04
# ==================================================================
s = new_slide()
section(s, '04', '营养调理', 'NUTRITION — 把方案落到一日三餐')

# ==================================================================
# 20 · 调理总原则
# ==================================================================
s = new_slide()
head(s, '20', '营养调理总原则', '不是"不能吃"，而是"换一种吃法"', 'PART 04 · 营养调理')
prin = [('01', '控总热量', '把体重和内脏脂肪降下来，\n是改善胰岛素抵抗的第一步'),
        ('02', '稳血糖曲线', '让血糖平稳波动，\n而不是餐后大起大落'),
        ('03', '补修复原料', '给胰岛与代谢系统\n提供它需要的营养底物'),
        ('04', '护肝是关键', '肝脏是血糖的"调度中心"，\n肝功能好，血糖才稳')]
for i, (num, title, desc) in enumerate(prin):
    x = Inches(0.7 + i * 3.06)
    rect(s, x, Inches(2.3), Inches(2.86), Inches(2.55), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    tb(s, x + Inches(0.22), Inches(2.45), Inches(2.4), Inches(0.7), num, sz=34, clr=C['YEL_D'], bold=True)
    tb(s, x + Inches(0.22), Inches(3.15), Inches(2.5), Inches(0.45), title, sz=17, clr=C['INK'], bold=True)
    tb(s, x + Inches(0.22), Inches(3.62), Inches(2.5), Inches(1.1), desc, sz=12.5, clr=C['DGRAY'], spacing=1.25)
band = rect(s, Inches(0.7), Inches(5.1), Inches(11.93), Inches(1.6), fill=C['INK'], border=None)
tf = band.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.35)
tf.margin_top = Inches(0.22)
p = tf.paragraphs[0]
p.text = '为什么是"多因素"调理？'
p.font.size = Pt(18)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
add_p(tf, '血糖失控从来不是单一原因造成的 —— 饮食、体重、肝脏、肠道、微量元素、睡眠压力，环环相扣。', sz=14, clr=C['WHITE'], sb=Pt(8))
add_p(tf, '只盯一个点，效果有限；多因素一起调，才可能把代谢拉回正轨。', sz=14, clr=C['YEL'], bold=True)

# ==================================================================
# 21 · 餐盘法则 + 进餐顺序
# ==================================================================
s = new_slide()
head(s, '21', '控糖餐盘 & 进餐顺序', '两个立刻能用的小改变，餐后血糖好看很多', 'PART 04 · 营养调理')
disc = rect(s, Inches(0.7), Inches(2.3), Inches(4.3), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = disc.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.28)
p = tf.paragraphs[0]
p.text = '控糖餐盘法则'
p.font.size = Pt(19)
p.font.color.rgb = C['INK']
p.font.bold = True
p.font.name = FONT
add_p(tf, '', sz=10)
add_p(tf, '½  非淀粉类蔬菜', sz=20, clr=C['INK'], bold=True, sb=Pt(10))
add_p(tf, '    绿叶菜、菌菇、瓜茄类', sz=12.5, clr=C['GRAY'])
add_p(tf, '¼  优质蛋白', sz=20, clr=C['INK'], bold=True, sb=Pt(12))
add_p(tf, '    鱼、去皮禽肉、蛋、豆制品', sz=12.5, clr=C['GRAY'])
add_p(tf, '¼  粗粮主食', sz=20, clr=C['INK'], bold=True, sb=Pt(12))
add_p(tf, '    全谷物、杂豆、薯类', sz=12.5, clr=C['GRAY'])
order = rect(s, Inches(5.3), Inches(2.3), Inches(7.33), Inches(4.4), fill=C['INK'], border=None)
tf = order.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.35)
tf.margin_top = Inches(0.28)
p = tf.paragraphs[0]
p.text = '进餐顺序：把主食放到最后'
p.font.size = Pt(19)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
add_p(tf, '', sz=10)
seq = [('① 先吃蔬菜', '膳食纤维先铺底，延缓糖吸收'),
       ('② 再吃蛋白质', '蛋白质与脂肪让胃排空变慢'),
       ('③ 最后吃主食', '此时升糖更平缓，峰值更低')]
for t, d in seq:
    add_p(tf, t, sz=18, clr=C['YEL'], bold=True, sb=Pt(14))
    add_p(tf, '     ' + d, sz=13, clr=C['WHITE'], sb=Pt(2))
add_p(tf, '', sz=6)
add_p(tf, '同样一顿饭，换个顺序，餐后血糖峰值可明显降低。', sz=13.5, clr=C['WHITE'], bold=True, sb=Pt(12))

# ==================================================================
# 22 · 主食革命 + 红线
# ==================================================================
s = new_slide()
head(s, '22', '主食革命 & 红线禁区', '该换的换掉，该避的避开', 'PART 04 · 营养调理')
left = rect(s, Inches(0.7), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
tf = left.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.26)
p = tf.paragraphs[0]
p.text = '主食革命：从"快糖"到"慢糖"'
p.font.size = Pt(18)
p.font.color.rgb = C['INK']
p.font.bold = True
p.font.name = FONT
add_p(tf, '', sz=8)
add_p(tf, '✅ 推荐的慢糖主食', sz=15, clr=C['INK'], bold=True, sb=Pt(8))
for t in ['全谷物：糙米、燕麦、荞麦、藜麦', '杂豆类：红豆、绿豆、鹰嘴豆', '薯类：红薯、紫薯、山药']:
    add_p(tf, '   ' + t, sz=13, clr=C['DGRAY'], sb=Pt(4))
add_p(tf, '', sz=8)
add_p(tf, '❌ 需要警惕的快糖主食', sz=15, clr=C['RED'], bold=True, sb=Pt(8))
for t in ['白米饭、白面条、白馒头、白粥', '精制糕点、饼干、即食麦片', '煮得越烂，升糖越快']:
    add_p(tf, '   ' + t, sz=13, clr=C['DGRAY'], sb=Pt(4))
add_p(tf, '', sz=10)
add_p(tf, '原则：粗细搭配，粗粮占一半。', sz=14, clr=C['INK'], bold=True, sb=Pt(8))
right = rect(s, Inches(6.8), Inches(2.3), Inches(5.83), Inches(4.4), fill=C['INK'], border=None)
tf = right.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.32)
tf.margin_top = Inches(0.26)
p = tf.paragraphs[0]
p.text = '红线禁区 · 尽量远离'
p.font.size = Pt(18)
p.font.color.rgb = C['RED']
p.font.bold = True
p.font.name = FONT
add_p(tf, '', sz=8)
for t, d in [('添加糖', '含糖饮料、甜点、糖果、果汁'),
             ('高 GI 水果', '西瓜、荔枝、桂圆、熟香蕉'),
             ('反式脂肪', '人造奶油、起酥油、反复油炸'),
             ('酒精', '干扰肝糖输出，易致血糖波动'),
             ('长期熬夜', '皮质醇升高，直接推高血糖')]:
    add_p(tf, '✕  ' + t, sz=16, clr=C['YEL'], bold=True, sb=Pt(12))
    add_p(tf, '      ' + d, sz=12.5, clr=C['WHITE'], sb=Pt(2))
add_p(tf, '', sz=6)
add_p(tf, '不必"绝对禁止"，但要清楚代价。', sz=12.5, clr=C['YEL'], bold=True, sb=Pt(10))

# ==================================================================
# 23 · 六大营养要素
# ==================================================================
s = new_slide()
head(s, '23', '营养调理的六个发力点', '多因素协同，而不是单打独斗', 'PART 04 · 营养调理')
elems = [('01', '轻断食', '给代谢系统一个休整窗口，\n有助于改善胰岛素敏感性'),
         ('02', '卵磷脂', '帮助脂代谢，\n改善细胞膜与肝脏脂肪堆积'),
         ('03', '护肝', '肝脏是血糖调度中心，\n肝功能稳，血糖才稳'),
         ('04', '高蛋白', '维持肌肉量 —— 肌肉是\n身体最大的"血糖仓库"'),
         ('05', '铬元素', '参与糖代谢，\n是胰岛素发挥作用的辅助因子'),
         ('06', '多维多矿', '补足膳食缺口，\n为代谢修复提供基础原料')]
for i, (num, title, desc) in enumerate(elems):
    col, row = i % 3, i // 3
    x = Inches(0.7 + col * 4.06)
    y = Inches(2.3 + row * 2.4)
    rect(s, x, y, Inches(3.86), Inches(2.15), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x + Inches(0.22), y + Inches(0.22), Inches(0.62), Inches(0.62), fill=C['INK'], border=None)
    tb(s, x + Inches(0.22), y + Inches(0.34), Inches(0.62), Inches(0.4), num, sz=15, clr=C['YEL'],
       bold=True, align=PP_ALIGN.CENTER)
    tb(s, x + Inches(1.0), y + Inches(0.3), Inches(2.7), Inches(0.5), title, sz=18, bold=True)
    tb(s, x + Inches(0.24), y + Inches(1.05), Inches(3.45), Inches(1.0), desc, sz=12.5, clr=C['DGRAY'], spacing=1.25)

# ==================================================================
# 24 · 三阶段流程
# ==================================================================
s = new_slide()
head(s, '24', '三阶段服务流程', '从"压住血糖"到"找回自己的能力"', 'PART 04 · 营养调理')
phases = [('第一阶段', '控糖期', '先稳住局面', ['完善血糖相关检查', '建立饮食 / 运动 / 监测习惯', '把血糖曲线先压平稳'], C['INK']),
          ('第二阶段', '恢复期', '修复与重建', ['持续营养支持', '逐步改善胰岛与代谢功能', '在专业指导下调整用药'], C['RED']),
          ('第三阶段', '逆转期', '稳住成果', ['重建代谢弹性', '恢复正常饮食耐受', '长期随访、防反弹'], C['INK'])]
for i, (label, name, sub, items, ac) in enumerate(phases):
    x = Inches(0.7 + i * 4.06)
    rect(s, x, Inches(2.3), Inches(3.86), Inches(4.4), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x, Inches(2.3), Inches(3.86), Inches(1.35), fill=ac, border=None)
    tb(s, x + Inches(0.25), Inches(2.42), Inches(3.4), Inches(0.4), label, sz=12,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True)
    tb(s, x + Inches(0.25), Inches(2.72), Inches(3.4), Inches(0.5), name, sz=26,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True)
    tb(s, x + Inches(0.25), Inches(3.22), Inches(3.4), Inches(0.35), sub, sz=12.5,
       clr=C['WHITE'] if ac == C['INK'] else C['YEL'])
    for j, it in enumerate(items):
        tb(s, x + Inches(0.25), Inches(3.9 + j * 0.62), Inches(3.45), Inches(0.55), '▪  ' + it,
           sz=13, clr=C['DGRAY'], spacing=1.15)
    if i < 2:
        tb(s, x + Inches(3.86), Inches(4.0), Inches(0.2), Inches(0.5), '▶', sz=14, clr=C['INK'],
           bold=True, align=PP_ALIGN.CENTER)
tb(s, Inches(0.7), Inches(6.85), Inches(11.93), Inches(0.45),
   '两大原则：危重患者不调 · 不信者不调；所有用药调整均需在专业医师指导下进行。',
   sz=12, clr=C['DGRAY'], bold=True, align=PP_ALIGN.CENTER)

# ==================================================================
# 25 · 生活方式六大支柱
# ==================================================================
s = new_slide()
head(s, '25', '六根支柱，缺一不可', '营养只是其中一根 —— 但它是能撬动全局的那根', 'PART 04 · 营养调理')
pillars = [('01', '营养干预', '控制热量 · 稳住血糖曲线'),
           ('02', '运动双引擎', '有氧 + 抗阻，肌肉是血糖仓库'),
           ('03', '睡眠管理', '睡不好，血糖一定不好'),
           ('04', '压力管理', '长期焦虑，血糖跟着波动'),
           ('05', '科学用药', '遵医嘱，阶梯式调整'),
           ('06', '血糖监测', '用数据导航，不凭感觉')]
for i, (num, title, desc) in enumerate(pillars):
    col, row = i % 3, i // 3
    x = Inches(0.7 + col * 4.06)
    y = Inches(2.3 + row * 2.4)
    rect(s, x, y, Inches(3.86), Inches(2.15), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x, y, Inches(3.86), Inches(0.13), fill=C['INK'], border=None)
    tb(s, x + Inches(0.22), y + Inches(0.38), Inches(1.2), Inches(0.7), num, sz=34, clr=C['YEL_D'], bold=True)
    tb(s, x + Inches(1.25), y + Inches(0.52), Inches(2.4), Inches(0.5), title, sz=18, bold=True)
    tb(s, x + Inches(0.24), y + Inches(1.28), Inches(3.45), Inches(0.7), desc, sz=12.5, clr=C['DGRAY'], spacing=1.25)
band = rect(s, Inches(0.7), Inches(6.35), Inches(11.93), Inches(0.75), fill=C['INK'], border=None)
tf = band.text_frame
tf.word_wrap = True
tf.margin_left = Inches(0.3)
p = tf.paragraphs[0]
p.text = '单一维度干预效果有限；六个方向一起走，才能放大改善的空间。'
p.font.size = Pt(14)
p.font.color.rgb = C['YEL']
p.font.bold = True
p.font.name = FONT
try:
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
except Exception:
    pass

# ==================================================================
# 26 · 案例
# ==================================================================
s = new_slide()
head(s, '26', '真实案例 · 数据说话', '以下为个体案例，效果因人而异', 'PART 04 · 营养调理')
cases = [('案例一', '37 岁女性 · 病程 5 年', ['糖化血红蛋白 8.6% → 5.6%', '减重 15.8 kg', '停用全部降糖药', 'C 肽恢复正常，2 年无反弹'], C['INK']),
         ('案例二', '28 岁男性 · 病程 5 年', ['体重 110 kg → 92 kg', '糖化血红蛋白 8.7% → 5.7%', '减重 18.0 kg', '完全停药，持续缓解 3 年'], C['INK']),
         ('案例三', '多项指标改善', ['血糖达标时间（TIR）显著提升', 'C 肽从 0.09 提升至 1.45', '恢复正常碳水耐受'], C['RED'])]
for i, (label, sub, items, ac) in enumerate(cases):
    x = Inches(0.7 + i * 4.06)
    rect(s, x, Inches(2.3), Inches(3.86), Inches(3.6), fill=C['WHITE'], border=C['INK'], bw=Pt(2))
    rect(s, x, Inches(2.3), Inches(3.86), Inches(0.95), fill=ac, border=None)
    tb(s, x + Inches(0.25), Inches(2.42), Inches(3.4), Inches(0.4), label, sz=12,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True)
    tb(s, x + Inches(0.25), Inches(2.72), Inches(3.4), Inches(0.45), sub, sz=14,
       clr=C['YEL'] if ac == C['INK'] else C['WHITE'], bold=True)
    for j, it in enumerate(items):
        tb(s, x + Inches(0.25), Inches(3.5 + j * 0.58), Inches(3.45), Inches(0.5), '▪  ' + it,
           sz=13, clr=C['DGRAY'], spacing=1.15)
tb(s, Inches(0.7), Inches(6.15), Inches(11.93), Inches(0.6),
   '⚠ 案例为个体经验分享，不代表所有人都会获得相同结果，不作为疗效承诺。',
   sz=12.5, clr=C['RED'], bold=True, align=PP_ALIGN.CENTER)

# ==================================================================
# 27 · 总结 + 行动号召
# ==================================================================
s = new_slide(C['INK'])
tag(s, Inches(0.7), Inches(0.35), 'SUMMARY', bg_c=C['YEL'], txt_c=C['INK'])
tb(s, Inches(0.7), Inches(0.9), Inches(12), Inches(0.8), '今天带走这几句话', sz=32, clr=C['YEL'], bold=True)
sums = [('流行病学', '每 10 个成年人就有 1 个糖尿病，3 个前期'),
        ('诊断', '三条线定诊断，C 肽看底牌，越早发现越好'),
        ('并发症', '糖尿病是全身血管病，保护血管才是真控糖'),
        ('营养调理', '换吃法、调顺序、补原料 —— 把方案落到三餐')]
for i, (t, d) in enumerate(sums):
    y = Inches(2.05 + i * 0.82)
    rect(s, Inches(0.7), y, Inches(7.6), Inches(0.68), fill=C['WHITE'], border=None)
    tb(s, Inches(0.95), y + Inches(0.14), Inches(1.9), Inches(0.45), t, sz=14, clr=C['INK'], bold=True)
    tb(s, Inches(2.9), y + Inches(0.16), Inches(5.3), Inches(0.45), d, sz=13, clr=C['DGRAY'])
quote = rect(s, Inches(8.6), Inches(2.05), Inches(4.03), Inches(3.32), fill=C['YEL'], border=None)
tf = quote.text_frame
tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = '糖尿病给我们的\n不是一个句号，\n而是一次重新\n认识身体的机会。'
p.font.size = Pt(22)
p.font.color.rgb = C['INK']
p.font.bold = True
p.font.name = FONT
p.alignment = PP_ALIGN.CENTER
cta = rect(s, Inches(0.7), Inches(5.6), Inches(11.93), Inches(0.95), fill=C['YEL'], border=None)
tf = cta.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = '行动，从下一顿饭开始 —— 先吃菜，再吃肉，最后吃主食。'
p.font.size = Pt(20)
p.font.color.rgb = C['INK']
p.font.bold = True
p.font.name = FONT
p.alignment = PP_ALIGN.CENTER
try:
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
except Exception:
    pass
tb(s, Inches(0.7), Inches(6.75), Inches(11.93), Inches(0.55),
   '以上为健康科普分享，不构成诊疗建议，个体情况请咨询专业医师。   |   感谢聆听',
   sz=11.5, clr=C['GRAY'], align=PP_ALIGN.CENTER)

# ==================================================================
out_dir = '/Users/mac/Documents/zfc最强大脑/A知识沉淀大脑/01-旧知识沉淀/01-营养知识/营养课件/糖尿病全解课件'
out = out_dir + '/糖尿病全解课件-流行病学·诊断·并发症·营养调理.pptx'
prs.save(out)
print('PPT saved to:', out)
print('Total slides:', len(prs.slides))
