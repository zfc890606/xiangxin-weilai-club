#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""张医生-事业指导-20261108-成事之道与业务路径.pptx 生成脚本
版式：16:9 · 米白底 · 砖红主色 · 无边框阴影 · 三档字号
来源：A知识沉淀大脑/00-新知识入库/逐字稿-安利成事之道与业务路径梳理.md
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os

# ---------- 设计令牌 ----------
RED   = RGBColor(0xD0, 0x02, 0x1B)
INK   = RGBColor(0x1F, 0x1F, 0x1F)
GRAY  = RGBColor(0x6B, 0x6B, 0x6B)
LIGHT = RGBColor(0xF6, 0xF4, 0xF1)
BG    = RGBColor(0xFC, 0xFB, 0xF9)
LINE  = RGBColor(0xE2, 0xDD, 0xD8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "苹方-简"
W, H = 13.333, 7.5
ML, MR = 0.95, 0.95
CW = W - ML - MR
FOOTER = "张医生 · AI营养师｜成事之道与安利业务路径 · 2026-11-08"

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]

TOTAL = 30


def _ea(run):
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', FONT)


def txt(slide, x, y, w, h, text, size=18, bold=False, color=INK,
        align=PP_ALIGN.LEFT, spacing=1.25, anchor=MSO_ANCHOR.TOP, margin=0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = FONT
        _ea(r)
    return box


def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE,
         radius=None, line_w=1.0):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    s.shadow.inherit = False
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        try:
            s.adjustments[0] = radius
        except Exception:
            pass
    return s


def page(slide, num):
    rect(slide, 0, 0, W, H, fill=BG)
    rect(slide, ML, H - 0.62, CW, 0.012, fill=LINE)
    txt(slide, ML, H - 0.55, 8.2, 0.3, FOOTER, 10, False, GRAY)
    txt(slide, W - MR - 1.2, H - 0.55, 1.2, 0.3, f"{num:02d} / {TOTAL}", 10,
        False, GRAY, PP_ALIGN.RIGHT)


def kicker(slide, text, y=0.62):
    rect(slide, ML, y + 0.06, 0.16, 0.16, fill=RED)
    txt(slide, ML + 0.32, y, 7.0, 0.3, text, 12, True, RED)


def title(slide, text, y=1.0, size=36, w=None):
    txt(slide, ML, y, w or CW, 1.5, text, size, True, INK, spacing=1.3)


def card(slide, x, y, w, h, head, body, head_size=19, body_size=14, badge=None):
    rect(slide, x, y, w, h, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    rect(slide, x, y + 0.22, 0.05, h - 0.44, fill=RED)
    ty = y + 0.24
    if badge:
        txt(slide, x + 0.32, ty, w - 0.6, 0.26, badge, 10.5, True, RED)
        ty += 0.32
    head_h = 0.46 if "\n" not in head else 0.46 + head_size * 0.018
    txt(slide, x + 0.32, ty, w - 0.62, head_h, head, head_size, True, INK, spacing=1.15)
    txt(slide, x + 0.32, ty + head_h + 0.08, w - 0.62,
        h - (ty - y) - head_h - 0.28, body, body_size, False, GRAY, spacing=1.3)


def quote(slide, text, note=None, size=30):
    rect(slide, ML, 2.0, 0.07, 2.6, fill=RED)
    txt(slide, ML + 0.55, 2.02, CW - 0.8, 1.8, text, size, True, INK, spacing=1.4)
    if note:
        txt(slide, ML + 0.55, 4.42, CW - 0.8, 0.5, note, 15, False, GRAY)


def divider(slide, num, cn, en, num_slide):
    rect(slide, 0, 0, W, H, fill=INK)
    txt(slide, ML, 1.32, 4.2, 1.95, num, 88, True, RED)
    rect(slide, ML + 0.05, 3.35, 1.5, 0.05, fill=RED)
    txt(slide, ML, 3.62, CW, 0.9, cn, 40, True, WHITE)
    txt(slide, ML, 4.62, CW, 0.5, en, 14, False, RGBColor(0x9A, 0x9A, 0x9A))
    txt(slide, W - MR - 1.2, H - 0.55, 1.2, 0.3, f"{num_slide:02d} / {TOTAL}", 10,
        False, RGBColor(0x7A, 0x7A, 0x7A), PP_ALIGN.RIGHT)


# ============ 01 封面 ============
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=BG)
rect(s, 0, 0, 0.42, H, fill=RED)
txt(s, ML + 0.3, 1.62, 9.0, 0.35, "张医生 · AI 营养师｜2026-11-08 伙伴辅导", 13, True, RED)
txt(s, ML + 0.3, 2.15, 10.6, 2.2, "成事之道\n与安利业务路径", 54, True, INK, spacing=1.28)
rect(s, ML + 0.35, 4.72, 2.0, 0.05, fill=RED)
txt(s, ML + 0.3, 5.02, 10.0, 1.0,
    "先立一条主线，学习才不是盲目堆料\n业务路径只有一条：基础工作 → 关键工作 → 2000 分 → 复制 2000 分",
    16, False, GRAY, spacing=1.55)
txt(s, W - MR - 2.6, H - 0.85, 2.6, 0.3, "逐字稿整理 · 配套课件", 11, False, GRAY, PP_ALIGN.RIGHT)

# ============ 02 今天只讲两件事 ============
s = prs.slides.add_slide(BLANK); page(s, 2)
kicker(s, "本场结构")
title(s, "30 分钟，只讲两件事", 1.0, 40)
card(s, ML, 2.25, CW / 2 - 0.18, 3.2, "一 · 成事之道\n找到你的价值点",
     "做成事不复杂：做对一件核心的事。\n\n先把价值点挖出来，再聚焦做到极致，\n然后用这个时代的语言让别人看见。",
     head_size=21, body_size=14.5, badge="PART 01")
card(s, ML + CW / 2 + 0.18, 2.25, CW / 2 - 0.18, 3.2, "二 · 业务路径\n2000 分与复制",
     "基础工作搭桥 → 关键工作收单 → 2000 分落地。\n\nOPP 找人，系统接住，复制 4 个 2000 分就是银章。",
     head_size=21, body_size=14.5, badge="PART 02")
txt(s, ML, 5.78, CW, 0.5, "路径清楚了，认知到位了，人才会真正用心去做。", 17, True, RED)

# ============ 03 分隔 01 ============
s = prs.slides.add_slide(BLANK); divider(s, "01", "成事之道", "先立主线，再谈方法", 3)

# ============ 04 主线 ============
s = prs.slides.add_slide(BLANK); page(s, 4)
kicker(s, "开场问题")
title(s, "不停在学，却说不出为什么学", 0.95, 36)
txt(s, ML, 2.05, CW, 0.5, "这是很多伙伴的共性：课没少上，人却越来越乱。", 15, False, GRAY)
for i, (tag, head, body) in enumerate([
    ("先定主线", "一条主线，就是一把筛子", "哪些该学、学了对自己这条路有没有用，\n一筛就清楚了。"),
    ("没有主线", "东学西学，越学越乱", "方法堆了一大堆，思路反而模糊，\n最后陷入迷茫。"),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    hi = (i == 0)
    rect(s, x, 2.65, wd, 2.15, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    rect(s, x, 2.65, wd, 0.58, fill=(RED if hi else LIGHT))
    txt(s, x + 0.3, 2.79, wd - 0.6, 0.3, tag, 12, True, (WHITE if hi else GRAY))
    txt(s, x + 0.3, 3.38, wd - 0.6, 0.4, head, 19, True, (RED if hi else INK))
    txt(s, x + 0.3, 3.9, wd - 0.6, 0.8, body, 13.5, False, GRAY, spacing=1.4)
txt(s, ML, 5.15, CW, 0.5, "先立起来这条线，学习才不至于变成盲目堆料。", 17, True, RED)
txt(s, ML, 5.75, CW, 0.4, "所以今天第一件事，不是讲方法，是帮你把这条主线梳理出来。", 14, False, GRAY)

# ============ 05 成事之道 ============
s = prs.slides.add_slide(BLANK); page(s, 5)
kicker(s, "核心命题")
title(s, "做成事，其实一点都不复杂", 0.95, 40)
txt(s, ML, 2.0, CW, 0.4, "简单到 —— 做对一件核心的事。", 16, True, GRAY)
flow = [
    ("①", "找到价值点", "回顾你真正做成过的事，\n把它挖出来。"),
    ("②", "用帮助别人来体现", "安利、自媒体都是载体，\n价值要落到人身上。"),
    ("③", "结果自然发生", "财富和社会认同，\n是价值被认可之后的结果。"),
]
for i, (n, hd, bd) in enumerate(flow):
    x = ML + i * (CW / 3 + 0.05)
    wd = CW / 3 - 0.2
    rect(s, x, 2.55, wd, 1.95, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.28, 2.74, wd - 0.56, 0.3, n, 12, True, RED)
    txt(s, x + 0.28, 3.06, wd - 0.56, 0.45, hd, 18, True, INK)
    txt(s, x + 0.28, 3.6, wd - 0.56, 0.75, bd, 13, False, GRAY, spacing=1.35)
    if i < 2:
        txt(s, x + wd, 2.55, 0.2, 1.95, "›", 20, True, RGBColor(0xC9, 0xC4, 0xBE),
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, ML, 4.85, CW, 1.25, fill=INK)
txt(s, ML + 0.4, 5.02, CW - 0.8, 0.9,
    "外在的方法和技巧，最后都要回到这条主线上来。\n主线不清，方法越多，人越乱。",
    18, True, WHITE, spacing=1.45)

# ============ 06 价值点藏在做成过的事里 ============
s = prs.slides.add_slide(BLANK); page(s, 6)
kicker(s, "挖掘原则")
title(s, "你的价值点，藏在你做成过的事里", 0.95, 32)
txt(s, ML, 1.95, CW, 0.4, "方法只有一个：回顾自己过去，真正做成的那一件事。", 14, False, GRAY)
items = [
    ("腰突改善了", "身体这件事，你自己走通过一遍", "健康方案"),
    ("把娃带到了 985", "教育这件事，你拿到了结果", "家庭教育"),
    ("减肥成功了", "体重这件事，你亲身验证过", "体重管理"),
]
cw3 = (CW - 0.44) / 3
for i, (h1, body, prod) in enumerate(items):
    x = ML + i * (cw3 + 0.22)
    rect(s, x, 2.5, cw3, 2.55, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.045)
    txt(s, x + 0.3, 2.75, cw3 - 0.6, 0.6, h1, 19, True, INK, spacing=1.2)
    txt(s, x + 0.3, 3.45, cw3 - 0.6, 0.85, body, 13, False, GRAY, spacing=1.4)
    txt(s, x + 0.3, 4.45, cw3 - 0.6, 0.3, "↓  可对应", 11, True, RED)
    txt(s, x + 0.3, 4.72, cw3 - 0.6, 0.3, prod, 14, True, RED)
txt(s, ML, 5.35, CW, 0.5, "这一步，其实就是做商业定位。", 18, True, RED)
txt(s, ML, 5.95, CW, 0.4, "人人都有价值，区别只在于：你挖没挖出来。", 14, False, GRAY)

# ============ 07 做到极致 ============
s = prs.slides.add_slide(BLANK); page(s, 7)
kicker(s, "难点在这里")
title(s, "挑一个挂得上钩的，做到极致", 0.95, 36)
card(s, ML, 2.15, CW / 2 - 0.16, 2.55, "减肥成功 ≠ 做到极致",
     "黄医生减下来了，但还不够。\n真做减肥赛道，练出腹肌马甲线，\n那才叫厉害、才拿得出手。",
     head_size=20, body_size=14.5, badge="举例")
card(s, ML + CW / 2 + 0.16, 2.15, CW / 2 - 0.16, 2.55, "遍地都是，凭什么成功是你",
     "你减了 20 斤又怎么样？遍地都是减了 20 斤的。\n普通，就是没有区别。",
     head_size=20, body_size=14.5, badge="扎心一问")
txt(s, ML, 4.95, CW, 0.4, "做到极致的那些人，才站到了台上：", 15, True, RED)
for i, (n, t) in enumerate([("李佳琦", "把表达做到极致"), ("小杨哥", "把内容做到极致"), ("真正的企业家", "把一个点做到极致")]):
    x = ML + i * (CW / 3 + 0.02)
    rect(s, x, 5.45, CW / 3 - 0.18, 0.82, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    txt(s, x + 0.26, 5.58, CW / 3 - 0.5, 0.3, n, 14, True, INK)
    txt(s, x + 0.26, 5.9, CW / 3 - 0.5, 0.3, t, 11.5, False, GRAY)

# ============ 08 金句 ============
s = prs.slides.add_slide(BLANK); page(s, 8)
kicker(s, "一句话")
quote(s, "赢一次，\n就可以赢一辈子。",
      "做成一件极致的事，你自己就是广告。", 36)
txt(s, ML + 0.55, 5.15, CW - 0.8, 0.9,
    "难的不是做，是先挖出那个价值点，再把它做到极致。", 17, True, RED)

# ============ 09 链接人 ============
s = prs.slides.add_slide(BLANK); page(s, 9)
kicker(s, "第二步 · 链接人")
title(s, "普通人链接人，靠的是被看见", 0.95, 34)
for i, (tag, head, body, hi) in enumerate([
    ("有资源的人", "天然能链接", "认识商会会长、秘书长，\n有价值又有资源，成功很简单。", False),
    ("我们普通人", "先让更多人看见你", "没有资源，就把自己放出去。\n看见你，才听得到你的价值。", True),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    rect(s, x, 2.2, wd, 2.35, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    rect(s, x, 2.2, wd, 0.58, fill=(RED if hi else LIGHT))
    txt(s, x + 0.3, 2.34, wd - 0.6, 0.3, tag, 12, True, (WHITE if hi else GRAY))
    txt(s, x + 0.3, 2.92, wd - 0.6, 0.45, head, 20, True, (RED if hi else INK))
    txt(s, x + 0.3, 3.48, wd - 0.6, 0.95, body, 13.5, False, GRAY, spacing=1.4)
rect(s, ML, 4.9, CW, 1.35, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML + 0.38, 5.05, CW - 0.8, 0.45, "每个时代，都有自己的语言和工具。", 20, True, INK)
txt(s, ML + 0.38, 5.62, CW - 0.8, 0.45, "当下这个时代的语言，就是自媒体 —— 懂它，别人才听得到你。", 15, False, GRAY)

# ============ 10 自媒体也要做到极致 ============
s = prs.slides.add_slide(BLANK); page(s, 10)
kicker(s, "工具即能力")
title(s, "自媒体不是随便拍拍", 0.95, 38)
txt(s, ML, 1.95, CW, 0.4, "做不好自媒体，本质是你不懂这个时代的语言。", 15, False, GRAY)
for i, (hd, bd) in enumerate([
    ("怎么表达", "同一件事，\n怎么讲才有人听"),
    ("怎么呈现", "打光、镜头、\n画面质感"),
    ("拍什么", "选题、场景、\n内容的取舍"),
]):
    x = ML + i * (CW / 3 + 0.02)
    wd = CW / 3 - 0.18
    rect(s, x, 2.55, wd, 1.75, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    rect(s, x, 2.55, wd, 0.06, fill=RED)
    txt(s, x + 0.3, 2.8, wd - 0.6, 0.4, hd, 19, True, INK)
    txt(s, x + 0.3, 3.3, wd - 0.6, 0.8, bd, 13.5, False, GRAY, spacing=1.4)
rect(s, ML, 4.6, CW, 1.5, fill=INK)
txt(s, ML + 0.4, 4.78, CW - 0.8, 1.1,
    "你学的很多东西，其实都在聚焦这一个点。\n这就是主线 —— 知道自己要做什么，学习效率自然就高。",
    19, True, WHITE, spacing=1.45)

# ============ 11 认知 ============
s = prs.slides.add_slide(BLANK); page(s, 11)
kicker(s, "认知决定成败")
title(s, "同样是练马甲线，两种认知，两种结果", 0.95, 31)
for i, (tag, head, body, hi) in enumerate([
    ("低认知", "只是为了让身材好看", "动力不够。\n没必要非好看不可，\n大概率练不出来。", False),
    ("高认知", "这是我人生的资产", "每一斤、每一块腹肌都有价值。\n会坚持，甚至会急迫，\n想尽快完成。", True),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    rect(s, x, 2.2, wd, 2.85, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.035)
    rect(s, x, 2.2, wd, 0.62, fill=(LIGHT if not hi else RED))
    txt(s, x + 0.3, 2.36, wd - 0.6, 0.3, tag, 12, True, (GRAY if not hi else WHITE))
    txt(s, x + 0.3, 2.98, wd - 0.6, 0.45, head, 20, True, (INK if not hi else RED))
    txt(s, x + 0.3, 3.55, wd - 0.6, 1.3, body, 14, False, GRAY, spacing=1.45)
txt(s, ML, 5.35, CW, 0.5, "人生所有东西，都源于认知上来。", 20, True, RED)
txt(s, ML, 5.95, CW, 0.4, "认知的高度，决定你能不能坚持、会不会用心。", 14, False, GRAY)

# ============ 12 我的两个东西 ============
s = prs.slides.add_slide(BLANK); page(s, 12)
kicker(s, "现场分享")
title(s, "我人生拿得出手的，就两件事", 0.95, 34)
card(s, ML, 2.15, CW / 2 - 0.16, 2.25, "营养行业十多年", "这是走过的路，也是我能真实分享的东西。",
     head_size=22, body_size=14.5, badge="01")
card(s, ML + CW / 2 + 0.16, 2.15, CW / 2 - 0.16, 2.25, "把自媒体打通了", "健康 + 表达，两条腿走通了，就能放大。",
     head_size=22, body_size=14.5, badge="02")
rect(s, ML, 4.72, CW, 1.3, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML + 0.38, 4.88, CW - 0.8, 1.0,
    "我一直在分享自己真实走过的路、做过的事，不去外面追一堆方法技巧。\n所以不乱，能做自己，每天做的都是自己喜欢、也清楚为什么做的事。",
    15, False, INK, spacing=1.45)
txt(s, ML, 6.15, CW, 0.35, "商业定位，就是从这里面挖，再跟当下要做的业务挂上钩。", 13.5, False, GRAY)

# ============ 13 分隔 02 ============
s = prs.slides.add_slide(BLANK); divider(s, "02", "安利的业务路径", "基础工作 → 关键工作 → 2000 分 → 复制", 13)

# ============ 14 路径不清晰 ============
s = prs.slides.add_slide(BLANK); page(s, 14)
kicker(s, "先说病灶")
title(s, "路径不清晰，人就会发懵", 0.95, 38)
txt(s, ML, 1.95, CW, 0.4, "群多、会多、线下学习多 —— 不拆清楚每天该做什么，人就散了。", 14, False, GRAY)
for i, t in enumerate([
    "不知道哪个工作是重点，胡子眉毛一把抓",
    "不知道自己每天应该做什么，效率很低",
    "不知道目标在哪，做着做着就没有方向",
]):
    y = 2.5 + i * 0.72
    rect(s, ML, y, 0.42, 0.42, fill=RED, shape=MSO_SHAPE.OVAL)
    txt(s, ML, y + 0.045, 0.42, 0.3, str(i + 1), 13, True, WHITE, PP_ALIGN.CENTER)
    txt(s, ML + 0.68, y + 0.02, CW - 0.7, 0.4, t, 17, False, INK)
rect(s, ML, 4.85, CW, 1.5, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
rect(s, ML, 4.85, 0.06, 1.5, fill=RED)
txt(s, ML + 0.42, 5.0, CW - 0.9, 0.4, "《认知觉醒》里说：一切焦虑和迷茫，都来自「模糊」。", 16, True, INK)
txt(s, ML + 0.42, 5.5, CW - 0.9, 0.7, "办法就是一直往下拆，拆到最后就清晰了。目标清楚了，效率才上得来。", 14.5, False, GRAY)

# ============ 15 价值怎么衡量 ============
s = prs.slides.add_slide(BLANK); page(s, 15)
kicker(s, "先看标尺")
title(s, "安利里，价值看你培养了多少个 DD", 0.9, 31)
txt(s, ML, 1.85, CW, 0.4, "衡量标准很直接：你培养的 DD 越多，价值越大，收入越高。", 14, False, GRAY)
for i, (n, t) in enumerate([("3 组", "百万年薪经理"), ("6 组", "更高层级"), ("12 组", "更高层级"), ("20 组", "更高层级")]):
    x = ML + i * (CW / 4 + 0.02)
    wd = CW / 4 - 0.18
    hi = (i == 0)
    rect(s, x, 2.45, wd, 1.6, fill=(RED if hi else WHITE), line=(None if hi else LINE),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x, 2.7, wd, 0.6, n, 30, True, (WHITE if hi else INK), PP_ALIGN.CENTER)
    txt(s, x, 3.4, wd, 0.4, t, 12.5, False, (RGBColor(0xFF, 0xE0, 0xE4) if hi else GRAY), PP_ALIGN.CENTER)
rect(s, ML, 4.4, CW, 1.2, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML + 0.4, 4.55, CW - 0.8, 0.52, "对应的业绩口径：12.5 万", 24, True, RED)
txt(s, ML + 0.4, 5.05, CW - 0.8, 0.4, "所以关键动作很清楚 ——", 13, False, GRAY)
txt(s, ML, 5.85, CW, 0.5, "先自己做到银章，再培养银章。", 20, True, INK)

# ============ 16 银章怎么拿下 ============
s = prs.slides.add_slide(BLANK); page(s, 16)
kicker(s, "关键目标")
title(s, "银章怎么拿下？", 0.95, 40)
txt(s, ML, 2.0, CW, 0.4, "你会做银章，就能培养很多银章；培养的银章多，价值和收入就上来了。", 15, False, GRAY)
for i, (hd, bd) in enumerate([
    ("先自己做到", "从 DD 到 SP，\n再到银章。"),
    ("再培养银章", "把你做银章的\n能力复制出去。"),
    ("银章多了", "价值更大，\n收入更高。"),
]):
    x = ML + i * (CW / 3 + 0.02)
    wd = CW / 3 - 0.18
    rect(s, x, 2.6, wd, 1.85, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.3, 2.82, wd - 0.6, 0.35, f"0{i + 1}", 12, True, RED)
    txt(s, x + 0.3, 3.18, wd - 0.6, 0.45, hd, 20, True, INK)
    txt(s, x + 0.3, 3.72, wd - 0.6, 0.65, bd, 13.5, False, GRAY, spacing=1.4)
rect(s, ML, 4.85, CW, 1.3, fill=INK)
txt(s, ML + 0.4, 5.02, CW - 0.8, 0.9,
    "银章一个月要做 12 万，不容易。\n所以问题不是「怎么做 12 万」，而是「怎么拿下 2000 分」。",
    18, True, WHITE, spacing=1.45)

# ============ 17 2000 分是目标单元 ============
s = prs.slides.add_slide(BLANK); page(s, 17)
kicker(s, "拆到最小单元")
title(s, "2000 分，是 12.5 万的目标单元", 0.95, 32)
rect(s, ML, 2.35, CW, 1.25, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
txt(s, ML, 2.5, CW, 0.95, "12.5 万   ÷   2000 分   =   可复制的单元", 25, True, INK,
    PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
for i, (t, b) in enumerate([
    ("做到 2000 分", "个人能做到，不靠运气"),
    ("复制 2000 分", "教会别人也做到"),
    ("4 个就上银章", "1 + 4 = 银章"),
]):
    x = ML + i * (CW / 3 + 0.02)
    wd = CW / 3 - 0.18
    rect(s, x, 4.0, wd, 1.5, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.28, 4.22, wd - 0.56, 0.4, t, 18, True, INK)
    txt(s, x + 0.28, 4.72, wd - 0.56, 0.55, b, 13, False, GRAY, spacing=1.35)
txt(s, ML, 5.85, CW, 0.5, "做到 2000 分，无非就是复制 2000 分。", 18, True, RED)

# ============ 18 2000 分怎么做 ============
s = prs.slides.add_slide(BLANK); page(s, 18)
kicker(s, "2000 分 = 结果")
title(s, "结果要倒推成 KR：两件关键工作", 0.95, 31)
card(s, ML, 2.1, CW / 2 - 0.16, 2.5, "① 把方案讲到客户愿意试",
     "轻断食 / 糖尿病 / 早餐 / 护肝 / 腰突……\n讲到对方听得懂、愿意尝试。\n本质就是销售 —— 学会讲项目。",
     head_size=19, body_size=14, badge="KR 1 · 讲方案")
card(s, ML + CW / 2 + 0.16, 2.1, CW / 2 - 0.16, 2.5, "② 找到客户",
     "一单 200–300 分，\n10–20 个客户就够 2000 分。\n这一步，个人完全做得到。",
     head_size=19, body_size=14, badge="KR 2 · 找客户")
rect(s, ML, 4.85, CW, 1.3, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML + 0.4, 5.02, CW - 0.8, 0.4, "但前提是基本功要到位。", 17, True, RED)
txt(s, ML + 0.4, 5.5, CW - 0.8, 0.55, "很多人课学了不少，学「会了」却未必做得了 KR —— 能力没到，动作就变形。", 14, False, INK)

# ============ 19 全景图 ============
s = prs.slides.add_slide(BLANK); page(s, 19)
kicker(s, "一张图看懂")
title(s, "安利业务路径全景", 0.92, 34)
stages = [
    ("01", "基础工作", "社群 · 朋友圈\n短视频", False),
    ("02", "关键工作", "讲产品\n讲项目·讲方案", False),
    ("03", "2000 分", "个人基本\n工作单元", True),
    ("04", "复制 2000 分", "找人 · OPP\n辅导他做到", False),
    ("05", "银章", "1 + 4 × 2000 分", False),
]
bw = 1.92
gap = (CW - bw * 5) / 4
for i, (n, hd, bd, hi) in enumerate(stages):
    x = ML + i * (bw + gap)
    rect(s, x, 2.35, bw, 2.35, fill=(RED if hi else WHITE),
         line=(None if hi else LINE), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.055)
    txt(s, x + 0.2, 2.55, bw - 0.4, 0.24, n, 10, True, (RGBColor(0xFF, 0xC9, 0xC9) if hi else RED))
    txt(s, x + 0.2, 2.92, bw - 0.4, 0.9, hd, 17, True, (WHITE if hi else INK), spacing=1.2)
    txt(s, x + 0.2, 3.86, bw - 0.4, 0.7, bd, 12, False, (RGBColor(0xFF, 0xE0, 0xE4) if hi else GRAY), spacing=1.35)
    if i < 4:
        txt(s, x + bw, 2.35, gap, 2.35, "›", 22, True, RGBColor(0xC9, 0xC4, 0xBE),
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
txt(s, ML, 5.05, CW, 0.5, "顺序不能乱：先做顺基础，再做关键；没有基础，关键工作没有开口。", 16, True, RED)
txt(s, ML, 5.6, CW, 0.4, "再往上：复制银章 → 冲钻石经理。路径是同一套，只是规模变大。", 14, False, GRAY)

# ============ 20 基础工作 ============
s = prs.slides.add_slide(BLANK); page(s, 20)
kicker(s, "基础工作")
title(s, "关键工作之前，先有人可讲", 0.95, 34)
txt(s, ML, 1.95, CW, 0.4, "基础工作不直接出业绩，但没有它，关键工作根本不会发生。", 14, False, GRAY)
for i, (hd, bd) in enumerate([
    ("朋友圈", "不发，谁看得到你？\n你又不是巷子里那口美酒，\n香味不会自己飘出去。"),
    ("社群运营", "本质也是分享。\n让人持续看见你、\n在群里被记住。"),
    ("短视频", "同一个逻辑，\n持续放大，\n让更多陌生人先认识你。"),
]):
    x = ML + i * (CW / 3 + 0.02)
    wd = CW / 3 - 0.18
    rect(s, x, 2.5, wd, 2.45, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.045)
    rect(s, x, 2.5, wd, 0.06, fill=RED)
    txt(s, x + 0.3, 2.76, wd - 0.6, 0.4, hd, 21, True, INK)
    txt(s, x + 0.3, 3.3, wd - 0.6, 1.35, bd, 13.5, False, GRAY, spacing=1.45)
txt(s, ML, 5.25, CW, 0.4, "把基础做足，人才找得到你；人找得到你，才有机会做关键工作。", 17, True, RED)
txt(s, ML, 5.85, CW, 0.4, "关键工作量足够多，2000 分就拿下了。", 14, False, GRAY)

# ============ 21 医生类比 ============
s = prs.slides.add_slide(BLANK); page(s, 21)
kicker(s, "一个类比")
title(s, "在医院，我也是这么干活的", 0.95, 36)
for i, (hd, bd, hi) in enumerate([
    ("术前会诊 · 看病人", "没有直接收益，\n但一步都不能少。", False),
    ("真正开刀", "产生收益的是这一下 ——\n这叫关键工作。", True),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    rect(s, x, 2.2, wd, 1.9, fill=(RED if hi else WHITE), line=(None if hi else LINE),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    txt(s, x + 0.32, 2.48, wd - 0.64, 0.45, hd, 21, True, (WHITE if hi else INK))
    txt(s, x + 0.32, 3.08, wd - 0.64, 0.8, bd, 14, False,
        (RGBColor(0xFF, 0xE0, 0xE4) if hi else GRAY), spacing=1.45)
txt(s, ML, 4.35, CW, 0.5, "想拿开刀那一下的结果，前面的会诊和看病人就必须做。", 16, True, INK)
rect(s, ML, 5.0, CW, 1.25, fill=INK)
txt(s, ML + 0.4, 5.18, CW - 0.8, 0.9,
    "安利一模一样：朋友圈、社群、分享课程都不直接出业绩，\n但没有它们，连人都找不到。",
    18, True, WHITE, spacing=1.45)

# ============ 22 认知不到 ============
s = prs.slides.add_slide(BLANK); page(s, 22)
kicker(s, "行为背后是认知")
title(s, "认知不到，行为就上不来", 0.95, 34)
for i, t in enumerate([
    "朋友圈随便发，应付一下就算完成",
    "群里不回应、不贡献内容",
    "本来有人在那里，你还不去发",
]):
    y = 2.15 + i * 0.72
    rect(s, ML, y + 0.08, 0.16, 0.16, fill=RED)
    txt(s, ML + 0.42, y, CW - 0.5, 0.4, t, 17, False, INK)
txt(s, ML, 4.35, CW, 0.4, "不是不会，是认知不到 —— 不认为它重要，自然不会认真。", 16, True, RED)
rect(s, ML, 4.95, CW, 1.3, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML + 0.4, 5.12, CW - 0.8, 0.95,
    "我这几年的短视频和朋友圈，都是认认真真在发。\n不是为了百日筑基凑三个打卡 —— 认真，才会影响到人，才容易链接到人。",
    14.5, False, INK, spacing=1.45)

# ============ 23 复制 2000 分 ============
s = prs.slides.add_slide(BLANK); page(s, 23)
kicker(s, "复制")
title(s, "复制 2000 分，是一个闭环", 0.95, 34)
flow = [
    ("①", "链接", "基础工作把人\n引过来"),
    ("②", "讲 OPP", "讲清事业价值，\n他才想跟你做"),
    ("③", "跟随系统", "想做 → 决定做，\n系统把人留下来"),
    ("④", "辅导", "把走过的能力\n全部教会他"),
]
for i, (n, hd, bd) in enumerate(flow):
    x = ML + i * (CW / 4 + 0.03)
    wd = CW / 4 - 0.18
    rect(s, x, 2.2, wd, 1.95, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.24, 2.4, wd - 0.48, 0.28, n, 12, True, RED)
    txt(s, x + 0.24, 2.7, wd - 0.48, 0.4, hd, 17, True, INK)
    txt(s, x + 0.24, 3.18, wd - 0.48, 0.85, bd, 12, False, GRAY, spacing=1.35)
    if i < 3:
        txt(s, x + wd, 2.2, 0.25, 1.95, "›", 19, True, RGBColor(0xC9, 0xC4, 0xBE),
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
txt(s, ML, 4.4, CW, 0.4, "辅导的内容，就是你走过的全部能力：", 15, True, RED)
for i, t in enumerate(["轻断食怎么讲", "项目方案怎么讲", "基础工作怎么分享", "百日筑基怎么带"]):
    x = ML + i * (CW / 4 + 0.01)
    rect(s, x, 4.95, CW / 4 - 0.16, 0.75, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    txt(s, x, 5.08, CW / 4 - 0.16, 0.5, t, 12.5, False, INK, PP_ALIGN.CENTER)
txt(s, ML, 5.98, CW, 0.4, "你带着他走一遍，他就有了一个 2000 分。", 14, False, GRAY)

# ============ 24 可复制单元 ============
s = prs.slides.add_slide(BLANK); page(s, 24)
kicker(s, "算术")
title(s, "2000 分 × N，就是层级", 0.95, 38)
txt(s, ML + 0.05, 2.0, 6.0, 1.6, "1 → 2 → 4", 68, True, RED)
txt(s, ML + 0.05, 3.75, 6.2, 0.8, "教一个、两个、四个\n就是 SP，再往上就是银章。", 14, False, GRAY, spacing=1.4)
for i, t in enumerate(["自己先做到 2000 分", "教会 1 个人做到", "教会 2 个人做到", "教会 4 个人做到 = 银章"]):
    y = 2.15 + i * 0.9
    rect(s, 7.3, y, 5.05, 0.72, fill=(RED if i == 3 else WHITE), line=(None if i == 3 else LINE),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    txt(s, 7.55, y + 0.2, 4.6, 0.35, t, 14.5, True, (WHITE if i == 3 else INK))
txt(s, 7.3, 5.78, 5.05, 0.4, "上了银章再继续复制 → 钻石经理不难", 12.5, False, RED)
rect(s, ML, 5.55, 6.3, 1.0, fill=INK)
txt(s, ML + 0.35, 5.72, 5.7, 0.7, "认知清晰了，才会聚焦。", 18, True, WHITE)

# ============ 25 朋友圈是 IP ============
s = prs.slides.add_slide(BLANK); page(s, 25)
kicker(s, "重点纠正")
title(s, "朋友圈不是广告，是你的 IP", 0.95, 34)
for i, (tag, head, body, hi) in enumerate([
    ("不叫朋友圈", "每日周报 · 健康早报", "那是广告。\n没有人喜欢看广告。", False),
    ("这才叫朋友圈", "五星级高质量内容", "你的思考、你的价值观、\n你的价值主张。", True),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    rect(s, x, 2.15, wd, 2.2, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    rect(s, x, 2.15, wd, 0.56, fill=(LIGHT if not hi else RED))
    txt(s, x + 0.3, 2.28, wd - 0.6, 0.3, tag, 12, True, (GRAY if not hi else WHITE))
    txt(s, x + 0.3, 2.85, wd - 0.6, 0.42, head, 19, True, (INK if not hi else RED))
    txt(s, x + 0.3, 3.4, wd - 0.6, 0.75, body, 13.5, False, GRAY, spacing=1.4)
rect(s, ML, 4.6, CW, 1.5, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML + 0.4, 4.75, CW - 0.8, 0.4, "现在有 AI，这件事好办很多：", 15, True, RED)
txt(s, ML + 0.4, 5.2, CW - 0.8, 0.8,
    "把我写好的朋友圈规则喂给豆包，说一句「今天参加超级个体俱乐部的感受」——\n它就能按你的风格、你的味道，写出一条属于你的朋友圈。",
    14, False, INK, spacing=1.45)

# ============ 26 先看见再谈 ============
s = prs.slides.add_slide(BLANK); page(s, 26)
kicker(s, "成交的真相")
title(s, "别人先看见你，才会问你的生意", 0.95, 31)
txt(s, ML, 1.9, CW, 0.4, "他没有先看见你 —— 沟通当然没效果。", 14, False, GRAY)
flow = [
    ("先看见", "你的成长、思考力、\n状态和气质的改变"),
    ("主动来问", "「你这个生意\n是怎么做的？」"),
    ("再讲", "这时候讲，\n成功率 80%–90%"),
]
for i, (hd, bd) in enumerate(flow):
    x = ML + i * (CW / 3 + 0.05)
    wd = CW / 3 - 0.2
    hi = (i == 2)
    rect(s, x, 2.4, wd, 1.85, fill=(RED if hi else WHITE), line=(None if hi else LINE),
         shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.3, 2.62, wd - 0.6, 0.45, hd, 19, True, (WHITE if hi else INK))
    txt(s, x + 0.3, 3.2, wd - 0.6, 0.85, bd, 13, False,
        (RGBColor(0xFF, 0xE0, 0xE4) if hi else GRAY), spacing=1.4)
    if i < 2:
        txt(s, x + wd, 2.4, 0.25, 1.85, "›", 19, True, RGBColor(0xC9, 0xC4, 0xBE),
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, ML, 4.65, CW, 1.35, fill=INK)
txt(s, ML + 0.4, 4.82, CW - 0.8, 1.0,
    "别人心里，你还是以前那个你 —— 他就不认为你有资格跟他谈生意。\n所以效果不靠话术，靠背后这些基础工作。",
    17, True, WHITE, spacing=1.45)

# ============ 27 分隔 03 ============
s = prs.slides.add_slide(BLANK); divider(s, "03", "认知与收口", "认知不到位，再重要的事也做不出效果", 27)

# ============ 28 想做 ≠ 决定做 ============
s = prs.slides.add_slide(BLANK); page(s, 28)
kicker(s, "留人")
title(s, "想做，不等于决定做", 0.95, 38)
for i, (tag, head, body, hi) in enumerate([
    ("想做", "一时有想法", "有想法的人很多，\n但想做的未必会留下。", False),
    ("决定做", "靠系统留下来", "跟随系统、团队接住，\n人才真正留下来。", True),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    rect(s, x, 2.15, wd, 2.3, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    rect(s, x, 2.15, wd, 0.58, fill=(LIGHT if not hi else RED))
    txt(s, x + 0.3, 2.29, wd - 0.6, 0.3, tag, 12, True, (GRAY if not hi else WHITE))
    txt(s, x + 0.3, 2.87, wd - 0.6, 0.45, head, 21, True, (INK if not hi else RED))
    txt(s, x + 0.3, 3.45, wd - 0.6, 0.9, body, 13.5, False, GRAY, spacing=1.4)
txt(s, ML, 4.75, CW, 0.4, "留下来之后，再靠辅导把能力复制出去：", 15, True, RED)
for i, t in enumerate(["人能做了", "2000 分坐得稳", "市场才不脆弱"]):
    x = ML + i * (CW / 3 + 0.02)
    rect(s, x, 5.3, CW / 3 - 0.18, 0.85, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    txt(s, x, 5.47, CW / 3 - 0.18, 0.5, t, 15, True, INK, PP_ALIGN.CENTER)

# ============ 29 今天讲的是心法 ============
s = prs.slides.add_slide(BLANK); page(s, 29)
kicker(s, "整场收口")
title(s, "今天讲的，其实是心法", 0.95, 40)
txt(s, ML, 2.0, CW, 0.45, "成功是简单的 —— 一直是这样。", 17, False, GRAY)
chain = ["把业务路径理清楚", "认知跟着上来", "真正重视、用心做", "做到极致"]
for i, t in enumerate(chain):
    x = ML + i * (CW / 4 + 0.02)
    wd = CW / 4 - 0.2
    rect(s, x, 2.75, wd, 1.35, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    txt(s, x + 0.15, 2.95, wd - 0.3, 0.95, t, 15, True, INK, PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE, spacing=1.3)
    if i < 3:
        txt(s, x + wd, 2.75, 0.2, 1.35, "›", 20, True, RGBColor(0xC9, 0xC4, 0xBE),
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
rect(s, ML, 4.55, CW, 1.6, fill=INK)
txt(s, ML + 0.45, 4.72, CW - 0.9, 1.3,
    "认知不到，你就不会认为这件事很重要，也就不会认真去做。\n认知足够高，你才看得见它的价值，才会用心做到极致。",
    19, True, WHITE, spacing=1.5)

# ============ 30 结尾 ============
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=INK)
rect(s, 0, 0, 0.42, H, fill=RED)
txt(s, ML + 0.35, 1.9, 10.8, 2.6,
    "认知不到位，\n再重要的事，也做不出效果。", 40, True, WHITE, spacing=1.4)
rect(s, ML + 0.4, 4.6, 1.6, 0.05, fill=RED)
txt(s, ML + 0.35, 4.95, 10.4, 0.9,
    "先把路径捋清楚，再把认知拔上去。\n剩下的，就是认认真真去做。", 15, False, RGBColor(0xB0, 0xB0, 0xB0), spacing=1.6)
txt(s, W - MR - 2.4, H - 0.75, 2.4, 0.3, "张医生 · AI 营养师", 11, False,
    RGBColor(0x8A, 0x8A, 0x8A), PP_ALIGN.RIGHT)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "张医生-事业指导-20261108-成事之道与业务路径.pptx")
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
