#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""张医生-事业指导-20261105-成功法则与业务路径.pptx 生成脚本
版式：16:9 · 米白底 · 砖红主色 · 无边框阴影 · 三档字号
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
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

prs = Presentation()
prs.slide_width = Inches(W)
prs.slide_height = Inches(H)
BLANK = prs.slide_layouts[6]

TOTAL = 23


def _ea(run):
    """设置中文字体（东亚字体）"""
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', FONT)


def txt(slide, x, y, w, h, text, size=18, bold=False, color=INK,
        align=PP_ALIGN.LEFT, spacing=1.25, anchor=MSO_ANCHOR.TOP, margin=0):
    """单段落文本框"""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(margin)
    tf.margin_top = tf.margin_bottom = 0
    lines = text.split("\n")
    for i, ln in enumerate(lines):
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


def page(slide, num, footer="张医生 · AI营养师｜一对一辅导整理 · 2026-11-05"):
    rect(slide, 0, 0, W, H, fill=BG)
    rect(slide, ML, H - 0.62, CW, 0.012, fill=LINE)
    txt(slide, ML, H - 0.55, 7.5, 0.3, footer, 10, False, GRAY)
    txt(slide, W - MR - 1.2, H - 0.55, 1.2, 0.3, f"{num:02d} / {TOTAL}", 10,
        False, GRAY, PP_ALIGN.RIGHT)


def kicker(slide, text, y=0.62):
    rect(slide, ML, y + 0.06, 0.16, 0.16, fill=RED)
    txt(slide, ML + 0.32, y, 6.0, 0.3, text, 12, True, RED)


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
txt(s, ML + 0.3, 1.62, 8.0, 0.35, "张医生 · AI 营养师｜2026-11-05 一对一辅导", 13, True, RED)
txt(s, ML + 0.3, 2.15, 10.2, 2.2, "成功法则\n与安利业务路径", 54, True, INK, spacing=1.28)
rect(s, ML + 0.35, 4.72, 2.0, 0.05, fill=RED)
txt(s, ML + 0.3, 5.02, 9.6, 0.9,
    "把做到过的事做成产品，聚焦一个点做到极致\n安利的工作只有一条路径：2000 分 → 复制 2000 分",
    16, False, GRAY, spacing=1.55)
txt(s, W - MR - 2.6, H - 0.85, 2.6, 0.3, "整理稿 · 配套课件", 11, False, GRAY, PP_ALIGN.RIGHT)

# ============ 02 今天只讲两件事 ============
s = prs.slides.add_slide(BLANK); page(s, 2)
kicker(s, "本场结构")
title(s, "41 分钟，只讲两个板块")
card(s, ML, 2.25, CW / 2 - 0.18, 3.2, "一 · 个人定位\n成功的法则",
     "产品思维：把你真正做到过的事，变成拿得出手的产品。\n\n认知高度：这件事为什么值得做，决定了你能不能做成。",
     head_size=21, body_size=14.5, badge="PART 01")
card(s, ML + CW / 2 + 0.18, 2.25, CW / 2 - 0.18, 3.2, "二 · 业务路径\n2000 分与复制",
     "基础工作铺路 → 关键工作收单 → 2000 分落地。\n\nOPP 找人，团队接住，复制 4 个 2000 分就是银章。",
     head_size=21, body_size=14.5, badge="PART 02")
txt(s, ML, 5.78, CW, 0.5, "把这两块吃透，上银章、做到经理，不难。", 17, True, RED)

# ============ 03 分隔 01 ============
s = prs.slides.add_slide(BLANK); divider(s, "01", "成功的法则", "成功是简单的——前提是你有产品思维", 3)

# ============ 04 成功是简单的 ============
s = prs.slides.add_slide(BLANK); page(s, 4)
kicker(s, "核心命题")
title(s, "成功是简单的，\n你一定要相信这个", 1.15, 40)
txt(s, ML, 3.55, CW, 0.6, "简单在哪——", 15, True, GRAY)
rect(s, ML, 4.05, CW, 1.75, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
rect(s, ML, 4.27, 0.05, 1.3, fill=RED)
txt(s, ML + 0.38, 4.28, CW - 0.8, 1.3,
    "把你真正做过的事情、走过的路、你的想法，\n想办法把它变成产品。",
    26, True, INK, spacing=1.45)

# ============ 05 产品思维三个例子 ============
s = prs.slides.add_slide(BLANK); page(s, 5)
kicker(s, "产品思维")
title(s, "你做成过的事，就是你的产品", 0.98, 34)
items = [
    ("你做成过的事", "你把家里两个小孩\n的身体调好了", "儿童追高方案"),
    ("你做成过的事", "你把自己\n练出了腹肌", "体重管理 · 腹肌方案"),
    ("你做成过的事", "你从医院跳出来创业\n从小白做到营养专家", "OPP · 对创业的理解"),
]
cw3 = (CW - 0.44) / 3
for i, (h0, h1, prod) in enumerate(items):
    x = ML + i * (cw3 + 0.22)
    rect(s, x, 2.15, cw3, 3.35, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.045)
    txt(s, x + 0.3, 2.42, cw3 - 0.6, 0.26, h0, 10.5, True, GRAY)
    txt(s, x + 0.3, 2.76, cw3 - 0.6, 1.35, h1, 17, True, INK, spacing=1.35)
    txt(s, x, 4.24, cw3, 0.3, "↓", 15, True, RED, PP_ALIGN.CENTER)
    txt(s, x + 0.3, 4.62, cw3 - 0.6, 0.6, prod, 16, True, RED)
txt(s, ML, 5.85, CW, 0.4, "这些产品不是编出来的，是你本身拥有、拿得出手的东西。", 15, False, GRAY)

# ============ 06 两步：打磨 + 分享 ============
s = prs.slides.add_slide(BLANK); page(s, 6)
kicker(s, "动作只有两步")
title(s, "打磨自己 → 学会分享", 0.98, 38)
card(s, ML, 2.2, CW / 2 - 0.16, 2.0, "第一步 · 打磨",
     "聚焦核心的那一个点——你最有感觉、\n做得最极致的那个。不是摊开一堆。",
     head_size=20, body_size=14.5, badge="PRODUCT")
card(s, ML + CW / 2 + 0.16, 2.2, CW / 2 - 0.16, 2.0, "第二步 · 分享",
     "自媒体、朋友圈、社群输出——凡是能分享的\n地方都放出去，让更多人看见，找到客户。",
     head_size=20, body_size=14.5, badge="CHANNEL")
rect(s, ML, 4.62, CW, 1.15, fill=INK)
txt(s, ML + 0.4, 4.86, CW - 0.8, 0.7,
    "不做到极致，你和普通人就没有区别。台上那些人，全是做到极致的。",
    19, True, WHITE, spacing=1.3)

# ============ 07 认知的高度 ============
s = prs.slides.add_slide(BLANK); page(s, 7)
kicker(s, "认知决定成败")
title(s, "同样是减 40 斤，两种认知，两种结果", 0.95, 32)
for i, (tag, head, body, col) in enumerate([
    ("低认知", "目标是外界给的", "领导要你减、指标要你减。\n靠意志力硬扛，做得很痛苦，\n很难坚持。", GRAY),
    ("高认知", "这是我人生的资产", "减下来我就有了拿得出手的东西，\n可以借这个点一劳永逸。\n由内而外，做得快乐。", RED),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    rect(s, x, 2.2, wd, 3.0, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.035)
    rect(s, x, 2.2, wd, 0.62, fill=(LIGHT if i == 0 else RED))
    txt(s, x + 0.3, 2.36, wd - 0.6, 0.3, tag, 12, True, (GRAY if i == 0 else WHITE))
    txt(s, x + 0.3, 3.0, wd - 0.6, 0.4, head, 21, True, (INK if i == 0 else RED))
    txt(s, x + 0.3, 3.58, wd - 0.6, 1.4, body, 14, False, GRAY, spacing=1.45)
    if i == 0:
        txt(s, x + 0.3, 4.55, wd - 0.6, 0.4, "→ 压力型目标", 13, True, GRAY)
    else:
        txt(s, x + 0.3, 4.55, wd - 0.6, 0.4, "→ 梦想型目标", 13, True, RED)
txt(s, ML, 5.55, CW, 0.6, "练腹肌也一样：认知停在「六块腹肌好看」，不足以让你做成这件事。", 15, False, GRAY)

# ============ 08 金句 ============
s = prs.slides.add_slide(BLANK); page(s, 8)
kicker(s, "一句话")
quote(s, "人生成功一次，\n就可以成功一辈子。", "认知的高度上去，你就不再把这件事当压力，而是当成有价值的事。", 34)

# ============ 09 我的 IP ============
s = prs.slides.add_slide(BLANK); page(s, 9)
kicker(s, "现场问答")
title(s, "我的 IP，来自两件拿得出手的事", 0.95, 32)
card(s, ML, 2.15, CW / 2 - 0.16, 2.35, "从医院跳出来，去创业", "并且成功了。这段过程本身就是内容，分享出去就能吸引人。",
     head_size=21, body_size=14.5, badge="01")
card(s, ML + CW / 2 + 0.16, 2.15, CW / 2 - 0.16, 2.35, "从小白到营养专家", "一路走过来的专业积累，让客户认你，IP 自然建立。",
     head_size=21, body_size=14.5, badge="02")
txt(s, ML, 4.85, CW, 0.5, "IP 不是设计出来的，是你真的走过、并且真的做成了。", 17, True, RED)
txt(s, ML, 5.45, CW, 0.5, "那你的呢？你的东西是什么？—— 这是每个人要先回答的问题。", 15, False, GRAY)

# ============ 10 凭什么交付你 ============
s = prs.slides.add_slide(BLANK); page(s, 10)
kicker(s, "扎心一问")
quote(s, "人家凭什么交付你？\n你都没有，\n凭什么交付你？",
      "所以别浮躁。沉下心，找到一个点去击破它。", 30)

# ============ 11 分隔 02 ============
s = prs.slides.add_slide(BLANK); divider(s, "02", "安利的业务路径", "2000 分 → 复制 2000 分", 11)

# ============ 12 路径不清晰 ============
s = prs.slides.add_slide(BLANK); page(s, 12)
kicker(s, "先说病灶")
title(s, "业务路径不清晰，每天就抓不住重点", 0.95, 32)
for i, t in enumerate([
    "不知道哪个工作是重点，胡子眉毛一把抓",
    "掌握不了这个事业的节奏",
    "不知道怎么就做到了，更不知道怎么做大",
]):
    y = 2.35 + i * 0.78
    rect(s, ML, y, 0.42, 0.42, fill=RED, shape=MSO_SHAPE.OVAL)
    txt(s, ML, y + 0.045, 0.42, 0.3, str(i + 1), 13, True, WHITE, PP_ALIGN.CENTER)
    txt(s, ML + 0.68, y + 0.02, CW - 0.7, 0.4, t, 18, False, INK)
rect(s, ML, 4.95, CW, 1.3, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML + 0.38, 5.12, CW - 0.8, 0.95,
    "就像你在学校做水电工：什么时候关闸、什么时候梳理、什么时候做宣教，\n流程清清楚楚，工作才做得好。回到安利，你清晰吗？",
    15, False, INK, spacing=1.45)

# ============ 13 全景图 ============
s = prs.slides.add_slide(BLANK); page(s, 13)
kicker(s, "一张图看懂")
title(s, "安利业务路径全景", 0.92, 34)
stages = [
    ("01", "基础工作", "社群 · 朋友圈\n短视频", False),
    ("02", "关键工作", "讲产品\n讲项目·讲方案", False),
    ("03", "2000 分", "个人基本\n工作单元", True),
    ("04", "复制 2000 分", "OPP 找人\n辅导他做到", False),
    ("05", "银章", "4 × 2000 分", False),
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
txt(s, ML, 5.6, CW, 0.5, "基础工作不立刻出业绩，但它决定了关键工作有没有机会发生。", 14, False, GRAY)

# ============ 14 2000 分 ============
s = prs.slides.add_slide(BLANK); page(s, 14)
kicker(s, "核心工作单元")
title(s, "2000 分：你自己的基本工作量", 0.95, 32)
txt(s, ML, 2.1, CW, 0.4, "注意范围：这 2000 分是你个人的，不算团队、不算龙卷风。", 14, False, GRAY)
rect(s, ML, 2.6, CW, 1.05, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
txt(s, ML, 2.72, CW, 0.8, "10 – 20 个顾客   →   2000 分", 26, True, INK, PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
txt(s, ML, 3.95, CW, 0.4, "顾客从哪来 —— 从你对产品、对项目有感觉来。", 15, True, RED)
cards = [
    ("对产品有感觉", "你说的不是洗洁精，是你主打的那个产品。"),
    ("对项目有感觉", "对轻断食有感觉，才会去找要减肥的人。"),
    ("对人群有感觉", "对糖尿病方案有感觉，才会去找血糖高的人。"),
]
for i, (hd, bd) in enumerate(cards):
    x = ML + i * (CW / 3 + 0.02)
    wd = CW / 3 - 0.18
    rect(s, x, 4.5, wd, 1.35, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.26, 4.68, wd - 0.5, 0.32, hd, 15, True, INK)
    txt(s, x + 0.26, 5.05, wd - 0.5, 0.7, bd, 12.5, False, GRAY, spacing=1.35)
txt(s, ML, 6.15, CW, 0.4, "你不可能找到你自己都没感觉的客户。", 14, False, GRAY)

# ============ 15 关键工作 ============
s = prs.slides.add_slide(BLANK); page(s, 15)
kicker(s, "KR · 关键工作")
title(s, "坐下来，把产品、项目、方案讲清楚", 0.95, 31)
txt(s, ML, 2.0, CW, 0.4, "不是 KOC，不是「让他看到好处」就完了。是讲。", 14, False, GRAY)
for i, (nm, scene, act) in enumerate([
    ("肥胖的人", "轻断食为什么能帮你减肥", "讲方案"),
    ("糖尿病的人", "我怎么帮你把糖降下来", "讲方案"),
    ("买产品的人", "做产品示范，讲清怎么用", "讲产品"),
]):
    y = 2.5 + i * 0.98
    rect(s, ML, y, CW, 0.82, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    txt(s, ML + 0.3, y + 0.24, 1.8, 0.35, nm, 15, True, INK)
    txt(s, ML + 2.2, y + 0.24, 6.6, 0.35, scene, 14, False, GRAY)
    txt(s, W - MR - 1.5, y + 0.24, 1.2, 0.35, act, 13, True, RED, PP_ALIGN.RIGHT)
rect(s, ML, 5.62, CW, 0.95, fill=INK)
txt(s, ML + 0.4, 5.78, CW - 0.8, 0.65, "讲通了，对方才会办卡、才会做方案 —— 2000 分才拿得下来。",
    17, True, WHITE)

# ============ 16 基础工作 ============
s = prs.slides.add_slide(BLANK); page(s, 16)
kicker(s, "基础工作")
title(s, "关键工作之前，先搭桥", 0.95, 32)
txt(s, ML, 1.95, CW, 0.4, "你逮住一个肚子大的人直接讲轻断食 —— 这叫没效率，因为他没有理由听你讲。", 14, False, GRAY)
for i, (hd, bd) in enumerate([
    ("社群", "把人拉进来养：输出内容、做群氛围，养出兴趣，关键工作才讲得通。"),
    ("朋友圈", "建立 IP 与专业感：他认定你专业，才愿意听你讲。"),
    ("短视频 3.0", "同一条逻辑，持续放大，让更多人先认识你。"),
]):
    x = ML + i * (CW / 3 + 0.02)
    wd = CW / 3 - 0.18
    rect(s, x, 2.55, wd, 2.55, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.045)
    rect(s, x, 2.55, wd, 0.06, fill=RED)
    txt(s, x + 0.3, 2.82, wd - 0.6, 0.4, hd, 21, True, INK)
    txt(s, x + 0.3, 3.42, wd - 0.6, 1.4, bd, 14, False, GRAY, spacing=1.45)
txt(s, ML, 5.35, CW, 0.5, "你老是一个水电工的形象，他认定你是水电工 —— 你跟他讲营养，有用吗？",
    16, True, INK)
txt(s, ML, 5.92, CW, 0.4, "基础工作不能马上带来业绩，但它是让关键工作有机会发生的工作。", 14, False, GRAY)

# ============ 17 复制 2000 分 ============
s = prs.slides.add_slide(BLANK); page(s, 17)
kicker(s, "复制")
title(s, "复制 2000 分 = 找到人 + 讲 OPP", 0.95, 32)
flow = [("①", "讲 OPP", "找到下一个\n能做 2000 分的人"),
        ("②", "跟随团队", "社群推动、\n团队复盘会接住"),
        ("③", "辅导他", "把你怎么做 2000 分\n一步步教给他")]
for i, (n, hd, bd) in enumerate(flow):
    x = ML + i * (CW / 3 + 0.03)
    wd = CW / 3 - 0.2
    rect(s, x, 2.15, wd, 1.95, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.28, 2.34, wd - 0.56, 0.3, n, 12, True, RED)
    txt(s, x + 0.28, 2.66, wd - 0.56, 0.4, hd, 19, True, INK)
    txt(s, x + 0.28, 3.14, wd - 0.56, 0.8, bd, 13, False, GRAY, spacing=1.35)
    if i < 2:
        txt(s, x + wd, 2.15, 0.2, 1.95, "›", 20, True, RGBColor(0xC9, 0xC4, 0xBE),
            PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
txt(s, ML, 4.35, CW, 0.4, "辅导的内容，就是你当初走过的全部能力：", 15, True, RED)
for i, t in enumerate(["轻断食怎么讲", "糖尿病方案怎么讲", "社群怎么一起运作", "朋友圈怎么发"]):
    x = ML + i * (CW / 4 + 0.01)
    rect(s, x, 4.95, CW / 4 - 0.16, 0.72, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
    txt(s, x, 5.06, CW / 4 - 0.16, 0.5, t, 13, False, INK, PP_ALIGN.CENTER)
txt(s, ML, 5.95, CW, 0.4, "一个人做 2000 分势单力薄；把能力复制出去，才做得大。", 14, False, GRAY)

# ============ 18 4 × 2000 ============
s = prs.slides.add_slide(BLANK); page(s, 18)
kicker(s, "算术")
title(s, "4 个 2000 分 = 银章", 0.95, 38)
txt(s, ML + 0.05, 1.95, 6.0, 1.75, "1 + 4", 76, True, RED)
txt(s, ML + 0.05, 3.78, 6.0, 0.55, "你自己的 2000 分 ＋\n复制出来的 4 个 2000 分", 14, False, GRAY, spacing=1.35)
for i, t in enumerate(["基本功赋能会", "线下精英班", "团队复盘会"]):
    y = 2.25 + i * 0.9
    rect(s, 7.3, y, 5.05, 0.72, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    txt(s, 7.55, y + 0.2, 4.6, 0.35, t, 15, True, INK)
txt(s, 7.3, 4.95, 5.05, 0.4, "↑ 本质都是在复制 2000 分的能力", 12, False, RED)
rect(s, ML, 5.55, CW, 1.0, fill=INK)
txt(s, ML + 0.4, 5.72, CW - 0.8, 0.7, "项目都讲不好，怎么出单？把这些学习的点看明白，上银章、做到经理，不难。",
    17, True, WHITE)

# ============ 19 分隔 03 ============
s = prs.slides.add_slide(BLANK); divider(s, "03", "认知与行动", "迷茫的根是模糊，破局靠拆解", 19)

# ============ 20 模糊 ============
s = prs.slides.add_slide(BLANK); page(s, 20)
kicker(s, "心法")
title(s, "一个人迷茫焦虑，原因就是「模糊」", 0.95, 32)
for i, (tag, head, body, hi) in enumerate([
    ("模糊", "「哎呀，又没成交」", "停留在情绪，问题还在原地，\n下一次还是同样的结果。", False),
    ("拆解", "「他为什么不买？\n我哪句没讲好？」", "找到具体原因，就知道改进什么，\n问题被解决掉。", True),
]):
    x = ML + i * (CW / 2 + 0.1)
    wd = CW / 2 - 0.1
    rect(s, x, 2.15, wd, 2.5, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    rect(s, x, 2.15, wd, 0.6, fill=(LIGHT if not hi else RED))
    txt(s, x + 0.3, 2.3, wd - 0.6, 0.3, tag, 12, True, (GRAY if not hi else WHITE))
    txt(s, x + 0.3, 2.92, wd - 0.6, 0.9, head, 20, True, (INK if not hi else RED), spacing=1.3)
    txt(s, x + 0.3, 3.85, wd - 0.6, 0.7, body, 13.5, False, GRAY, spacing=1.4)
rect(s, ML, 4.95, CW, 1.35, fill=INK)
txt(s, ML + 0.4, 5.1, CW - 0.8, 1.0,
    "创业者就是解决问题。\n遇到问题记下来，解决它；自己想不通，向老师请教。",
    19, True, WHITE, spacing=1.45)

# ============ 21 认知从哪来 ============
s = prs.slides.add_slide(BLANK); page(s, 21)
kicker(s, "最后落点")
title(s, "要么看书，要么请教老师", 0.95, 36)
txt(s, ML, 1.95, CW, 0.45, "因为但凡认知，都是因为别人的思想打开了你的闭环。老师，就是书。",
    16, False, GRAY)
for i, (hd, bd) in enumerate([
    ("读书", "把别人的思想搬进来，\n打开自己的闭环。"),
    ("请教", "不喜欢看书就打电话，\n问到很细的工作层面。"),
    ("信息差", "你掌握不了成功的信息，\n老师掌握了。"),
]):
    x = ML + i * (CW / 3 + 0.02)
    wd = CW / 3 - 0.18
    rect(s, x, 2.65, wd, 1.85, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    txt(s, x + 0.28, 2.85, wd - 0.56, 0.4, hd, 19, True, RED)
    txt(s, x + 0.28, 3.35, wd - 0.56, 0.9, bd, 13.5, False, GRAY, spacing=1.4)
rect(s, ML, 4.9, CW, 1.35, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
txt(s, ML, 5.05, CW, 0.5, "认知  →  行动  →  反馈复盘  →  再认知", 22, True, INK, PP_ALIGN.CENTER)
txt(s, ML, 5.72, CW, 0.4, "不断转这个循环，人才会越来越厉害。", 13.5, False, GRAY, PP_ALIGN.CENTER)

# ============ 22 行动清单 ============
s = prs.slides.add_slide(BLANK); page(s, 22)
kicker(s, "落地")
title(s, "接下来要做的", 0.95, 36)
acts = [
    ("做规划 · 做布局", "两个板块吃透：个人定位 ＋ 业务路径"),
    ("抢自己的 2000 分", "先把基础工作做到位，关键工作才有开口"),
    ("自媒体定一个点", "围绕轻断食（后转搏击）持续输出"),
    ("学对标，不自己琢磨", "看别人怎么表达、怎么做情绪、怎么调用信息"),
    ("打磨身体这块产品", "把块头练出来，产品配得上人设"),
    ("多打电话 · 多复盘", "不清晰就问，问到很细的工作层面"),
    ("遇到问题就拆解", "不让它模糊，不让它侵袭内心"),
]
for i, (hd, bd) in enumerate(acts):
    col, row = i // 4, i % 4
    x = ML + col * (CW / 2 + 0.06)
    y = 2.05 + row * 1.12
    wd = CW / 2 - 0.12
    rect(s, x, y, wd, 0.94, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    txt(s, x + 0.24, y + 0.14, 0.4, 0.3, f"{i + 1:02d}", 11, True, RED)
    txt(s, x + 0.76, y + 0.12, wd - 1.0, 0.32, hd, 15, True, INK)
    txt(s, x + 0.76, y + 0.5, wd - 1.0, 0.36, bd, 11.5, False, GRAY)

# ============ 23 结尾 ============
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, fill=INK)
rect(s, 0, 0, 0.42, H, fill=RED)
txt(s, ML + 0.35, 1.9, 10.5, 2.6,
    "把认知拔到那个高度，\n然后把事情做到极致。", 40, True, WHITE, spacing=1.4)
rect(s, ML + 0.4, 4.6, 1.6, 0.05, fill=RED)
txt(s, ML + 0.35, 4.95, 10.0, 0.6,
    "难，是因为很多东西没做到位、认知没上来。\n这些都做到位了，你就知道该往哪里补。", 15, False, RGBColor(0xB0, 0xB0, 0xB0), spacing=1.6)
txt(s, W - MR - 2.4, H - 0.75, 2.4, 0.3, "张医生 · AI 营养师", 11, False,
    RGBColor(0x8A, 0x8A, 0x8A), PP_ALIGN.RIGHT)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "张医生-事业指导-20261105-成功法则与业务路径.pptx")
prs.save(out)
print("saved:", out, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))
