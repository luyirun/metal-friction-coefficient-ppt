from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math
import random

OUT = Path(r"F:\Users\user\Documents\测金属摩擦系数\outputs")
OUT.mkdir(exist_ok=True)

FONT_REG = r"C:\Windows\Fonts\Noto Sans SC (TrueType).otf"
FONT_MED = r"C:\Windows\Fonts\Noto Sans SC Medium (TrueType).otf"
FONT_BOLD = r"C:\Windows\Fonts\Noto Sans SC Bold (TrueType).otf"


slides = [
    ("金属与冰摩擦系数的测量", "搭建可重复、可调控的测量系统", ["金属-冰接触面", "实验装置", "规律分析"]),
    ("选题背景", "冰上摩擦影响极地航运阻力与能耗", ["航程缩短 25%-55%", "油耗增加 13%-15%", "船冰阻力"]),
    ("题目解读", "完成机制分析、装置搭建、结果评估", ["滑动摩擦机制", "测量方案", "不确定度"]),
    ("方法选择", "动态斜面法反推出动摩擦系数", ["ma = mg sinθ - μmg cosθ", "运动学采集", "受力模型"]),
    ("装置设计", "斜面冰面与竖直冰面双结构验证", ["斜面冰槽", "低温传感器", "数据采集"]),
    ("变量设计", "冰体维度与金属维度分变量控制", ["温度", "压强", "材质", "粗糙度", "面积"]),
    ("问题改进", "解决冰面平整、低温保持与传感稳定", ["去离子水", "亚克力盖板", "外置传感器"]),
    ("数据处理", "由 x、t、θ 得到 a，再计算 μ", ["x,t → a", "μ 反演", "不确定度传播"]),
    ("结果分析", "比较不同条件下摩擦系数变化趋势", ["μ-温度", "μ-压力", "μ-粗糙度"]),
    ("总结展望", "自制装置、多变量测量、工程应用", ["双结构验证", "低温优化", "极地装备"]),
]


def font(size, weight="reg"):
    p = {"reg": FONT_REG, "med": FONT_MED, "bold": FONT_BOLD}.get(weight, FONT_REG)
    return ImageFont.truetype(p, size)


def wrap_text(draw, text, fnt, max_w):
    if not text:
        return []
    lines, cur = [], ""
    for ch in text:
        test = cur + ch
        if draw.textbbox((0, 0), test, font=fnt)[2] <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = ch
    if cur:
        lines.append(cur)
    return lines


def rounded(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def draw_text_block(draw, xy, text, fnt, fill, max_w, line_gap=5, max_lines=3):
    x, y = xy
    for line in wrap_text(draw, text, fnt, max_w)[:max_lines]:
        draw.text((x, y), line, font=fnt, fill=fill)
        y += fnt.size + line_gap
    return y


def draw_footer(draw, w, h, idx, style):
    if style == "polar":
        draw.line((42, h - 34, w - 42, h - 34), fill=(169, 202, 213), width=1)
        draw.text((42, h - 28), "Metal-Ice Friction Measurement", font=font(13, "reg"), fill=(86, 111, 119))
        draw.text((w - 72, h - 30), f"{idx:02d}", font=font(16, "bold"), fill=(12, 53, 68))
    elif style == "lab":
        draw.rectangle((0, h - 42, w, h), fill=(16, 22, 27))
        draw.text((38, h - 30), "LOW TEMPERATURE TRIBOLOGY LAB", font=font(12, "med"), fill=(116, 255, 210))
        draw.text((w - 70, h - 33), f"{idx:02d}", font=font(19, "bold"), fill=(235, 242, 240))
    else:
        draw.rectangle((0, h - 36, w, h), fill=(238, 242, 244))
        draw.text((36, h - 26), "金属与冰摩擦系数测量", font=font(12, "med"), fill=(48, 58, 64))
        draw.text((w - 68, h - 29), f"{idx:02d}", font=font(18, "bold"), fill=(0, 104, 129))


def slide_polar(idx, title, point, bullets):
    w, h = 640, 360
    im = Image.new("RGB", (w, h), (238, 247, 248))
    d = ImageDraw.Draw(im)
    for y in range(h):
        c = int(248 - y * 0.05)
        d.line((0, y, w, y), fill=(c, min(255, c + 4), min(255, c + 6)))
    d.polygon([(0, 250), (145, 205), (310, 255), (510, 190), (640, 230), (640, 360), (0, 360)], fill=(205, 229, 232))
    d.polygon([(0, 292), (210, 245), (420, 310), (640, 260), (640, 360), (0, 360)], fill=(177, 214, 220))
    if idx in (1, 2, 5):
        d.rounded_rectangle((365, 75, 592, 215), 18, fill=(18, 73, 91))
        d.arc((390, 95, 530, 235), 185, 340, fill=(241, 250, 251), width=12)
        d.rectangle((420, 163, 572, 176), fill=(116, 170, 180))
        d.polygon([(492, 122), (575, 176), (464, 176)], fill=(236, 245, 246))
    else:
        d.rounded_rectangle((368, 68, 590, 230), 18, fill=(248, 252, 252), outline=(145, 188, 198), width=2)
        d.line((405, 195, 555, 118), fill=(0, 119, 145), width=5)
        d.rectangle((452, 125, 508, 150), fill=(77, 91, 99))
        d.line((452, 150, 508, 150), fill=(28, 40, 47), width=3)
        d.text((396, 78), "μ", font=font(46, "bold"), fill=(0, 104, 129))
    d.text((42, 42), title, font=font(30, "bold"), fill=(8, 44, 58))
    draw_text_block(d, (42, 91), point, font(18, "med"), (0, 91, 111), 285, 5, 3)
    y = 176
    for b in bullets[:4]:
        d.ellipse((44, y + 6, 52, y + 14), fill=(0, 132, 161))
        y = draw_text_block(d, (62, y), b, font(15, "reg"), (38, 61, 68), 250, 2, 1) + 7
    draw_footer(d, w, h, idx, "polar")
    return im


def slide_lab(idx, title, point, bullets):
    w, h = 640, 360
    im = Image.new("RGB", (w, h), (8, 13, 17))
    d = ImageDraw.Draw(im)
    for x in range(0, w, 32):
        d.line((x, 0, x, h), fill=(18, 31, 37), width=1)
    for y in range(0, h, 32):
        d.line((0, y, w, y), fill=(18, 31, 37), width=1)
    d.rectangle((0, 0, 12, h), fill=(0, 214, 170))
    d.text((38, 34), f"0{idx}", font=font(19, "bold"), fill=(0, 214, 170))
    d.text((38, 66), title, font=font(29, "bold"), fill=(239, 247, 245))
    draw_text_block(d, (38, 112), point, font(17, "med"), (167, 230, 216), 292, 6, 3)
    if idx in (4, 8):
        d.rounded_rectangle((358, 62, 594, 240), 8, fill=(14, 29, 35), outline=(0, 214, 170), width=2)
        d.line((392, 198, 560, 114), fill=(0, 214, 170), width=4)
        d.text((387, 77), "ma = mg sinθ - μmg cosθ", font=font(15, "med"), fill=(239, 247, 245))
        d.arc((404, 170, 468, 234), 300, 355, fill=(255, 189, 89), width=3)
    elif idx in (6, 9):
        for i, b in enumerate(bullets[:5]):
            x = 354 + (i % 2) * 120
            y = 68 + (i // 2) * 58
            d.rounded_rectangle((x, y, x + 104, y + 38), 6, fill=(18, 42, 48), outline=(49, 95, 97))
            d.text((x + 12, y + 9), b[:7], font=font(13, "med"), fill=(228, 242, 237))
    else:
        d.rounded_rectangle((358, 62, 594, 240), 8, fill=(14, 29, 35), outline=(49, 95, 97))
        for i in range(4):
            d.line((378, 206 - i * 31, 570, 177 - i * 18), fill=(0, 214, 170, ), width=3)
        d.ellipse((438, 94, 514, 170), outline=(255, 189, 89), width=5)
        d.text((460, 114), "μ", font=font(34, "bold"), fill=(255, 189, 89))
    y = 184
    for b in bullets[:3]:
        d.rectangle((40, y + 4, 48, y + 12), fill=(255, 189, 89))
        y = draw_text_block(d, (59, y), b, font(14, "reg"), (224, 233, 231), 250, 2, 1) + 8
    draw_footer(d, w, h, idx, "lab")
    return im


def slide_report(idx, title, point, bullets):
    w, h = 640, 360
    im = Image.new("RGB", (w, h), (250, 250, 247))
    d = ImageDraw.Draw(im)
    d.rectangle((0, 0, w, 62), fill=(236, 241, 241))
    d.rectangle((0, 0, 16, h), fill=(0, 111, 135))
    d.text((42, 25), title, font=font(27, "bold"), fill=(25, 36, 42))
    d.rounded_rectangle((42, 88, 310, 152), 6, fill=(255, 255, 255), outline=(219, 226, 226))
    draw_text_block(d, (58, 105), point, font(16, "med"), (0, 94, 116), 228, 5, 2)
    if idx in (2, 9):
        axes = (358, 92, 585, 236)
        d.line((axes[0], axes[3], axes[2], axes[3]), fill=(70, 80, 82), width=2)
        d.line((axes[0], axes[1], axes[0], axes[3]), fill=(70, 80, 82), width=2)
        pts = [(370, 220), (420, 180), (470, 190), (520, 128), (570, 108)]
        d.line(pts, fill=(0, 111, 135), width=4)
        for p in pts:
            d.ellipse((p[0]-4, p[1]-4, p[0]+4, p[1]+4), fill=(231, 98, 69))
        d.text((366, 244), "变量", font=font(12), fill=(80, 89, 92))
        d.text((548, 74), "μ", font=font(18, "bold"), fill=(0, 111, 135))
    elif idx in (4, 5, 8):
        d.rounded_rectangle((352, 84, 590, 235), 6, fill=(255, 255, 255), outline=(210, 219, 220))
        d.line((384, 202, 558, 128), fill=(0, 111, 135), width=4)
        d.rectangle((440, 139, 502, 164), fill=(77, 84, 88))
        d.text((386, 96), "θ", font=font(25, "bold"), fill=(231, 98, 69))
    else:
        d.rounded_rectangle((352, 84, 590, 235), 6, fill=(255, 255, 255), outline=(210, 219, 220))
        cx, cy = 470, 157
        for a, b in zip(range(0, 360, 60), bullets * 2):
            r = math.radians(a)
            x, y = cx + math.cos(r) * 83, cy + math.sin(r) * 57
            d.line((cx, cy, x, y), fill=(194, 207, 209), width=2)
            d.ellipse((x-22, y-14, x+22, y+14), fill=(238, 242, 242), outline=(0, 111, 135))
            d.text((x-16, y-8), b[:2], font=font(12, "med"), fill=(0, 91, 111))
        d.ellipse((cx-30, cy-22, cx+30, cy+22), fill=(0, 111, 135))
        d.text((cx-12, cy-14), "μ", font=font(25, "bold"), fill=(255, 255, 255))
    y = 176
    for b in bullets[:4]:
        d.rounded_rectangle((46, y, 220, y + 26), 5, fill=(239, 243, 243), outline=(224, 231, 231))
        d.text((58, y + 5), b, font=font(12, "med"), fill=(44, 55, 60))
        y += 34
    draw_footer(d, w, h, idx, "report")
    return im


def make_board(name, renderer, bg, title, subtitle):
    thumb_w, thumb_h = 640, 360
    gap, margin = 34, 54
    cols, rows = 2, 5
    header = 120
    board = Image.new("RGB", (cols * thumb_w + (cols - 1) * gap + margin * 2, rows * thumb_h + (rows - 1) * gap + margin * 2 + header), bg)
    d = ImageDraw.Draw(board)
    d.text((margin, 34), title, font=font(42, "bold"), fill=(245, 248, 248) if sum(bg) < 300 else (20, 32, 38))
    d.text((margin, 84), subtitle, font=font(18, "med"), fill=(178, 210, 214) if sum(bg) < 300 else (80, 96, 102))
    for i, s in enumerate(slides, 1):
        x = margin + ((i - 1) % cols) * (thumb_w + gap)
        y = margin + header + ((i - 1) // cols) * (thumb_h + gap)
        sh = renderer(i, *s)
        board.paste(sh, (x, y))
        ImageDraw.Draw(board).rectangle((x, y, x + thumb_w, y + thumb_h), outline=(255, 255, 255) if sum(bg) < 300 else (198, 210, 212), width=2)
    path = OUT / name
    board.save(path, quality=95)
    return path


paths = [
    make_board("ppt_style_01_polar_engineering.jpg", slide_polar, (218, 236, 239), "方案 01  极地工程清透风", "冰川浅色背景 + 青蓝科研图表 + 真实工程感，适合正式答辩"),
    make_board("ppt_style_02_dark_lab_dashboard.jpg", slide_lab, (5, 10, 13), "方案 02  暗色实验室仪表风", "深色网格 + 荧光数据 + 传感器界面感，突出科技与装置创新"),
    make_board("ppt_style_03_academic_report.jpg", slide_report, (232, 237, 238), "方案 03  学术报告图册风", "白底论文式排版 + 数据图表 + 稳重页眉页脚，适合评委阅读"),
]

for p in paths:
    print(p)
