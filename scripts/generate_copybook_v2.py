from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime

# ===================== 精准配置：8mm 米字格 A4 字帖 =====================
DPI = 300
A4_W = int(210 * DPI / 25.4)
A4_H = int(297 * DPI / 25.4)

# 单格 = 0.8 cm (8mm)
GRID_SIZE_MM = 10
GRID_SIZE_PX = int(GRID_SIZE_MM * DPI / 25.4)

# A4 排版（左右留边）
MARGIN = int(15 * DPI / 25.4)  # 1.5cm边距
TITLE_AREA_HEIGHT = int(35 * DPI / 25.4)  # 标题区域高度 3.5cm

COLS = (A4_W - 2 * MARGIN) // GRID_SIZE_PX  # 自动算列数
ROWS = (A4_H - 2 * MARGIN - TITLE_AREA_HEIGHT) // GRID_SIZE_PX  # 自动算行数

# 字帖内容（从文件读取）
text_file = os.path.join(os.path.dirname(__file__), "copybook_text.txt")
title = "字帖"
text = ""
if os.path.exists(text_file):
    with open(text_file, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]
        if lines:
            # 第一行作为标题
            title = lines[0]
            # 剩余行作为正文
            text = "".join(lines[1:])
            print(f"成功提取标题: {title}")
        else:
            text = ""
else:
    text = "请在 scripts 目录下创建 copybook_text.txt 文件并写入内容。"

GRID_COLOR = "#FF4500"  # 橙红色格线
TEXT_COLOR = "#808080"  # 灰色文字
# ==========================================================================

# 字体路径处理
fonts_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "fonts")
available_fonts = [
    f for f in os.listdir(fonts_dir) if f.lower().endswith((".ttf", ".otf"))
]

if available_fonts:
    print("\n可用字体列表:")
    for i, f in enumerate(available_fonts):
        print(f"{i + 1}. {f}")
    
    try:
        choice = input(f"\n请选择字体编号 (1-{len(available_fonts)}, 默认 1): ").strip()
        if not choice:
            idx = 0
        else:
            idx = int(choice) - 1
            if idx < 0 or idx >= len(available_fonts):
                print(f"无效选择，使用默认字体")
                idx = 0
    except ValueError:
        print("输入无效，使用默认字体")
        idx = 0
        
    selected_font = available_fonts[idx]
    font_path = os.path.join(fonts_dir, selected_font)
    print(f"已选择字体: {selected_font}\n")
else:
    font_path = "C:/Windows/Fonts/simkai.ttf"
    print("未在 fonts 目录找到字体，尝试系统字体")

# 新增功能：选择每个字留 n 个练习格
try:
    practice_count = input("每个字后面留几个练习格？(默认 0): ").strip()
    practice_count = int(practice_count) if practice_count else 0
except ValueError:
    print("输入无效，默认不留练习格")
    practice_count = 0

# 构建最终要绘制的字符列表
raw_chars = [c for c in text.strip() if c.strip()]
chars_to_draw = []
for c in raw_chars:
    chars_to_draw.append(c)  # 添加原字
    for _ in range(practice_count):
        chars_to_draw.append(None)  # 添加练习格标记

for page in range((len(chars_to_draw) + COLS * ROWS - 1) // (COLS * ROWS)):
    img = Image.new("RGB", (A4_W, A4_H), "white")
    draw = ImageDraw.Draw(img)

    # 加载字体
    try:
        font = ImageFont.truetype(font_path, int(GRID_SIZE_PX * 0.8))
        title_font_size = int(TITLE_AREA_HEIGHT * 0.5)
        title_font = ImageFont.truetype(font_path, title_font_size)
    except Exception as e:
        print(f"无法加载字体: {e}")
        font = ImageFont.load_default()
        title_font = ImageFont.load_default()

    # 1. 绘制标题
    try:
        bbox = draw.textbbox((0, 0), title, font=title_font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        tx = (A4_W - tw) // 2 - bbox[0]
        ty = MARGIN + (TITLE_AREA_HEIGHT - th) // 2 - bbox[1]
        draw.text((tx, ty), title, font=title_font, fill="#000000")
    except:
        tw, th = draw.textsize(title, font=title_font)
        draw.text(((A4_W - tw) // 2, MARGIN + (TITLE_AREA_HEIGHT - th) // 2), title, font=title_font, fill="#000000")

    # 2. 绘制米字格和正文
    for i in range(COLS * ROWS):
        idx = page * COLS * ROWS + i
        if idx >= len(chars_to_draw):
            break
        c = chars_to_draw[idx]

        col = i % COLS
        row = i // COLS

        gx = MARGIN + col * GRID_SIZE_PX
        gy = MARGIN + TITLE_AREA_HEIGHT + row * GRID_SIZE_PX

        # 外框
        draw.rectangle(
            [gx, gy, gx + GRID_SIZE_PX, gy + GRID_SIZE_PX], outline=GRID_COLOR, width=2
        )

        # 米字格虚线
        cx, cy = gx + GRID_SIZE_PX // 2, gy + GRID_SIZE_PX // 2
        draw.line([gx, gy, gx + GRID_SIZE_PX, gy + GRID_SIZE_PX], fill=GRID_COLOR, width=1)
        draw.line([gx + GRID_SIZE_PX, gy, gx, gy + GRID_SIZE_PX], fill=GRID_COLOR, width=1)
        draw.line([cx, gy, cx, gy + GRID_SIZE_PX], fill=GRID_COLOR, width=1)
        draw.line([gx, cy, gx + GRID_SIZE_PX, cy], fill=GRID_COLOR, width=1)

        # 如果不是练习格，则绘制文字
        if c is not None:
            try:
                bbox = draw.textbbox((0, 0), c, font=font)
                w = bbox[2] - bbox[0]
                h = bbox[3] - bbox[1]
                draw.text(
                    (
                        gx + (GRID_SIZE_PX - w) // 2 - bbox[0],
                        gy + (GRID_SIZE_PX - h) // 2 - bbox[1],
                    ),
                    c,
                    font=font,
                    fill=TEXT_COLOR,
                )
            except:
                w, h = draw.textsize(c, font=font)
                draw.text((gx + (GRID_SIZE_PX - w) // 2, gy + (GRID_SIZE_PX - h) // 2), c, font=font, fill=TEXT_COLOR)

    # 保存文件
    font_name = os.path.splitext(os.path.basename(font_path))[0]
    now = datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H%M%S")

    out_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "out", date_str)
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)

    output_name = f"{title}_{font_name}_{time_str}_第{page+1}页.png"
    output_path = os.path.join(out_dir, output_name)
    img.save(output_path)
    print(f"已生成：{output_path}")
