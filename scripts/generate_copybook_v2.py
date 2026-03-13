from PIL import Image, ImageDraw, ImageFont
import os

# ===================== 精准配置：8mm 米字格 A4 字帖 =====================
DPI = 300
A4_W = int(210 * DPI / 25.4)
A4_H = int(297 * DPI / 25.4)

# 单格 = 0.8 cm (8mm)
GRID_SIZE_MM = 10
GRID_SIZE_PX = int(GRID_SIZE_MM * DPI / 25.4)

# A4 排版（左右留边）
MARGIN = int(15 * DPI / 25.4)  # 1.5cm边距
COLS = (A4_W - 2 * MARGIN) // GRID_SIZE_PX  # 自动算列数
ROWS = (A4_H - 2 * MARGIN) // GRID_SIZE_PX  # 自动算行数

# 字帖内容（从文件读取）
text_file = os.path.join(os.path.dirname(__file__), "copybook_text.txt")
if os.path.exists(text_file):
    with open(text_file, "r", encoding="utf-8") as f:
        text = f.read()
else:
    text = "请在 scripts 目录下创建 copybook_text.txt 文件并写入内容。"

GRID_COLOR = "#FF4500"  # 更淡的灰色，用于描写
TEXT_COLOR = "#808080"  # 浅灰色文字，方便临摹描写
# ==========================================================================

chars = [c for c in text.strip() if c.strip()]

# 字体路径处理
font_path = "方正粗黑宋简体.ttf"
if not os.path.exists(font_path):
    font_path = "C:/Windows/Fonts/方正粗黑宋简体.ttf"

for page in range((len(chars) + COLS * ROWS - 1) // (COLS * ROWS)):
    img = Image.new("RGB", (A4_W, A4_H), "white")
    draw = ImageDraw.Draw(img)

    # 楷书
    try:
        font = ImageFont.truetype(font_path, int(GRID_SIZE_PX * 0.8))
    except Exception as e:
        print(f"无法加载字体: {e}")
        # 尝试使用默认字体
        font = ImageFont.load_default()

    for i in range(COLS * ROWS):
        idx = page * COLS * ROWS + i
        if idx >= len(chars):
            break
        c = chars[idx]

        col = i % COLS
        row = i // COLS

        gx = MARGIN + col * GRID_SIZE_PX
        gy = MARGIN + row * GRID_SIZE_PX

        # 外框
        draw.rectangle(
            [gx, gy, gx + GRID_SIZE_PX, gy + GRID_SIZE_PX], outline=GRID_COLOR, width=2
        )

        # 米字格虚线
        cx, cy = gx + GRID_SIZE_PX // 2, gy + GRID_SIZE_PX // 2

        # 对角线
        draw.line(
            [gx, gy, gx + GRID_SIZE_PX, gy + GRID_SIZE_PX], fill=GRID_COLOR, width=1
        )
        draw.line(
            [gx + GRID_SIZE_PX, gy, gx, gy + GRID_SIZE_PX], fill=GRID_COLOR, width=1
        )
        # 十字线
        draw.line([cx, gy, cx, gy + GRID_SIZE_PX], fill=GRID_COLOR, width=1)
        draw.line([gx, cy, gx + GRID_SIZE_PX, cy], fill=GRID_COLOR, width=1)

        # 文字居中
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
            # 兼容旧版本 Pillow
            w, h = draw.textsize(c, font=font)
            draw.text(
                (gx + (GRID_SIZE_PX - w) // 2, gy + (GRID_SIZE_PX - h) // 2),
                c,
                font=font,
                fill=TEXT_COLOR,
            )

    output_name = f"字帖_A4_8mm米字格_第{page+1}页.png"
    img.save(output_name)
    print(f"已生成：{output_name}")
