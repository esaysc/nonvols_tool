import os
from PIL import Image, ImageDraw, ImageFont

def generate_copybook(
    char_list="一二三四五六七八九十",
    output_path="米字格字帖.png",
    grid_size=120,
    rows=2,
    cols=5,
    font_path="C:/Windows/Fonts/simhei.ttf",
    font_size=80,
    bg_color=(255, 255, 255),
    line_color=(180, 180, 180),
    text_color=(0, 0, 0)
):
    """
    生成米字格字帖图片
    """
    # 检查字体文件
    if not os.path.exists(font_path):
        # 尝试一些常见的 Windows 字体路径
        alt_fonts = [
            "C:/Windows/Fonts/msyh.ttc",  # 微软雅黑
            "C:/Windows/Fonts/simsun.ttc", # 宋体
        ]
        found = False
        for alt in alt_fonts:
            if os.path.exists(alt):
                font_path = alt
                found = True
                break
        if not found:
            print(f"错误: 找不到字体文件 {font_path}")
            return

    # 初始化画布
    canvas_w = cols * grid_size
    canvas_h = rows * grid_size
    img = Image.new("RGB", (canvas_w, canvas_h), bg_color)
    draw = ImageDraw.Draw(img)

    # 加载字体
    try:
        font = ImageFont.truetype(font_path, font_size)
    except Exception as e:
        print(f"加载字体失败: {e}")
        return

    # 绘制米字格+文字
    for row in range(rows):
        for col in range(cols):
            # 计算当前格子左上角坐标
            x0 = col * grid_size
            y0 = row * grid_size
            x1 = x0 + grid_size
            y1 = y0 + grid_size
            center_x, center_y = (x0 + x1) / 2, (y0 + y1) / 2

            # 1. 画外框
            draw.rectangle([x0, y0, x1, y1], outline=line_color, width=2)
            
            # 2. 画虚线效果（通过绘制小段直线模拟）
            def draw_dash_line(p1, p2, dash_length=4):
                # 简化版：米字格内部线条使用细线
                draw.line([p1, p2], fill=line_color, width=1)

            # 横、竖中线
            draw_dash_line((x0, center_y), (x1, center_y))
            draw_dash_line((center_x, y0), (center_x, y1))
            # 对角线
            draw_dash_line((x0, y0), (x1, y1))
            draw_dash_line((x0, y1), (x1, y0))

            # 3. 绘制文字（居中）
            idx = row * cols + col
            if idx < len(char_list):
                char = char_list[idx]
                
                # 使用 getbbox 替代已弃用的 textsize
                # bbox 返回 (left, top, right, bottom)
                bbox = font.getbbox(char)
                w = bbox[2] - bbox[0]
                h = bbox[3] - bbox[1]
                
                # 计算偏移量以实现真正的视觉居中
                # font.getbbox(char) 的 top 可能是负值或非零值，取决于字体设计
                text_x = center_x - (bbox[2] + bbox[0]) / 2
                text_y = center_y - (bbox[3] + bbox[1]) / 2
                
                draw.text((text_x, text_y), char, fill=text_color, font=font)

    # 保存图片
    img.save(output_path)
    print(f"米字格生成完成！保存至: {output_path}")

if __name__ == "__main__":
    # 可以根据需要调整参数
    generate_copybook(
        char_list="永和九年岁在癸丑暮春之初",
        rows=3,
        cols=4,
        grid_size=150,
        font_size=100
    )
