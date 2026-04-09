from PIL import Image
import os

def compress_image(input_path, output_path, target_kb=10, target_size=(295, 413)):
    """
    压缩图片到指定大小以内，并调整为指定像素尺寸。
    """
    try:
        original_size = os.path.getsize(input_path) / 1024
        print(f"原始文件大小: {original_size:.2f} KB")

        with Image.open(input_path) as img:
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            
            # 1. 等比例缩放并居中裁剪 (不拉伸)
            target_w, target_h = target_size
            orig_w, orig_h = img.size
            
            # 计算缩放比例（取较大者以覆盖目标区域）
            ratio = max(target_w / orig_w, target_h / orig_h)
            new_w = int(orig_w * ratio)
            new_h = int(orig_h * ratio)
            
            print(f"等比例缩放至: {new_w}x{new_h}，然后裁剪为: {target_w}x{target_h}")
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            
            # 居中裁剪
            left = (new_w - target_w) / 2
            top = (new_h - target_h) / 2
            right = (new_w + target_w) / 2
            bottom = (new_h + target_h) / 2
            img = img.crop((left, top, right, bottom))
            
            # 2. 迭代调整质量以满足文件大小要求
            quality = 95
            while True:
                img.save(output_path, "JPEG", quality=quality, optimize=True)
                compressed_size = os.path.getsize(output_path) / 1024
                
                if compressed_size <= target_kb or quality <= 5:
                    break
                
                quality -= 5

        print(f"最终文件大小: {compressed_size:.2f} KB (Quality: {quality}, Size: {target_size[0]}x{target_size[1]})")
        print(f"已保存至: {output_path}")

    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == "__main__":
    input_file = "11.jpg"
    output_file = "11_compressed.jpg"
    
    if os.path.exists(input_file):
        # 目标：10KB以内，尺寸 295x413
        compress_image(input_file, output_file, target_kb=10, target_size=(295, 413))
    else:
        print(f"错误: 找不到文件 {input_file}")
