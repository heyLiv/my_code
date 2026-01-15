import rawpy
from PIL import Image
import os

# ---------- 配置区域 ----------
source_folder = r"D:\RAW_photos"    # 母文件夹，替换成你的RAW根目录
output_folder = r"D:\JPEG_output"   # 输出根目录
jpeg_quality = 90                   # JPEG压缩质量
# -------------------------------

# 遍历母文件夹及所有子文件夹
for root, dirs, files in os.walk(source_folder):
    for file in files:
        if file.lower().endswith(('.cr2', '.nef', '.arw', '.rw2', '.orf', '.raf')):  # 常见RAW格式
            raw_path = os.path.join(root, file)
            # 构建输出路径，保持原目录结构
            relative_path = os.path.relpath(root, source_folder)
            output_dir = os.path.join(output_folder, relative_path)
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, os.path.splitext(file)[0] + '.jpg')
            
            # 读取RAW并保存为JPEG
            try:
                with rawpy.imread(raw_path) as raw:
                    rgb = raw.postprocess()
                    img = Image.fromarray(rgb)
                    img.save(output_path, 'JPEG', quality=jpeg_quality)
                print(f"已处理: {raw_path} → {output_path}")
            except Exception as e:
                print(f"转换失败: {raw_path}, 错误: {e}")

print("批量转换完成！")
