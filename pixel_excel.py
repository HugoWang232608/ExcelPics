import sys
import os
import argparse
import xlsxwriter
from PIL import Image


def create_pixel_art(input_path, output_path, width_cells, max_colors, show_grid):
    """
    核心转换逻辑
    """
    print(f"[-] 正在读取图片: {input_path}")

    try:
        img = Image.open(input_path).convert('RGB')
    except Exception as e:
        print(f"[!] 错误: 无法打开图片 - {e}")
        sys.exit(1)

    # 1. 计算尺寸
    aspect_ratio = img.size[1] / img.size[0]
    height_cells = int(width_cells * aspect_ratio)
    print(f"[-] 目标网格尺寸: {width_cells} (宽) x {height_cells} (高)")

    # 2. 调整大小 & 颜色量化 (核心算法)
    img_resized = img.resize((width_cells, height_cells), Image.Resampling.NEAREST)

    print(f"[-] 正在进行色彩量化 (最大颜色数: {max_colors})...")
    # 必须先 quantize 再转回 RGB 才能获取量化后的色值
    img_quantized = img_resized.quantize(colors=max_colors).convert('RGB')

    # 3. 生成 Excel
    print(f"[-] 正在生成 Excel 文件: {output_path}")
    try:
        workbook = xlsxwriter.Workbook(output_path)
        worksheet = workbook.add_worksheet("Pixel Art")

        # 隐藏默认网格线 (视图更干净)
        worksheet.hide_gridlines(2)

        # 设置行列尺寸 (模拟正方形)
        # 列宽 2.2 (字符) ≈ 行高 15 (磅)
        worksheet.set_column(0, width_cells - 1, 2.2)
        for row in range(height_cells):
            worksheet.set_row(row, 15)

        # 格式缓存
        format_cache = {}
        pixel_data = img_quantized.load()

        for y in range(height_cells):
            for x in range(width_cells):
                r, g, b = pixel_data[x, y]
                hex_color = '#{:02x}{:02x}{:02x}'.format(r, g, b)

                if hex_color not in format_cache:
                    props = {'bg_color': hex_color}
                    if show_grid:
                        props.update({'border': 1, 'border_color': '#E0E0E0'})  # 极细浅灰边框
                    format_cache[hex_color] = workbook.add_format(props)

                worksheet.write_blank(y, x, '', format_cache[hex_color])

        workbook.close()
        print(f"[√] 成功! 文件已保存至: {output_path}")
        print(f"[-] 最终使用颜色数: {len(format_cache)}")

    except Exception as e:
        print(f"[!] 保存 Excel 失败: {e}")
        print("    提示: 请确保目标文件没有被其他程序打开。")


def main():
    parser = argparse.ArgumentParser(
        description="将图片转换为 Excel 像素画工具 (Pixel Art to Excel Converter)",
        formatter_class=argparse.RawTextHelpFormatter
    )

    # 必须参数
    parser.add_argument("input", help="输入图片的路径 (例如: input.jpg)")

    # 可选参数
    parser.add_argument("-o", "--output", help="输出 Excel 的路径 (默认: [原文件名].xlsx)")
    parser.add_argument("-w", "--width", type=int, default=150, help="Excel 横向格数，决定精细度 (默认: 150)")
    parser.add_argument("-c", "--colors", type=int, default=48, help="最大颜色数量限制 (默认: 48)")
    parser.add_argument("--no-grid", action="store_true", help="如果添加此标记，则不显示像素格之间的微弱网格线")

    args = parser.parse_args()

    # 处理输入路径
    if not os.path.exists(args.input):
        print(f"[!] 错误: 找不到文件 '{args.input}'")
        sys.exit(1)

    # 处理输出路径 (如果没有指定，自动生成)
    output_path = args.output
    if not output_path:
        base_name = os.path.splitext(args.input)[0]
        output_path = f"{base_name}.xlsx"

    # 执行转换
    create_pixel_art(
        input_path=args.input,
        output_path=output_path,
        width_cells=args.width,
        max_colors=args.colors,
        show_grid=not args.no_grid
    )


if __name__ == "__main__":
    main()