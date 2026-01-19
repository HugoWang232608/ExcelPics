# Pixel Excel Tool (Excel 电子刺绣/像素画生成器)

这是一个基于 Python 的命令行工具，可以将任何图像（JPG, PNG 等）转换为高精度的 Excel 像素画。它通过读取图像像素，自动调整 Excel 单元格大小并填充背景色来实现“电子刺绣”效果。

## ✨ 功能特点

* **自动调整网格**：自动设置行列宽高，使单元格呈正方形。
* **智能色彩量化**：自动提取图片主题色，避免杂乱，同时保持丰富细节。
* **高度可定制**：可通过命令行参数控制画面的精细度（格数）和色彩丰富度。
* **轻量级**：仅依赖 Pillow 和 XlsxWriter。

## 📦 安装指南

在使用之前，请确保你已经安装了 Python 3.x。

1. **安装依赖库**：
   在终端或命令行中运行以下命令：

   ```bash
   pip install Pillow XlsxWriter
2. 下载脚本：确保 pixel_excel.py 在你的工作目录中。

🚀 使用方法基础用法最简单的用法只需提供图片路径。程序会自动生成同名的 Excel 文件。
```Bash
python pixel_excel.py input.jpg
```
这将在同目录下生成 input.xlsx进阶用法你可以通过参数控制输出效果。
```Bash
python pixel_excel.py input.jpg -o my_art.xlsx -w 200 -c 64
```

参数说明

| 参数        | 缩写      | 说明                   | 默认值         |
|-----------|---------|----------------------|-------------|
| input     | N/A(必填) | 输入图片路径               | N/A         |
| --output  | -o      | 输出 Excel 文件路径        | [原文件名].xlsx |
| --width   | -w      | 宽度格子数量 (精细度)。数值越大越清晰 | 150         |
| --colors  | -c      | 限制最大颜色数量。数值越大色彩越丰富   | 48          |
| --no-grid | N/A     | 添加此标记将隐藏格子间的细网格线     | 默认显示网格      |

🎨 效果调整建议
1. 想要复古游戏风格 (8-bit):使用较少的格子和较少的颜色。 
```Bash
python pixel_excel.py mario.png -w 60 -c 16
```
2. 想要照片级精细还原:增加格子数量和颜色上限。
```Bash
python pixel_excel.py photo.jpg -w 250 -c 128
```
3. 去除网格线:如果你希望色块之间无缝连接（纯平涂风格）：
```Bash
python pixel_excel.py logo.png --no-grid
```
⚠️ 注意事项
1. 生成过程中请不要打开同名的 Excel 文件，否则程序会因为无法写入而报错。
2. 图片过大（如 -w 500 以上）会导致 Excel 文件体积剧增并可能导致打开缓慢。