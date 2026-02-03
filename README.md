# 🧩 电子刺绣/拼豆底稿生成器
# 工具一：Pixel Excel Tool (Excel 电子刺绣生成器)

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


# 工具二：拼豆专用生成器 (pixel_tool.py)

适用于实体制作（拼豆、十字绣）。它会读取 colorMap.json，强制将图片颜色对应到你拥有的真实材料色号上。

✨ 核心功能

- 真实色号映射：支持加载 JSON 色卡，将像素点转换为具体的色号名称（如 "A01", "S14"）。 
- 智能色系过滤：支持严格模式：精确锁定特定品牌（如 Mard-221），自动排除名称相似但品牌不同的系列（如 优肯Mard-221）。
- 支持模糊搜索：输入简称（如 COCO）即可自动匹配相关系列。
- 双工作表输出：Sheet1: 图纸，正方形单元格，内置色号文字（如 A10），方便对照制作。Sheet2: 采购清单，列出所需色号、系列、预览色及精确的颗粒数量。
- 可视化辅助：自动调整单元格尺寸以容纳色号文字，保持视觉正方形。

🚀 使用方法
1. 查看可用色系不知道 colorMap.json 里有哪些牌子的豆子？先查一下： 
```Bash
python pixel_tool.py --list colorMap.json
```
2. 生成图纸 (指定色系)推荐用法。指定你要用的豆子系列（例如 COCO-291）：
```Bash
python pixel_tool.py input.jpg colorMap.json -s COCO-291 -w 50
```
这会生成 pixel_pattern.xlsx，包含图纸和用量清单。
3. 自动匹配如果不指定 -s，工具会自动选择 JSON 中包含颜色数量最多的那个系列：
```Bash
python pixel_tool.py input.jpg colorMap.json
```

参数说明

|参数| 缩写   | 说明                | 默认值         |
|-----|------|-------------------|-------------|
|image| (必填) |输入图片路径| N/A         |
|config| (必填) | 色号配置文件路径(通常是 colorMap.json) | N/A         |
|--list|N/A|仅列出 JSON 中包含的所有色系名称 (不生成图纸)| N/A         |
|--series|-s|指定色系名称 (如 Mard-221)。支持严格匹配和模糊匹配。| 自动选择颜色最多的系列 |
|--width|-w|宽度格子数量 (建议 30-80 之间，取决于拼豆板大小)|50|
|--output|-o|输出文件名|pixel_pattern.xlsx|

🎨 效果调整建议

制作大型拼豆挂画使用 pixel_tool.py，根据你的拼豆板数量计算宽度（例如 4块大板子 ≈ 100格）：
```Bash
python pixel_tool.py landscape.jpg colorMap.json -w 100 -s COCO-291
```

⚠️ 注意事项

- 文件占用：生成过程中请不要打开同名的 Excel 文件，否则程序会报错（Permission denied）。
- 性能：图片过大（如 -w 500 以上）会导致 Excel 文件体积剧增，打开缓慢。
- JSON 格式：pixel_tool.py 强依赖 colorMap.json 的结构，请确保 JSON 格式正确（Key为Hex颜色，Value为包含 colorName, colorTitle 的列表）。