import sys
import os
import json
import argparse
import xlsxwriter
from PIL import Image


def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4))


def get_all_series_titles(data):
    """提取所有存在的色系名称"""
    titles = set()
    for items in data.values():
        for item in items:
            t = item.get('colorTitle')
            if t:
                titles.add(t)
    return titles


def analyze_series(data):
    """分析 JSON 中包含哪些色系，并返回颜色最多的那个作为默认值"""
    series_counts = {}
    for items in data.values():
        for item in items:
            title = item.get('colorTitle', 'Unknown')
            series_counts[title] = series_counts.get(title, 0) + 1
    sorted_series = sorted(series_counts.items(), key=lambda x: x[1], reverse=True)
    return sorted_series


def load_palette_from_json(json_path, target_series=None):
    print(f"[-] 正在加载配置: {json_path}")
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"[!] JSON 读取失败: {e}")
        sys.exit(1)

    # 1. 确定目标色系
    if not target_series:
        sorted_series = analyze_series(data)
        if not sorted_series:
            print("[!] 配置文件中没有有效的色系信息！")
            sys.exit(1)
        target_series = sorted_series[0][0]
        print(f"[*] 未指定色系，自动选择颜色最全的: {target_series} ({sorted_series[0][1]}色)")

    # --- 关键改进：智能严格模式 ---
    all_titles = get_all_series_titles(data)

    # 如果用户指定的名称精确存在于数据库中（例如 "Mard-221"），则启用严格模式
    is_strict_mode = target_series in all_titles

    if is_strict_mode:
        print(f"[*] 检测到精确色系名称 '{target_series}'，已启用严格过滤模式 (拒绝 '优肯Mard-221' 等近似项)")
    else:
        print(f"[*] 色系 '{target_series}' 未完全匹配，启用模糊搜索模式...")

    palette_list = []
    palette_info = {}

    # 2. 遍历筛选
    match_count = 0
    for hex_key, items in data.items():
        if not items: continue

        selected_item = None

        # 策略A: 精确匹配 (最高优先级)
        for item in items:
            if item.get('colorTitle') == target_series:
                selected_item = item
                break

        # 策略B: 模糊匹配 (仅当不处于严格模式时才允许)
        if not selected_item and not is_strict_mode:
            for item in items:
                if target_series in item.get('colorTitle', ''):
                    selected_item = item
                    break

        if selected_item:
            try:
                rgb = hex_to_rgb(hex_key)
                color_name = selected_item.get('colorName', '???')
                color_title = selected_item.get('colorTitle', '')

                # 去重
                if color_name not in palette_info:
                    palette_list.append((rgb, color_name))
                    palette_info[color_name] = {
                        'name': color_name,
                        'title': color_title,
                        'rgb': rgb,
                        'hex': hex_key
                    }
                    match_count += 1
            except:
                continue

    if match_count == 0:
        print(f"[!] 警告: 在色系 '{target_series}' 下没有找到任何颜色！")
        if is_strict_mode:
            print(f"    (严格模式已开启，请检查 JSON 中是否有完全等于 '{target_series}' 的项)")
        sys.exit(1)

    print(f"[-] 色板构建完成: 有效颜色 {len(palette_list)} 种")
    return palette_list, palette_info


def get_closest_id(target_rgb, palette_list):
    min_dist = float('inf')
    closest_id = None
    r1, g1, b1 = target_rgb

    for p_rgb, p_id in palette_list:
        r2, g2, b2 = p_rgb
        dist_sq = (r1 - r2) ** 2 + (g1 - g2) ** 2 + (b1 - b2) ** 2
        if dist_sq < min_dist:
            min_dist = dist_sq
            closest_id = p_id
    return closest_id


def get_text_color(bg_rgb):
    lum = 0.299 * bg_rgb[0] + 0.587 * bg_rgb[1] + 0.114 * bg_rgb[2]
    return '#FFFFFF' if lum < 128 else '#000000'


def main():
    parser = argparse.ArgumentParser(description="拼豆图纸生成器 (智能匹配版)")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("image", nargs='?', help="输入图片路径")
    group.add_argument("--list", action="store_true", help="列出所有色系")

    parser.add_argument("config", help="色号JSON路径")
    parser.add_argument("-o", "--output", help="输出Excel路径")
    parser.add_argument("-w", "--width", type=int, default=50, help="宽度格子数")
    parser.add_argument("-s", "--series", help="指定色系 (智能匹配)")

    args = parser.parse_args()

    if args.list:
        with open(args.config, 'r', encoding='utf-8') as f:
            data = json.load(f)
        stats = analyze_series(data)
        print(f"{'色系名称':<25} | {'包含颜色'}")
        print("-" * 40)
        for t, c in stats: print(f"{t:<25} | {c}")
        return

    # 加载数据
    palette_list, palette_info = load_palette_from_json(args.config, args.series)

    # 图片处理
    try:
        img = Image.open(args.image).convert('RGB')
    except Exception as e:
        print(f"[!] 图片错误: {e}")
        sys.exit(1)

    aspect = img.size[1] / img.size[0]
    height = int(args.width * aspect)
    img = img.resize((args.width, height), Image.Resampling.NEAREST)
    print(f"[-] 图纸尺寸: {args.width} x {height}")

    # Excel 生成
    out_path = args.output if args.output else "pixel_pattern.xlsx"
    wb = xlsxwriter.Workbook(out_path)

    # Sheet 1: 图纸
    ws = wb.add_worksheet("图纸")
    ws.hide_gridlines(2)

    # 参数修正：宽度 3.8, 高度 26
    ws.set_column(0, args.width - 1, 3.8)
    for r in range(height):
        ws.set_row(r, 26)

    format_cache = {}
    usage_stats = {k: 0 for k in palette_info.keys()}
    pixel_data = img.load()

    print("[-] 正在生成...")
    for y in range(height):
        for x in range(args.width):
            c_id = get_closest_id(pixel_data[x, y], palette_list)
            usage_stats[c_id] += 1
            info = palette_info[c_id]

            if c_id not in format_cache:
                format_cache[c_id] = wb.add_format({
                    'bg_color': info['hex'],
                    'font_color': get_text_color(info['rgb']),
                    'align': 'center',
                    'valign': 'vcenter',
                    'font_size': 9,
                    'font_name': 'Arial',
                    'border': 1,
                    'border_color': '#C0C0C0'
                })
            ws.write_string(y, x, info['name'], format_cache[c_id])

    # Sheet 2: 采购清单
    ws_bom = wb.add_worksheet("采购清单")
    headers = ["色号", "系列", "预览", "数量"]
    for i, h in enumerate(headers):
        ws_bom.write(0, i, h, wb.add_format({'bold': True, 'bg_color': '#EEE', 'border': 1}))
        ws_bom.set_column(i, i, 15)

    row = 1
    for c_id, count in sorted(usage_stats.items(), key=lambda x: x[1], reverse=True):
        if count == 0: continue
        info = palette_info[c_id]
        ws_bom.write(row, 0, info['name'], wb.add_format({'align': 'center'}))
        ws_bom.write(row, 1, info['title'])
        ws_bom.write(row, 2, "", wb.add_format({'bg_color': info['hex'], 'border': 1}))
        ws_bom.write(row, 3, count)
        row += 1

    wb.close()
    print(f"[√] 完成！已保存至 {out_path}")


if __name__ == "__main__":
    main()