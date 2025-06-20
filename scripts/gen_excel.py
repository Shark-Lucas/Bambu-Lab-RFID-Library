#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Bambu Lab RFID数据库统计脚本
遍历所有目录，解析RFID标签数据，生成Excel报告
"""

import os
import sys
import pandas as pd
from pathlib import Path
import importlib.util
import traceback
from collections import defaultdict

def load_parser_module():
    """加载parse.py模块"""
    try:
        # 脚本在scripts目录下运行，parse.py在同一目录下
        spec = importlib.util.spec_from_file_location("parse", "parse.py")
        parse_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(parse_module)
        return parse_module
    except Exception as e:
        print(f"错误：无法加载parse.py模块: {e}")
        return None

def translate_color(english_color):
    """将英文颜色翻译为中文"""
    color_map = {
        'Black': '黑色',
        'White': '白色',
        'Red': '红色',
        'Blue': '蓝色',
        'Green': '绿色',
        'Yellow': '黄色',
        'Orange': '橙色',
        'Purple': '紫色',
        'Pink': '粉色',
        'Gray': '灰色',
        'Grey': '灰色',
        'Silver': '银色',
        'Gold': '金色',
        'Bronze': '青铜色',
        'Cyan': '青色',
        'Magenta': '洋红色',
        'Beige': '米色',
        'Brown': '棕色',
        'Transparent': '透明',
        'Clear': '透明',
        'Translucent': '半透明',
        'Hot Pink': '热粉色',
        'Dark Gray': '深灰色',
        'Dark Grey': '深灰色',
        'Light Gray': '浅灰色',
        'Light Grey': '浅灰色',
        'Blue Gray': '蓝灰色',
        'Blue Grey': '蓝灰色',
        'Bambu Green': '拓竹绿',
        'Apple Green': '苹果绿',
        'Grass Green': '草绿色',
        'Forest Green': '森林绿',
        'Lime Green': '柠檬绿',
        'Lake Blue': '湖蓝色',
        'Ice Blue': '冰蓝色',
        'Sky Blue': '天蓝色',
        'Marine Blue': '海蓝色',
        'Royal Purple': '皇室紫',
        'Lilac Purple': '丁香紫',
        'Scarlet Red': '猩红色',
        'Lemon Yellow': '柠檬黄',
        'Ivory White': '象牙白',
        'Jade White': '玉白色',
        'Cream': '奶油色',
        'Desert Tan': '沙漠棕',
        'Peanut Brown': '花生棕',
        'Dark Brown': '深棕色',
        'Terracotta': '赤陶色',
        'Charcoal': '炭黑色',
        'Ash Gray': '灰白色',
        'Nardo Gray': '纳多灰',
        'Nebulae': '星云色',
        'Blue Hawaii (Blue-Green)': '夏威夷蓝（蓝绿色）',
        'Gilded Rose (Pink-Gold)': '镀金玫瑰（粉金色）',
        'Midnight Blaze (Blue-Red)': '午夜烈焰（蓝红色）',
        'Neon City (Blue-Magenta)': '霓虹城市（蓝洋红色）'
    }
    
    return color_map.get(english_color, english_color)

def extract_path_info(file_path):
    """从文件路径中提取材料、材料详细信息、颜色信息"""
    path_parts = Path(file_path).parts
    
    # 获取相对路径部分（去掉'..'开头）
    relevant_parts = [part for part in path_parts if part not in ['.', '..']]
    
    if len(relevant_parts) >= 3:
        material_category = relevant_parts[0]  # 例如：PLA, ABS, PETG
        material_type = relevant_parts[1]      # 例如：PLA Basic, PLA Matte
        color = relevant_parts[2]              # 例如：Black, White, Red
        uid = relevant_parts[3] if len(relevant_parts) > 3 else "未知"
        
        # 翻译颜色
        color_chinese = translate_color(color)
        color_display = f"{color_chinese}/{color}" if color_chinese != color else color
        
        return material_category, material_type, color, uid, color_display
    
    return "未知", "未知", "未知", "未知", "未知"

def find_rfid_files():
    """查找所有RFID数据文件（只查找dump.bin文件）"""
    rfid_files = []
    base_path = Path("..")  # 回到项目根目录
    
    # 遍历所有目录
    for root, dirs, files in os.walk(base_path):
        # 跳过根目录下的直接文件
        if root == str(base_path):
            # 从dirs中移除scripts目录，这样os.walk就不会进入该目录
            if 'scripts' in dirs:
                dirs.remove('scripts')
            continue
        
        # 跳过scripts目录及其子目录
        if 'scripts' in Path(root).parts:
            continue
            
        # 检查是否是UID目录（通常包含8位十六进制字符）
        current_dir = Path(root).name
        if len(current_dir) == 8 and all(c in '0123456789ABCDEFabcdef' for c in current_dir):
            # 在UID目录中查找dump.bin文件
            for file in files:
                if file.endswith('dump.bin'):
                    file_path = Path(root) / file
                    rfid_files.append(file_path)
    
    return rfid_files

def parse_rfid_file(file_path, parse_module):
    """解析单个RFID文件"""
    try:
        with open(file_path, 'rb') as f:
            data = f.read()
        
        tag = parse_module.Tag(str(file_path), data)
        return tag
    except Exception as e:
        print(f"解析文件 {file_path} 时出错: {e}")
        return None

def main():
    print("Bambu Lab RFID数据库统计工具")
    print("=" * 50)
    
    # 加载解析模块
    parse_module = load_parser_module()
    if not parse_module:
        return
    
    # 查找所有RFID文件
    print("正在搜索RFID文件...")
    rfid_files = find_rfid_files()
    print(f"找到 {len(rfid_files)} 个RFID文件")
    
    # 准备数据列表
    data_list = []
    stats = defaultdict(int)
    
    # 解析每个文件
    print("\n正在解析文件...")
    for i, file_path in enumerate(rfid_files, 1):
        print(f"进度: {i}/{len(rfid_files)} - {file_path}")
        
        # 从路径提取信息
        material_category, material_type, color, uid_from_path, color_display = extract_path_info(file_path)
        
        # 解析文件
        tag = parse_rfid_file(file_path, parse_module)
        
        if tag:
            # 获取温度信息
            min_temp = tag.data.get('temperatures', {}).get('min_hotend', {}).value if hasattr(tag.data.get('temperatures', {}).get('min_hotend', {}), 'value') else None
            max_temp = tag.data.get('temperatures', {}).get('max_hotend', {}).value if hasattr(tag.data.get('temperatures', {}).get('max_hotend', {}), 'value') else None
            
            # 格式化温度范围
            if min_temp is not None and max_temp is not None:
                temp_range = f"{min_temp}-{max_temp}°C"
            elif min_temp is not None:
                temp_range = f"{min_temp}°C"
            elif max_temp is not None:
                temp_range = f"{max_temp}°C"
            else:
                temp_range = '未知'
            
            # 提取数据（只保留需要的字段）
            data = {
                '材料类型': material_type,
                '颜色': color_display,
                'UID': tag.data.get('uid', '未知'),
                '颜色代码': tag.data.get('filament_color', '未知'),
                '打印温度': temp_range
            }
            
            data_list.append(data)
            
            # 统计信息
            stats[f'材料类型_{material_type}'] += 1
            stats[f'颜色_{color}'] += 1
            
        else:
            # 即使解析失败，也记录基本信息
            data = {
                '材料类型': material_type,
                '颜色': color_display,
                'UID': '解析失败',
                '颜色代码': '解析失败',
                '打印温度': '解析失败'
            }
            data_list.append(data)
    
    # 创建DataFrame
    df = pd.DataFrame(data_list)
    
    # 创建统计DataFrame
    stats_data = []
    for key, count in sorted(stats.items()):
        category, value = key.split('_', 1)
        stats_data.append({'类别': category, '值': value, '数量': count})
    
    stats_df = pd.DataFrame(stats_data)
    
    # 保存到Excel文件
    output_file = f'xlsx/Bambu_Lab_RFID_数据库报告.xlsx'
    print(f"\n正在生成Excel报告: {output_file}")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 主数据表
        df.to_excel(writer, sheet_name='RFID数据', index=False)
        
        # 统计数据表
        stats_df.to_excel(writer, sheet_name='统计汇总', index=False)
        
        # 材料分类汇总
        material_summary = df.groupby(['材料类型', '颜色']).size().reset_index(name='数量')
        material_summary.to_excel(writer, sheet_name='材料分类汇总', index=False)
        
        # 调整列宽
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    print(f"\n报告生成完成！")
    print(f"总共处理文件: {len(rfid_files)}")
    print(f"成功解析: {len([d for d in data_list if d['UID'] != '解析失败'])}")
    print(f"解析失败: {len([d for d in data_list if d['UID'] == '解析失败'])}")
    print(f"Excel文件已保存为: {output_file}")
    
    # 显示材料类型统计
    print(f"\n材料类型统计:")
    material_counts = df['材料类型'].value_counts()
    for material, count in material_counts.items():
        print(f"  {material}: {count}个")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户中断操作")
    except Exception as e:
        print(f"\n发生错误: {e}")
        traceback.print_exc() 