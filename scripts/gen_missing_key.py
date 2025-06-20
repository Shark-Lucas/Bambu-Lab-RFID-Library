#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
生成缺失的key.bin文件脚本
遍历所有目录，找到有dump.json但没有key.bin的地方，生成key.bin文件
"""

import os
import json
from pathlib import Path
import traceback

def find_missing_key_files():
    """查找需要生成key.bin文件的目录"""
    missing_key_dirs = []
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
            # 检查是否有dump.json文件
            json_files = [f for f in files if f.endswith('dump.json')]
            # 检查是否有key.bin文件
            key_files = [f for f in files if f.endswith('key.bin')]
            
            if json_files and not key_files:
                # 有json文件但没有key文件
                json_file = json_files[0]  # 取第一个json文件
                missing_key_dirs.append((Path(root), json_file))
    
    return missing_key_dirs

def generate_key_file(directory, json_filename):
    """从JSON文件生成key.bin文件"""
    try:
        json_path = directory / json_filename
        
        # 读取JSON文件
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # 获取UID
        uid = data['Card']['UID']
        
        # 获取扇区密钥
        sector_keys = data['SectorKeys']
        
        # 收集密钥字节：先收集KeyA（0-15扇区），然后收集KeyB（0-15扇区）
        key_bytes = bytearray()
        
        # KeyA 前半部分
        for sector in range(16):
            sec_str = str(sector)
            keyA_hex = sector_keys.get(sec_str, {}).get('KeyA', 'FFFFFFFFFFFF')
            key_bytes += bytes.fromhex(keyA_hex)
        
        # KeyB 后半部分
        for sector in range(16):
            sec_str = str(sector)
            keyB_hex = sector_keys.get(sec_str, {}).get('KeyB', 'FFFFFFFFFFFF')
            key_bytes += bytes.fromhex(keyB_hex)
        
        # 生成输出文件路径
        output_path = directory / f'hf-mf-{uid}-key.bin'
        
        # 写入key.bin文件
        with open(output_path, 'wb') as f:
            f.write(key_bytes)
        
        return True, len(key_bytes), str(output_path)
        
    except Exception as e:
        return False, str(e), None

def main():
    print("Bambu Lab RFID密钥文件生成工具")
    print("=" * 50)
    
    # 查找需要生成key.bin文件的目录
    print("正在搜索需要生成key.bin文件的目录...")
    missing_key_dirs = find_missing_key_files()
    
    if not missing_key_dirs:
        print("没有找到需要生成key.bin文件的目录。")
        return
    
    print(f"找到 {len(missing_key_dirs)} 个需要生成key.bin文件的目录")
    
    # 统计信息
    success_count = 0
    failed_count = 0
    failed_dirs = []
    
    # 为每个目录生成key.bin文件
    print("\n正在生成key.bin文件...")
    for i, (directory, json_filename) in enumerate(missing_key_dirs, 1):
        print(f"进度: {i}/{len(missing_key_dirs)} - {directory}")
        
        success, result, output_path = generate_key_file(directory, json_filename)
        
        if success:
            print(f"  ✓ 成功生成: {output_path} ({result} 字节)")
            success_count += 1
        else:
            print(f"  ✗ 生成失败: {result}")
            failed_count += 1
            failed_dirs.append(str(directory))
    
    # 显示最终统计
    print(f"\n生成完成！")
    print(f"总共处理目录: {len(missing_key_dirs)}")
    print(f"成功生成: {success_count}")
    print(f"生成失败: {failed_count}")
    
    if failed_dirs:
        print(f"\n失败的目录列表:")
        for failed_dir in failed_dirs:
            print(f"  - {failed_dir}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n用户中断操作")
    except Exception as e:
        print(f"\n发生错误: {e}")
        traceback.print_exc() 