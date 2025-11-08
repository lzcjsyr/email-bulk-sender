#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
证书文件名解析工具
解析捐赠证书文件名中的姓名和邮箱，保存为Excel文件供确认
"""

import os
import re
import pandas as pd
from pathlib import Path
import sys

def parse_filename(filename):
    """
    从文件名中解析姓名和邮箱
    格式：姓名邮箱@domain.com.jpg 或 姓名_邮箱@domain.com.jpg
    """
    # 移除文件扩展名
    name_without_ext = os.path.splitext(filename)[0]
    
    # 尝试匹配中文姓名 + 邮箱格式
    # 模式1：小君xiaojun.zeng@buy42.com
    pattern1 = r'([\u4e00-\u9fa5]+)([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    match1 = re.search(pattern1, name_without_ext)
    
    if match1:
        name = match1.group(1).strip()
        email = match1.group(2).strip()
        return name, email
    
    # 模式2：小君_xiaojun.zeng@buy42.com (用下划线分隔)
    pattern2 = r'([\u4e00-\u9fa5]+)[-_]([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    match2 = re.search(pattern2, name_without_ext)
    
    if match2:
        name = match2.group(1).strip()
        email = match2.group(2).strip()
        return name, email
    
    # 模式3：纯英文姓名 + 邮箱
    # JohnSmith_john@example.com 或 John.Smith_john@example.com
    pattern3 = r'([a-zA-Z\s\.]+)[-_]([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    match3 = re.search(pattern3, name_without_ext)
    
    if match3:
        name = match3.group(1).strip()
        email = match3.group(2).strip()
        return name, email
    
    # 模式4：只包含邮箱
    # xiaojun.zeng@buy42.com
    pattern4 = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
    match4 = re.search(pattern4, name_without_ext)
    
    if match4:
        email = match4.group(1).strip()
        # 从邮箱用户名部分提取姓名
        name_part = email.split('@')[0]
        name = name_part.replace('.', ' ').replace('_', ' ').title()
        return name, email
    
    return None, None

def scan_certificates_folder(folder_path):
    """
    扫描证书文件夹，解析所有文件名
    """
    certificates = []
    
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 {folder_path} 不存在")
        return certificates
    
    # 支持的图片格式
    supported_formats = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.pdf'}
    
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        
        # 只处理文件，跳过文件夹
        if not os.path.isfile(file_path):
            continue
        
        # 检查文件格式
        file_ext = os.path.splitext(filename)[1].lower()
        if file_ext not in supported_formats:
            continue
        
        # 解析文件名
        name, email = parse_filename(filename)
        
        if name and email:
            certificates.append({
                '文件名': filename,
                '姓名': name,
                '邮箱': email,
                '文件路径': file_path,
                '状态': '待确认'
            })
        else:
            certificates.append({
                '文件名': filename,
                '姓名': '解析失败',
                '邮箱': '解析失败',
                '文件路径': file_path,
                '状态': '需要手动处理'
            })
    
    return certificates

def main():
    """
    主函数
    """
    print("=== 证书文件名解析工具 ===")
    print("功能：解析证书文件名中的姓名和邮箱，保存为Excel文件")
    print()
    
    # 默认证书文件夹路径
    default_folder = "捐赠证书"
    
    # 获取证书文件夹路径
    folder_path = input(f"请输入证书文件夹路径（直接回车使用默认路径'{default_folder}'）：").strip()
    if not folder_path:
        folder_path = default_folder
    
    # 扫描证书文件
    print(f"正在扫描文件夹：{folder_path}")
    certificates = scan_certificates_folder(folder_path)
    
    if not certificates:
        print("未找到任何证书文件！")
        return
    
    print(f"找到 {len(certificates)} 个证书文件")
    
    # 显示解析结果
    print("\n=== 解析结果 ===")
    success_count = 0
    fail_count = 0
    
    for cert in certificates:
        if cert['状态'] == '待确认':
            print(f"✓ {cert['文件名']} -> {cert['姓名']} | {cert['邮箱']}")
            success_count += 1
        else:
            print(f"✗ {cert['文件名']} -> 解析失败")
            fail_count += 1
    
    print(f"\n解析成功：{success_count} 个")
    print(f"解析失败：{fail_count} 个")
    
    # 保存为Excel文件
    output_file = "证书信息确认表.xlsx"
    try:
        df = pd.DataFrame(certificates)
        df.to_excel(output_file, index=False)
        print(f"\n结果已保存到：{output_file}")
        print("请打开Excel文件确认信息是否正确")
        
        # 显示Excel文件统计信息
        print(f"\nExcel文件包含以下列：")
        for col in df.columns:
            print(f"  - {col}")
            
    except Exception as e:
        print(f"保存Excel文件时出错：{e}")
        return
    
    print("\n=== 下一步操作 ===")
    print("1. 打开 证书信息确认表.xlsx 文件")
    print("2. 检查并修正姓名和邮箱信息")
    print("3. 在'状态'列中标记要发送的证书（保持'待确认'状态）")
    print("4. 保存Excel文件")
    print("5. 运行第二步脚本：python send_certificates.py")

if __name__ == "__main__":
    main()