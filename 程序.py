#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
宿舍信息自动填写
功能：将Excel中的学生住宿信息按宿舍分组，并根据Word模板生成每个宿舍的文档
备注：此程序使用ai辅助编写，功能暂不完整，仅用于自动填写宿舍信息
"""

import os
import pandas as pd
from docx import Document

def main():
    print("=== 宿舍信息自动填写 ===")
    
    # 设置当前工作目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)
    print(f"工作目录: {current_dir}")
    
    # 查找文件
    files = os.listdir('.')
    excel_files = [f for f in files if f.endswith('.xlsx') and not f.startswith('~$')]
    doc_files = [f for f in files if f.endswith('.docx') and not f.startswith('~$') and "信息表" in f]
    
    # 如果没有找到包含"信息表"的模板，使用所有docx文件
    if not doc_files:
        doc_files = [f for f in files if f.endswith('.docx') and not f.startswith('~$')]
    
    if not excel_files:
        print("错误: 未找到Excel文件 (.xlsx)")
        input("按回车键退出...")
        return
    
    if not doc_files:
        print("错误: 未找到Word模板文件 (.docx)")
        input("按回车键退出...")
        return
    
    # 显示并让用户选择Excel文件
    print("\n可用的Excel文件:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file}")
    
    excel_choice = input("请选择要使用的Excel文件 (默认为1): ").strip()
    try:
        excel_idx = int(excel_choice) - 1 if excel_choice else 0
        excel_file = excel_files[excel_idx]
    except (ValueError, IndexError):
        print("选择无效，使用第一个文件")
        excel_file = excel_files[0]
    
    # 显示并让用户选择Word模板
    print("\n可用的Word模板:")
    for i, file in enumerate(doc_files, 1):
        print(f"{i}. {file}")
    
    doc_choice = input("请选择要使用的Word模板 (默认为1): ").strip()
    try:
        doc_idx = int(doc_choice) - 1 if doc_choice else 0
        doc_file = doc_files[doc_idx]
    except (ValueError, IndexError):
        print("选择无效，使用第一个文件")
        doc_file = doc_files[0]
    
    print(f"\n使用Excel文件: {excel_file}")
    print(f"使用Word模板: {doc_file}")
    
    try:
        # 读取Excel数据
        print("\n正在读取Excel数据...")
        df = pd.read_excel(excel_file)
        print(f"成功读取 {len(df)} 行数据")
        
        # 显示列名
        print("\nExcel文件包含以下列:")
        for i, col in enumerate(df.columns, 1):
            print(f"{i}. {col}")
        
        # 获取用户输入
        print("\n请根据上面的列名输入对应的列名:")
        
        building_col = input("宿舍楼名称所在的列名: ").strip()
        while building_col not in df.columns:
            print(f"'{{building_col}}' 不是有效列名")
            building_col = input("宿舍楼名称所在的列名: ").strip()
        
        floor_col = input("楼层名称所在的列名: ").strip()
        while floor_col not in df.columns:
            print(f"'{{floor_col}}' 不是有效列名")
            floor_col = input("楼层名称所在的列名: ").strip()
        
        dorm_col = input("宿舍名称所在的列名: ").strip()
        while dorm_col not in df.columns:
            print(f"'{{dorm_col}}' 不是有效列名")
            dorm_col = input("宿舍名称所在的列名: ").strip()
        
        name_col = input("姓名所在的列名: ").strip()
        while name_col not in df.columns:
            print(f"'{{name_col}}' 不是有效列名")
            name_col = input("姓名所在的列名: ").strip()
        
        bed_col = input("床位号所在的列名: ").strip()
        while bed_col not in df.columns:
            print(f"'{{bed_col}}' 不是有效列名")
            bed_col = input("床位号所在的列名: ").strip()
        
        # 询问是否有职务列
        has_position = input("是否有职务列? (y/n): ").strip().lower()
        position_col = ""
        if has_position == 'y':
            position_col = input("职务所在的列名: ").strip()
            while position_col not in df.columns:
                print(f"'{{position_col}}' 不是有效列名")
                position_col = input("职务所在的列名: ").strip()
        
        # 按宿舍分组
        print("\n正在处理数据...")
        grouped = df.groupby([building_col, floor_col, dorm_col])
        
        success_count = 0
        
        for (building, floor, dorm), group in grouped:
            try:
                # 清理文件名
                filename = f"{clean_filename(str(building))}{clean_filename(str(floor))}{clean_filename(str(dorm))}.docx"
                
                # 处理Word文档
                process_word_template(doc_file, filename, group, name_col, position_col, bed_col)
                
                print(f"✓ 已生成: {filename}")
                success_count += 1
                
            except Exception as e:
                print(f"✗ 处理宿舍 {building}{floor}{dorm} 失败: {e}")
        
        print(f"\n处理完成! 成功生成 {success_count} 个宿舍文档")
        
    except Exception as e:
        print(f"程序执行出错: {e}")
        import traceback
        traceback.print_exc()
    
    input("按回车键退出...")

def clean_filename(name):
    """清理文件名中的非法字符"""
    invalid_chars = '<>"\\/|?*\n'
    for char in invalid_chars:
        name = name.replace(char, '')
    return name.strip()

def process_word_template(template_path, output_path, data, name_col, position_col, bed_col):
    """处理Word模板并生成新文档"""
    doc = Document(template_path)
    
    # 创建床位到姓名的映射
    bed_name_map = {}
    bed_position_map = {}
    
    for _, row in data.iterrows():
        bed = str(row[bed_col]).strip()
        if '.' in bed:
            bed = bed.split('.')[0]
        
        name = str(row[name_col]).strip()
        bed_name_map[bed] = name
        
        if position_col:
            position = str(row[position_col]).strip() if pd.notna(row[position_col]) else ""
            bed_position_map[bed] = position
    
    # 处理所有表格，查找包含序号、姓名、备注的表格
    for table in doc.tables:
        # 查找序号、姓名、备注行
        seq_row = name_row = remark_row = None
        
        for row_idx, row in enumerate(table.rows):
            if row.cells and row.cells[0]:
                first_cell = row.cells[0].text.strip()
                if first_cell == "序号":
                    seq_row = row_idx
                elif first_cell == "姓名":
                    name_row = row_idx
                elif first_cell == "备注":
                    remark_row = row_idx
        
        # 如果找到所有需要的行
        if seq_row is not None and name_row is not None:
            seq_cells = table.rows[seq_row].cells
            name_cells = table.rows[name_row].cells
            
            # 遍历序号行的单元格
            for cell_idx, cell in enumerate(seq_cells):
                if cell_idx == 0:  # 跳过表头单元格
                    continue
                    
                seq_num = cell.text.strip()
                if seq_num.isdigit() and seq_num in bed_name_map:
                    # 替换姓名 - 直接设置文本，不调用格式设置函数
                    if cell_idx < len(name_cells):
                        name_cell = name_cells[cell_idx]
                        name_cell.text = bed_name_map[seq_num]
                        
                    # 替换备注（如果有）
                    if remark_row is not None and cell_idx < len(table.rows[remark_row].cells):
                        remark_cell = table.rows[remark_row].cells[cell_idx]
                        if position_col and seq_num in bed_position_map:
                            remark_cell.text = bed_position_map[seq_num]
    
    doc.save(output_path)

if __name__ == "__main__":
    main()