"""
utils.py

该模块包含一些通用的辅助函数，用于处理资源路径、文件操作等，
以支持应用程序在开发和打包后的环境中正常运行。
"""
import os
import sys
import shutil
from pathlib import Path
import pandas as pd  # 新增导入 pandas
from openpyxl import Workbook  # 新增导入 openpyxl

def resource_path(relative_path):
    """
    获取应用程序资源的绝对路径，兼容 PyInstaller 打包。
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath('.'), relative_path)

def setup_excel_files():
    """
    检查并创建程序所需的 resources 文件夹和默认 Excel 文件。
    如果文件不存在，将创建包含默认表头的空文件。
    """
    resources_dir = resource_path('resources')
    file_list_path = os.path.join(resources_dir, 'file_list.xlsx')
    file_list_updated_path = os.path.join(resources_dir, 'file_list_updated.xlsx')

    # 1. 检查并创建 resources 文件夹
    if not os.path.exists(resources_dir):
        os.makedirs(resources_dir)
        print(f"Created resources directory at: {resources_dir}")

    # 2. 检查并创建 file_list.xlsx
    if not os.path.exists(file_list_path):
        print(f"File not found: {file_list_path}. Creating new file...")
        df = pd.DataFrame(columns=['文件名'])
        df.to_excel(file_list_path, index=False)
        print("Created default file_list.xlsx with a header.")

    # 3. 检查并创建 file_list_updated.xlsx
    if not os.path.exists(file_list_updated_path):
        print(f"File not found: {file_list_updated_path}. Creating new file...")
        df = pd.DataFrame(columns=['文件名'])
        df.to_excel(file_list_updated_path, index=False)
        print("Created default file_list_updated.xlsx with a header.")
    
    return file_list_path, file_list_updated_path

# ----------------------------------------------------------------------
# 注意：原有的 ensure_embedded_excels 函数逻辑已不再需要，
# 因为 setup_excel_files 函数提供了更健壮的解决方案。
# 如果你的其他代码仍在调用 ensure_embedded_excels，请改为调用 setup_excel_files。
# ----------------------------------------------------------------------
def ensure_embedded_excels():
    """
    为了向后兼容，但其功能已被 setup_excel_files 替代。
    """
    return setup_excel_files()