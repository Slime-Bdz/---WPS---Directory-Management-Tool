import os
import sys
import shutil
from pathlib import Path

def resource_path(rel):
    """
    获取应用程序资源的绝对路径，兼容 PyInstaller 打包。
    """
    if hasattr(sys, '_MEIPASS'): #
        return os.path.join(sys._MEIPASS, rel) #
    return os.path.join(os.path.abspath('.'), rel) #

def ensure_embedded_excels():
    """
    确保 file_list.xlsx 和 file_list_updated.xlsx 存在于当前工作目录。
    如果不存在，则从资源路径复制过来。
    """
    current_dir = Path('.') #
    
    excels = ['file_list.xlsx', 'file_list_updated.xlsx'] #
    paths = {}
    for name in excels: #
        dst = current_dir / name
        if not dst.exists(): #
            src = resource_path(name) #
            shutil.copy2(src, dst) #
        paths[name] = str(dst) #
    return paths['file_list.xlsx'], paths['file_list_updated.xlsx'] #

# 保存和加载设置的函数也可以放在这里，或者更复杂的配置管理模块
# 为了保持与 UniApp 类的耦合性，这里暂时不移动 SETTINGS_FILE 相关的逻辑。
# 但如果以后有多个地方需要访问设置，就应该放在这里。