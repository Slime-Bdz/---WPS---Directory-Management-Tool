# ---Directory-Management-Tool
# 目录管理终端

一个强大的桌面工具，旨在简化批量文件查找与整理流程，助您高效管理大量文件。

## 功能特性

* **Excel 驱动**: 通过 Excel 列表进行批量查找与复制，告别手动操作。
* **智能匹配**: 支持文件名及“词干”匹配，提高查找成功率。
* **实时报告**: 即时查看成功/失败日志，任务完成后自动生成带标记的更新版 Excel 报告。
* **内置编辑**: 直接在界面中编辑 Excel 列表，支持复制、粘贴、删除单元格内容。

## 如何使用

1.  **下载**: 从 [Release 页面](https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME/releases) 下载最新版本的 `DirectoryManager.exe` 可执行文件。
    * **提示**: 首次使用时，请点击界面中的“创建/刷新 Excel 表”按钮，以确保 `file_list.xlsx` 和 `file_list_updated.xlsx` 文件存在于程序运行目录下。
2.  **准备 Excel 列表**: 确保你的 `file_list.xlsx` 文件位于 `DirectoryManager.exe` 同一目录下。此文件应包含你要查找的文件名列表（第一列）。
3.  **运行程序**: 双击 `DirectoryManager.exe`。
4.  **配置路径**:
    * **Excel 列表**: 默认会自动加载同目录下的 `file_list.xlsx`。
    * **目标文件夹**: 选择文件复制的目的地。
    * **查找根目录**: 选择程序开始查找文件的根目录。
5.  **开始执行**: 点击“开始执行”按钮。
6.  **查看报告**: 执行完成后，`file_list_updated.xlsx` 将在同一目录下生成，并高亮显示未找到的文件。

## 本地开发设置 (针对开发者)

如果你想在本地运行或开发此项目，请遵循以下步骤：

1.  **克隆仓库**:
    ```bash
    git clone [https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git)
    cd YOUR_REPOSITORY_NAME # 将 YOUR_REPOSITORY_NAME 替换为你的仓库名称
    ```
2.  **创建虚拟环境并安装依赖**:
    ```bash
    python -m venv venv
    # macOS/Linux:
    source venv/bin/activate
    # Windows:
    .\venv\Scripts\activate
    pip install -r requirements.txt
    ```
    * **注意**: 如果 `requirements.txt` 文件不存在，你可以通过运行 `pip freeze > requirements.txt` 来生成当前环境中安装的所有库。本项目主要依赖 `PyQt5`, `pandas`, `openpyxl`。
3.  **运行**:
    ```bash
    python ui_app.py
    ```
