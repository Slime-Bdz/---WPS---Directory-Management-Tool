# 目录管理终端

一个强大的桌面工具，旨在简化批量文件查找与整理流程，助您高效迁移指定的多个文件。

[返回主页](README.md) | English README

---

### 功能特性

* **多核加速**: 利用多核处理器进行并行目录扫描，并使用多线程进行并发文件复制，确保在处理大型目录时也能快速完成任务。

* **Excel 驱动**: 通过 Excel 列表进行批量查找与复制，告别手动操作。

* **智能匹配**: 支持**精确匹配 (包含)**、**模糊匹配 (85%)** 和**正则表达式**三种模式，提高查找成功率。

* **实时报告**: 即时查看成功/失败日志，任务完成后自动生成带有标记的更新版 Excel 报告。

* **内置编辑**: 直接在界面中编辑 Excel 列表，支持复制、粘贴、删除单元格内容。

---

### 如何使用

1.  **准备 Excel 列表**: 确保你的 `file_list.xlsx` 文件位于程序运行目录下。此文件应包含你要查找的文件名列表（第一列）。
    * **提示**: 首次使用时，请点击“创建/刷新 Excel 表”按钮，以确保 `file_list.xlsx` 和 `file_list_updated.xlsx` 文件存在。

2.  **运行程序**: 双击可执行文件。

3.  **配置路径**:
    * **Excel 列表**: 默认会自动加载同目录下的 `file_list.xlsx`。
    * **目标文件夹**: 选择文件复制的目的地。
    * **查找根目录**: 选择程序开始查找文件的根目录。

4.  **开始执行**: 点击“开始执行”按钮。

5.  **查看报告**: 执行完成后，`file_list_updated.xlsx` 将在同一目录下生成，并高亮显示未找到的文件。

---

### 本地开发设置 (针对开发者)

1.  **克隆仓库**:
    ```bash
    git clone [https://github.com/Slime-Bdz/Directory-Management-Tool.git](https://github.com/Slime-Bdz/Directory-Management-Tool.git)
    cd Directory-Management-Tool 
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
    * **注意**: 本项目主要依赖 `PyQt5`, `pandas`, `openpyxl`, `fuzzywuzzy`。

3.  **运行**:
    ```bash
    python main_app.py
    ```
