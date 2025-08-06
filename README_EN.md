# Directory Management Tool

A powerful desktop tool designed to streamline bulk file search and organization, enabling you to efficiently manage and transfer multiple specified files.

[Back to main](README.md) | 简体中文 README

---

### Key Features

* **High Performance**: Leverages multi-core processing for parallel directory scanning and multi-threading for concurrent file copying, ensuring rapid task completion even with large directories.

* **Excel-Driven**: Use an Excel list to perform bulk searches and copies, eliminating tedious manual operations.

* **Intelligent Matching**: Supports **Exact (contains)**, **Fuzzy (85%)**, and **Regular Expression** matching modes to enhance search success rates.

* **Real-time Reporting**: Instantly view success/failure logs. Upon completion, an updated Excel report with highlighted statuses is automatically generated.

* **Built-in Editor**: Directly edit the Excel list within the interface, with support for copy, paste, and delete operations.

---

### How to Use

1.  **Prepare Excel List**: Ensure your `file_list.xlsx` file is in the same directory as the executable. This file should contain a list of filenames you want to find in the first column.
    * **Tip**: For first-time use, click "Create/Refresh Excel" to ensure `file_list.xlsx` and `file_list_updated.xlsx` exist.

2.  **Run the Program**: Double-click the executable.

3.  **Configure Paths**:
    * **Excel List**: The `file_list.xlsx` from the current directory is loaded by default.
    * **Target Folder**: Select the destination for copied files.
    * **Search Root Directory**: Select the root folder where the program should begin its search.

4.  **Start Execution**: Click the "Start" button.

5.  **View Report**: After execution, `file_list_updated.xlsx` will be generated in the same directory, highlighting files that were not found.

---

### For Developers

1.  **Clone the repository**:
    ```bash
    git clone [https://github.com/Slime-Bdz/Directory-Management-Tool.git](https://github.com/Slime-Bdz/Directory-Management-Tool.git)
    cd Directory-Management-Tool
    ```

2.  **Create a virtual environment and install dependencies**:
    ```bash
    python -m venv venv
    # macOS/Linux:
    source venv/bin/activate
    # Windows:
    .\venv\Scripts\activate
    pip install -r requirements.txt
    ```
    * **Note**: This project primarily depends on `PyQt5`, `pandas`, `openpyxl`, `fuzzywuzzy`.

3.  **Run the application**:
    ```bash
    python main_app.py
    ```
