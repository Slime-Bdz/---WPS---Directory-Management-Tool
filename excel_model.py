"""
excel_model.py

该模块包含与Excel文件处理相关的模型类和表格视图类，
负责数据的加载、保存、显示和编辑功能。
"""
import sys
from pathlib import Path
from PyQt5.QtCore import QAbstractTableModel, Qt, QModelIndex, QVariant
from PyQt5.QtWidgets import QTableView, QApplication, QHeaderView, QAbstractItemView
from PyQt5.QtGui import QKeySequence
import pandas as pd


class ExcelTableModel(QAbstractTableModel):
    """
    Excel 表格数据模型，基于 QAbstractTableModel 实现。
    支持从 Excel 文件加载数据，并提供编辑和保存功能。
    """

    def __init__(self, is_read_only=False, parent=None):
        """
        初始化模型
        
        Args:
            is_read_only (bool): 是否为只读模式
            parent: 父对象
        """
        super().__init__(parent)
        self.df = pd.DataFrame()
        self.is_read_only = is_read_only

    def rowCount(self, parent=QModelIndex()):
        """返回行数"""
        return len(self.df)

    def columnCount(self, parent=QModelIndex()):
        """返回列数"""
        return len(self.df.columns) if not self.df.empty else 0

    def data(self, index, role=Qt.DisplayRole):
        """获取指定索引位置的数据"""
        if not index.isValid():
            return QVariant()

        if role == Qt.DisplayRole or role == Qt.EditRole:
            try:
                value = self.df.iloc[index.row(), index.column()]
                return str(value) if pd.notna(value) else ""
            except (IndexError, KeyError):
                return ""
        
        return QVariant()

    def setData(self, index, value, role=Qt.EditRole):
        """设置指定索引位置的数据"""
        if not index.isValid() or self.is_read_only:
            return False

        if role == Qt.EditRole:
            try:
                self.df.iloc[index.row(), index.column()] = value
                self.dataChanged.emit(index, index, [Qt.DisplayRole])
                return True
            except (IndexError, KeyError):
                return False
        
        return False

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """获取表头数据"""
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                if section < len(self.df.columns):
                    return str(self.df.columns[section])
            elif orientation == Qt.Vertical:
                return str(section + 1)
        return QVariant()

    def flags(self, index):
        """返回指定索引位置的标志"""
        if not index.isValid():
            return Qt.NoItemFlags
        
        flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if not self.is_read_only:
            flags |= Qt.ItemIsEditable
        return flags

    def insertRows(self, row, count, parent=QModelIndex()):
        """插入行"""
        if self.is_read_only:
            return False
            
        self.beginInsertRows(parent, row, row + count - 1)
        
        # 创建新的空行数据
        new_rows = pd.DataFrame([['' for _ in range(len(self.df.columns))] for _ in range(count)], 
                                columns=self.df.columns)
        
        # 分割原数据框并插入新行
        if row == 0:
            self.df = pd.concat([new_rows, self.df], ignore_index=True)
        elif row >= len(self.df):
            self.df = pd.concat([self.df, new_rows], ignore_index=True)
        else:
            top_part = self.df.iloc[:row]
            bottom_part = self.df.iloc[row:]
            self.df = pd.concat([top_part, new_rows, bottom_part], ignore_index=True)
        
        self.endInsertRows()
        return True

    def removeRows(self, row, count, parent=QModelIndex()):
        """删除行"""
        if self.is_read_only:
            return False
            
        if row < 0 or row + count > len(self.df):
            return False
            
        self.beginRemoveRows(parent, row, row + count - 1)
        self.df = self.df.drop(self.df.index[row:row + count]).reset_index(drop=True)
        self.endRemoveRows()
        return True

    def appendRow(self):
        """在末尾添加一行"""
        if self.is_read_only:
            return False
        return self.insertRows(len(self.df), 1)
    
    def cleanup_empty_rows(self):
        """
        [新增] 实时清理所有空行并重新排列数据
        """
        if self.df.empty:
            return

        initial_row_count = len(self.df)

        # 1. 将空字符串临时替换为 NaN，以便使用 dropna
        temp_df = self.df.replace('', pd.NA)
        
        # 2. 移除所有单元格都为 NaN 的行，并重置索引
        self.df = temp_df.dropna(how='all').reset_index(drop=True)
        
        # 3. 如果数据框变为空，创建一个默认行
        if self.df.empty:
            self.df = pd.DataFrame({'文件名': ['']})
        else:
            # 4. 再次填充 NaN 为空字符串，以保持数据一致性
            self.df = self.df.fillna('')

        final_row_count = len(self.df)

        if initial_row_count != final_row_count:
            # 如果行数发生变化，通知视图重置模型以更新显示
            self.beginResetModel()
            self.endResetModel()

    def load(self, excel_path):
        """从 Excel 文件加载数据"""
        try:
            path = Path(excel_path)
            if path.exists():
                self.beginResetModel()
                self.df = pd.read_excel(excel_path)
                
                # 确保至少有一列
                if self.df.empty or len(self.df.columns) == 0:
                    self.df = pd.DataFrame({'文件名': ['']})
                
                # 填充NaN值为空字符串
                self.df = self.df.fillna('')
                
                self.endResetModel()
            else:
                # 文件不存在时创建默认的DataFrame
                self.beginResetModel()
                self.df = pd.DataFrame({'文件名': ['']})
                self.endResetModel()
        except Exception as e:
            print(f"加载Excel文件失败: {e}")
            self.beginResetModel()
            self.df = pd.DataFrame({'文件名': ['']})
            self.endResetModel()

    def save(self, excel_path):
        """将数据保存到 Excel 文件"""
        try:
            # 在保存前也进行一次最终清理，确保保存结果干净
            self.cleanup_empty_rows()
            
            # 创建目录（如果不存在）
            path = Path(excel_path)
            path.parent.mkdir(parents=True, exist_ok=True)
            
            # 保存到Excel文件
            self.df.to_excel(excel_path, index=False)
            return True
        except Exception as e:
            print(f"保存Excel文件失败: {e}")
            return False


class CustomTableView(QTableView):
    """
    自定义表格视图，支持复制粘贴、删除等操作
    """

    def __init__(self, parent=None):
        """初始化表格视图"""
        super().__init__(parent)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
    def keyPressEvent(self, event):
        """处理键盘事件，实现复制粘贴、删除等功能"""
        selected_indexes = self.selectedIndexes()
        
        if not selected_indexes:
            super().keyPressEvent(event)
            return
            
        # 处理删除键 (Delete 和 Backspace)
        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            for index in selected_indexes:
                self.model().setData(index, "", Qt.EditRole)
            
            # [新增] 删除后立即进行空行清理
            self.model().cleanup_empty_rows()
                
        # 处理复制 (Ctrl+C)
        elif event.matches(QKeySequence.Copy):
            self._copy_selection()
            
        # 处理粘贴 (Ctrl+V)
        elif event.matches(QKeySequence.Paste):
            self._paste_selection()
            
        else:
            super().keyPressEvent(event)

    def _copy_selection(self):
        """复制选中的内容到剪贴板"""
        selected_indexes = self.selectedIndexes()
        if not selected_indexes:
            return
            
        # 按行列排序索引
        selected_indexes.sort(key=lambda x: (x.row(), x.column()))
        
        # 获取选择区域的边界
        min_row = min(idx.row() for idx in selected_indexes)
        max_row = max(idx.row() for idx in selected_indexes)
        min_col = min(idx.column() for idx in selected_indexes)
        max_col = max(idx.column() for idx in selected_indexes)
        
        # 构建复制的文本
        rows_data = []
        for row in range(min_row, max_row + 1):
            row_data = []
            for col in range(min_col, max_col + 1):
                index = self.model().index(row, col)
                data = self.model().data(index, Qt.DisplayRole) or ""
                row_data.append(str(data))
            rows_data.append("\t".join(row_data))
        
        clipboard_text = "\n".join(rows_data)
        QApplication.clipboard().setText(clipboard_text)

    def _paste_selection(self):
        """从剪贴板粘贴内容"""
        clipboard_text = QApplication.clipboard().text()
        if not clipboard_text:
            return
            
        selected_indexes = self.selectedIndexes()
        if not selected_indexes:
            return
            
        # 获取粘贴起始位置（选中区域的左上角）
        start_index = min(selected_indexes, key=lambda x: (x.row(), x.column()))
        start_row = start_index.row()
        start_col = start_index.column()
        
        # 解析剪贴板数据
        rows_data = clipboard_text.strip().split('\n')
        paste_rows = len(rows_data)
        paste_cols = max(len(row.split('\t')) for row in rows_data) if rows_data else 0
        
        # 检查是否需要添加新行
        model = self.model()
        current_rows = model.rowCount()
        required_rows = start_row + paste_rows
        
        if required_rows > current_rows:
            # 需要添加新行
            rows_to_add = required_rows - current_rows
            model.insertRows(current_rows, rows_to_add)
        
        # 检查是否需要添加新列（如果模型支持的话）
        current_cols = model.columnCount()
        required_cols = start_col + paste_cols
        
        # 粘贴数据
        for row_idx, row_data in enumerate(rows_data):
            cells = row_data.split('\t')
            for col_idx, cell_data in enumerate(cells):
                target_row = start_row + row_idx
                target_col = start_col + col_idx
                
                # 确保不超出列范围
                if target_col < current_cols:
                    index = model.index(target_row, target_col)
                    model.setData(index, cell_data, Qt.EditRole)

        # [新增] 粘贴后立即进行空行清理
        self.model().cleanup_empty_rows()


# 测试代码
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 创建测试窗口
    view = CustomTableView()
    model = ExcelTableModel(is_read_only=False)
    
    # 创建测试数据
    test_df = pd.DataFrame({
        '文件名': ['file1.txt', '', 'file3.txt'],
        '状态': ['待处理', '', '待处理'],
        '备注': ['', '', '']
    })
    model.df = test_df
    
    view.setModel(model)
    view.show()
    
    sys.exit(app.exec_())