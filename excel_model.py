# excel_model.py
import os
import pandas as pd
from openpyxl import Workbook
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt, pyqtSignal
from PyQt5.QtWidgets import QTableView, QApplication, QAbstractItemView
from PyQt5.QtGui import QKeySequence
import traceback 

class ExcelTableModel(QAbstractTableModel):
    dataChangedSignal = pyqtSignal(int, int)

    def __init__(self, parent=None, is_read_only=False):
        super().__init__(parent)
        self.headers = []
        self.data_list = []
        self.is_read_only = is_read_only

    def load(self, path):
        self.beginResetModel()
        if not os.path.exists(path):
            wb = Workbook()
            ws = wb.active
            ws.append(["文件名"])
            wb.save(path)
        try:
            df = pd.read_excel(path, dtype=str).fillna("")
            self.headers = df.columns.tolist()
            self.data_list = df.values.tolist()
        except Exception as e:
            traceback.print_exc()
            self.headers = ["文件名"]
            self.data_list = []
        finally:
            self.endResetModel()

    def save(self, path):
        df = pd.DataFrame(self.data_list, columns=self.headers or ["文件名"])
        df.to_excel(path, index=False)
        
    def append_row(self):
        self.beginInsertRows(QModelIndex(), self.rowCount(), self.rowCount())
        new_row = [''] * self.columnCount()
        self.data_list.append(new_row)
        self.endInsertRows()

    def rowCount(self, parent=QModelIndex()):
        return len(self.data_list)

    def columnCount(self, parent=QModelIndex()):
        return len(self.headers) if self.headers else 1

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == Qt.DisplayRole or role == Qt.EditRole:
            if index.row() < len(self.data_list) and \
               index.column() < len(self.data_list[index.row()]):
                return str(self.data_list[index.row()][index.column()])
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid() or self.is_read_only:
            return False
        
        if role == Qt.EditRole:
            try:
                # 1. 确保目标行存在
                while index.row() >= len(self.data_list):
                    self.beginInsertRows(QModelIndex(), len(self.data_list), len(self.data_list))
                    initial_cols_for_new_row = self.columnCount()
                    self.data_list.append([''] * initial_cols_for_new_row)
                    self.endInsertRows()

                # 2. 确保表头有足够的列，并为所有现有行扩展这些列
                while index.column() >= len(self.headers):
                    new_col_name = f"Column{len(self.headers) + 1}"
                    self.headers.append(new_col_name)
                    for r_idx in range(len(self.data_list)):
                        if r_idx < len(self.data_list) and len(self.data_list[r_idx]) < len(self.headers):
                            self.data_list[r_idx].append('')
                    self.headerDataChanged.emit(Qt.Horizontal, len(self.headers) - 1, len(self.headers) - 1)
                
                # 3. 确保目标行（data_list[index.row()]）的内部列表有足够的列
                while index.row() < len(self.data_list) and index.column() >= len(self.data_list[index.row()]):
                    self.data_list[index.row()].append('')

                self.data_list[index.row()][index.column()] = str(value)
                self.dataChanged.emit(index, index, [Qt.DisplayRole, Qt.EditRole])
                return True
            except Exception as e:
                traceback.print_exc()
                return False
        return False

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                if section < len(self.headers):
                    return self.headers[section]
                return f"Column{section + 1}"
            if orientation == Qt.Vertical:
                return f"{section + 1}"
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        
        if self.is_read_only:
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable
        else:
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

class CustomTableView(QTableView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setTabKeyNavigation(True)

    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            self.copySelected()
        elif event.matches(QKeySequence.Paste):
            self.pasteFromClipboard()
        elif event.key() == Qt.Key_Delete or event.key() == Qt.Key_Backspace:
            self.deleteSelected()
        else:
            super().keyPressEvent(event)

    def copySelected(self):
        selection = self.selectionModel().selectedIndexes()
        if not selection:
            return

        min_row = min(index.row() for index in selection)
        max_row = max(index.row() for index in selection)
        min_col = min(index.column() for index in selection)
        max_col = max(index.column() for index in selection)

        table_data = []
        for r in range(min_row, max_row + 1):
            row_data = []
            for c in range(min_col, max_col + 1):
                index = self.model().index(r, c)
                if index in selection:
                    row_data.append(str(self.model().data(index, Qt.DisplayRole)))
                else:
                    row_data.append('')
            table_data.append(row_data)

        text_to_copy = '\n'.join(['\t'.join(row) for row in table_data])
        QApplication.clipboard().setText(text_to_copy)

    def deleteSelected(self):
        selection = self.selectionModel().selectedIndexes()
        if not selection:
            return
        
        if self.model().is_read_only:
            return

        min_row = min(index.row() for index in selection)
        max_row = max(index.row() for index in selection)
        min_col = min(index.column() for index in selection)
        max_col = max(index.column() for index in selection)
        
        model = self.model()
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                index = model.index(r, c)
                if index.isValid():
                    model.setData(index, '', Qt.EditRole)

    def pasteFromClipboard(self):
        if self.model().is_read_only:
            return

        clipboard = QApplication.clipboard()
        text = clipboard.text()

        if not text:
            return
        
        if text.endswith('\n'):
            text = text[:-1]

        rows_data = [line.split('\t') for line in text.split('\n')]
        pasted_rows_count = len(rows_data)

        model = self.model()
        selection_model = self.selectionModel()
        selected_indexes = selection_model.selectedIndexes()

        start_row = 0
        start_col = 0
        selected_rows_count = 0
        clear_min_col = 0
        clear_max_col = 0

        if not selected_indexes:
            current_index = self.currentIndex()
            if current_index.isValid():
                start_row = current_index.row()
                start_col = current_index.column()
        else:
            min_row = min(index.row() for index in selected_indexes)
            max_row = max(index.row() for index in selected_indexes)
            min_col = min(index.column() for index in selected_indexes)
            max_col = max(index.column() for index in selected_indexes)

            start_row = min_row
            start_col = min_col
            selected_rows_count = max_row - min_row + 1
            
            clear_min_col = min_col
            clear_max_col = max_col

        for r_offset, row_values in enumerate(rows_data):
            target_row = start_row + r_offset
            for c_offset, value in enumerate(row_values):
                target_col = start_col + c_offset
                target_index = model.index(target_row, target_col)
                if target_index.isValid():
                    model.setData(target_index, value, Qt.EditRole)

        if selected_rows_count > 0 and pasted_rows_count < selected_rows_count:
            rows_to_clear_start = start_row + pasted_rows_count
            rows_to_clear_end = start_row + selected_rows_count - 1

            actual_clear_end = min(rows_to_clear_end, model.rowCount() - 1)

            for r in range(rows_to_clear_start, actual_clear_end + 1):
                for c in range(clear_min_col, clear_max_col + 1):
                    index_to_clear = model.index(r, c)
                    if index_to_clear.isValid():
                        model.setData(index_to_clear, '', Qt.EditRole)