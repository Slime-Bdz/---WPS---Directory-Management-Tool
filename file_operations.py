# file_operations.py
import os
import shutil
import traceback
import re
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from fuzzywuzzy import fuzz
import concurrent.futures
from PyQt5.QtCore import QObject, pyqtSignal

def _scan_root_process(root_dir, names_to_find_set, match_mode, min_fuzzy_score):
    """
    一个独立的、可被多进程调用的函数，用于在单个根目录中查找文件。
    该函数只接受可被序列化的参数。
    """
    found_files = {}
    
    if not os.path.exists(root_dir):
        return found_files

    # 预处理，提高查找效率
    if match_mode == 'regex':
        try:
            # 预编译所有正则表达式
            regex_patterns = {name: re.compile(name) for name in names_to_find_set}
        except re.error:
            # 如果正则表达式无效，返回空结果
            return found_files
    
    for dirpath, _, filenames in os.walk(root_dir):
        current_dir_files_set = set(filenames)

        if match_mode == 'exact':
            # 修正：将精确匹配改为子字符串匹配，符合“包含”的定义
            for name_to_find in names_to_find_set:
                if name_to_find not in found_files:
                    for filename in current_dir_files_set:
                        if name_to_find in filename:
                            found_files[name_to_find] = str(Path(dirpath) / filename)
                            break
        
        elif match_mode == 'fuzzy':
            # 模糊匹配需要逐一检查
            for filename in current_dir_files_set:
                for name_to_find in names_to_find_set:
                    if name_to_find not in found_files:
                        if fuzz.ratio(name_to_find, filename) >= min_fuzzy_score:
                            found_files[name_to_find] = str(Path(dirpath) / filename)
                            break
        
        elif match_mode == 'regex':
            for filename in current_dir_files_set:
                for name_to_find, pattern in regex_patterns.items():
                    if name_to_find not in found_files:
                        if pattern.search(filename):
                            found_files[name_to_find] = str(Path(dirpath) / filename)
                            break
        
        # 如果已经找到所有文件，则提前退出
        if len(found_files) == len(names_to_find_set):
            break
    
    return found_files

class SearchWorker(QObject):
    finished = pyqtSignal()
    success = pyqtSignal(str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, int, str)
    
    _is_stopped = False

    def __init__(self, excel, target, roots, updated_excel_path, match_mode='exact', min_fuzzy_score=85, parent=None):
        super().__init__(parent)
        self.excel = excel
        self.target = target
        self.roots = roots
        self.updated_excel_path = updated_excel_path
        self.match_mode = match_mode
        self.min_fuzzy_score = min_fuzzy_score
        self._is_stopped = False
        self.executor = None

    def stop(self):
        self._is_stopped = True
        self.failed.emit("用户请求取消任务，正在停止...")
        if self.executor:
            self.executor.shutdown(wait=False, cancel_futures=True)

    def run(self):
        try:
            self._work()
        except Exception:
            traceback.print_exc()
            self.failed.emit("任务执行出错，请检查日志。")
        finally:
            self.finished.emit()

    def _copy_single_file(self, name_to_find, src_path, target_dir):
        if self._is_stopped:
            return {'status': 'stopped', 'message': f"任务已中断。", 'name': name_to_find}
        
        if src_path and os.path.exists(src_path):
            dst_name = os.path.basename(src_path)
            dst = os.path.join(target_dir, dst_name)
            try:
                if os.path.isfile(src_path):
                    shutil.copy2(src_path, dst)
                else:
                    if os.path.exists(dst):
                        shutil.rmtree(dst)
                    shutil.copytree(src_path, dst)
                return {'status': 'success', 'message': f"✅ 已复制: {dst_name}", 'name': name_to_find}
            except Exception as e:
                return {'status': 'failed', 'message': f"❌ 复制失败 ({name_to_find}): {e}", 'name': name_to_find}
        else:
            return {'status': 'failed', 'message': f"❌ 未找到: {name_to_find}", 'name': name_to_find}
    
    def _find_files_in_roots(self, names_to_find_set):
        """并行扫描多个根目录并汇总结果。"""
        combined_found_files = {}
        total_roots = len(self.roots)
        completed_roots = 0
        
        with concurrent.futures.ProcessPoolExecutor(max_workers=os.cpu_count() or 4) as executor:
            self.executor = executor
            futures = {
                executor.submit(_scan_root_process, root, names_to_find_set, self.match_mode, self.min_fuzzy_score): root
                for root in self.roots
            }
            
            for future in concurrent.futures.as_completed(futures):
                if self._is_stopped:
                    for f in futures:
                        f.cancel()
                    break
                
                try:
                    result = future.result()
                    for name, path in result.items():
                        if name not in combined_found_files:
                            combined_found_files[name] = path
                            self.success.emit(f"🔍 找到文件: {name}")
                except Exception as e:
                    root_dir = futures[future]
                    self.failed.emit(f"❌ 扫描目录 {root_dir} 发生错误: {e}")
                
                completed_roots += 1
                search_progress_value = int((completed_roots / total_roots) * 70)
                self.progress.emit(search_progress_value, 100, f"🔎 正在扫描: {completed_roots}/{total_roots} 个目录")
                
                if len(combined_found_files) == len(names_to_find_set):
                    self.success.emit("✅ 已找到所有文件，提前结束搜索。")
                    for f in futures:
                        f.cancel()
                    break
        
        self.executor = None
        return combined_found_files

    def _work(self):
        self.progress.emit(0, 100, "⚙️ 正在初始化...")
        os.makedirs(self.target, exist_ok=True)
        
        try:
            wb_original = load_workbook(self.excel)
            ws_original = wb_original.active
            
            names_to_find = [
                str(row[0]).strip() 
                for row in ws_original.iter_rows(min_row=2, values_only=True) 
                if row and row[0] is not None
            ]
            names_to_find = [name for name in names_to_find if name]
            names_to_find_set = set(names_to_find)

        except Exception as e:
            self.failed.emit(f"❌ 无法读取 Excel 文件: {self.excel} - {e}")
            return

        if not names_to_find_set:
            self.progress.emit(100, 100, "⚠️ Excel 中未找到文件名。")
            return
        
        self.success.emit(f"🔎 开始在 {len(self.roots)} 个目录中查找 {len(names_to_find)} 个文件...")

        found_files = self._find_files_in_roots(names_to_find_set)
        
        self.progress.emit(70, 100, "✅ 搜索阶段完成，准备复制文件...")
        
        if self._is_stopped:
            self.failed.emit("任务已中断。")
            return
            
        total_files_to_process = len(names_to_find)
        copied_count = 0
        copy_results = []
        
        if total_files_to_process > 0:
            self.progress.emit(70, 100, "📁 正在并发复制文件...")
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=os.cpu_count() * 2 if os.cpu_count() else 4) as executor:
                futures = [executor.submit(self._copy_single_file, name, found_files.get(name), self.target)
                           for name in names_to_find]

                for future in concurrent.futures.as_completed(futures):
                    if self._is_stopped:
                        for f in futures:
                            f.cancel()
                        break
                    
                    copied_count += 1
                    try:
                        result = future.result()
                        copy_results.append(result)
                        if result['status'] == 'success':
                            self.success.emit(result['message'])
                        elif result['status'] == 'failed':
                            self.failed.emit(result['message'])
                        elif result['status'] == 'stopped':
                            self.failed.emit(result['message'])
                            for f in futures:
                                f.cancel()
                            break
                    except Exception as e:
                        self.failed.emit(f"❌ 任务处理异常: {e}")
                    
                    copy_progress_value = 70 + int((copied_count / total_files_to_process) * 30)
                    self.progress.emit(copy_progress_value, 100, f"🚀 正在复制文件: {copied_count}/{total_files_to_process}")
        else:
            self.success.emit("没有需要复制的文件。")
            self.progress.emit(100, 100, "任务完成。")
            
        if not self._is_stopped:
            self._finalize_excel_report(self.updated_excel_path, names_to_find, copy_results)
            self.progress.emit(100, 100, "任务完成。")
            self.success.emit(f'✅ 已保存更新表：{Path(self.updated_excel_path).name}')
        else:
            self.failed.emit("任务已中断。")

    def _finalize_excel_report(self, updated_excel_path, names_to_find, copy_results):
        try:
            if not os.path.exists(updated_excel_path):
                wb = Workbook()
                ws = wb.active
                ws.cell(row=1, column=1, value="文件名")
                ws.cell(row=1, column=2, value="状态")
                wb.save(updated_excel_path)
            
            wb = load_workbook(updated_excel_path)
            ws = wb.active

            results_map = {res['name']: res for res in copy_results}
            
            for i in range(len(names_to_find) - ws.max_row):
                ws.append(['', ''])

            for idx, name_to_find in enumerate(names_to_find):
                row_index = idx + 2
                result = results_map.get(name_to_find)
                
                if result:
                    status = result['status']
                    if status == 'success':
                        fill = PatternFill(fill_type='solid', start_color='00FF00', end_color='00FF00')
                        status_text = "✅ 已找到"
                    else:
                        fill = PatternFill(fill_type='solid', start_color='FFC0CB', end_color='FFC0CB')
                        status_text = "❌ 未找到或复制失败"
                else:
                    fill = PatternFill(fill_type='solid', start_color='FFFF00', end_color='FFFF00')
                    status_text = "❌ 未找到"

                ws.cell(row=row_index, column=1, value=name_to_find)
                ws.cell(row=row_index, column=2, value=status_text)
                ws.cell(row=row_index, column=2).fill = fill
                    
            wb.save(updated_excel_path)
        except Exception as e:
            self.failed.emit(f"❌ 无法保存更新的 Excel 报告: {e}")
            traceback.print_exc()