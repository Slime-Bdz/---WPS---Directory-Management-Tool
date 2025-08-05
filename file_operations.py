# file_operations.py
import os
import shutil
import traceback
from pathlib import Path
import concurrent.futures
from PyQt5.QtCore import QObject, pyqtSignal
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

class SearchWorker(QObject):
    finished = pyqtSignal()
    success = pyqtSignal(str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, int, str)
    
    _is_stopped = False

    def __init__(self, excel, target, roots, updated_excel_path):
        super().__init__()
        self.excel = excel
        self.target = target
        self.roots = roots
        self.updated_excel_path = updated_excel_path
        self._is_stopped = False

    def stop(self):
        self._is_stopped = True
        self.failed.emit("用户请求取消任务，正在停止...")

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

        except Exception as e:
            self.failed.emit(f"❌ 无法读取 Excel 文件: {self.excel} - {e}")
            return

        if not names_to_find:
            self.progress.emit(100, 100, "⚠️ Excel 中未找到文件名。")
            return

        # ==================== 1. 搜索阶段 (0% - 70%) ====================
        self.success.emit(f"🔎 开始在 {len(self.roots)} 个目录中查找 {len(names_to_find)} 个文件...")
        
        # 预先计算总共需要遍历的目录数量，以便实现平滑的进度条
        total_dirs_to_scan = 0
        for root in self.roots:
            if os.path.exists(root):
                total_dirs_to_scan += sum(1 for _ in os.walk(root))

        if total_dirs_to_scan == 0:
            self.progress.emit(70, 100, "⚠️ 搜索根目录为空或不存在。")
        
        dirs_scanned_count = 0
        found_files = {}

        for root in self.roots:
            if self._is_stopped:
                self.failed.emit("任务已中断。")
                return

            if not os.path.exists(root):
                self.failed.emit(f"❌ 查找根目录不存在: {root}")
                dirs_scanned_count += 1
                continue

            for dirpath, dirnames, filenames in os.walk(root):
                if self._is_stopped:
                    self.failed.emit("任务已中断。")
                    return
                
                dirs_scanned_count += 1
                
                # 计算搜索阶段的进度 (0% - 70%)
                if total_dirs_to_scan > 0:
                    search_progress_value = int((dirs_scanned_count / total_dirs_to_scan) * 70)
                    self.progress.emit(search_progress_value, 100, f"🔎 正在扫描: {dirpath}")
                else:
                    # 如果根目录为空，进度条直接到70%
                    self.progress.emit(70, 100, "🔎 搜索完成")
                
                for filename in filenames:
                    if len(found_files) == len(names_to_find):
                        # 所有文件已找到，提前退出当前目录的遍历
                        break
                    
                    for name_to_find in names_to_find:
                        if name_to_find in filename and name_to_find not in found_files:
                            found_files[name_to_find] = str(Path(dirpath) / filename)
                            self.success.emit(f"🔍 找到文件: {name_to_find}")
                            break
            
            if len(found_files) == len(names_to_find):
                self.success.emit("✅ 已找到所有文件，提前结束搜索。")
                break
        
        # 确保搜索阶段的进度条达到70%
        self.progress.emit(70, 100, "✅ 搜索阶段完成，准备复制文件...")
        
        if self._is_stopped:
            self.failed.emit("任务已中断。")
            return
            
        # ==================== 2. 复制阶段 (70% - 100%) ====================
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
                    
                    # 复制阶段的进度，从70%到100%
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