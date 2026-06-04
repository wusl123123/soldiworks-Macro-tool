# -*- coding: utf-8 -*-
"""
SW批量宏命令脚本
制作人: wsl
制作日期: 2026-06-03
用于在SolidWorks中对选中的文件批量执行宏程序
"""

import os
import sys
import time
import threading
import pythoncom
import ctypes
from ctypes import wintypes

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QListWidget, QListWidgetItem, QComboBox,
    QProgressBar, QTextEdit, QGroupBox, QMessageBox, QScrollArea, QSplitter,
    QSizePolicy
)
from PySide6.QtGui import QTextCursor
from PySide6.QtCore import Qt, QThread, Signal, QTimer

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False
    print("警告: 未安装pywin32库，请安装: pip install pywin32")

# SolidWorks文件类型常量
swDocNONE = 0
swDocPART = 1
swDocASSEMBLY = 2
swDocDRAWING = 3
swAllMethods = 3

class WorkerThread(QThread):
    """工作线程，用于执行批量宏命令（支持多宏串联执行）"""
    progress_update = Signal(int, str)
    task_complete = Signal(bool, str)
    log_message = Signal(str)
    file_status_update = Signal(str, str)  # 文件路径, 状态(completed/failed)

    def __init__(self, file_list, macro_configs, parent=None):
        super().__init__(parent)
        self.file_list = file_list
        self.macro_configs = macro_configs  # 宏配置列表，每个元素包含{'macro_path': ..., 'module_name': ...}
        self.is_running = True
        self.is_paused = False
        self.pause_event = threading.Event()
        self.pause_event.set()  # 初始状态为运行

    def run(self):
        """执行批量宏命令"""
        try:
            pythoncom.CoInitialize()
            
            # 连接到SolidWorks
            # 使用 Dispatch 而非 DispatchEx，避免创建新进程
            sw_app = win32com.client.Dispatch("SldWorks.Application")
            sw_app.Visible = True  # 必须可见才能正常操作
            
            # 设置 SW 选项，禁止弹出对话框
            # swUserPreferenceIntegerValue_e 中：
            # swAllowMultipleDocuments = 270 (允许多文档)
            try:
                sw_app.SetUserPreferenceIntegerValue(270, 1)
            except Exception:
                pass
            
            total_files = len(self.file_list)
            total_macros = len(self.macro_configs)
            success_count = 0
            error_messages = []
            
            for index, file_path in enumerate(self.file_list):
                if not self.is_running:
                    self.log_message.emit("用户取消操作")
                    break
                
                # 检查暂停状态
                self.pause_event.wait()
                if not self.is_running:
                    self.log_message.emit("用户取消操作")
                    break
                
                try:
                    # 获取文件类型
                    file_ext = os.path.splitext(file_path)[1].lower()
                    if file_ext == ".sldprt":
                        doc_type = swDocPART
                    elif file_ext == ".sldasm":
                        doc_type = swDocASSEMBLY
                    elif file_ext == ".slddrw":
                        doc_type = swDocDRAWING
                    else:
                        self.log_message.emit(f"跳过未知文件类型: {file_path}")
                        continue
                    
                    # 确保文件路径是字符串
                    file_path_str = str(file_path)
                    
                    # 使用 OpenDoc6（SW 2013+ 可用，参数完整，支持静默打开）
                    # OpenDoc6(FileName, Type, Options, Configuration, Errors, Warnings)
                    # Options: swOpenDocOptions_Silent = 1 (静默模式)
                    #          swOpenDocOptions_ReadOnly = 2 (只读)
                    #          swOpenDocOptions_RapidDraft = 4 (快速工程图)
                    # 组合使用 Silent + ReadOnly = 3
                    open_options = 1  # Silent 静默打开，不弹对话框
                    errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                    
                    self.log_message.emit(f"打开文件: {os.path.basename(file_path)}")
                    sw_doc = sw_app.OpenDoc6(
                        file_path_str,
                        doc_type,
                        open_options,
                        "",        # Configuration（空=默认配置）
                        errors,
                        warnings
                    )
                    
                    if sw_doc is None:
                        raise Exception("无法打开文档")
                    self.log_message.emit(f"打开完成: {os.path.basename(file_path)}")
                    
                    # 多宏串联执行：按顺序执行每个宏
                    for macro_index, macro_config in enumerate(self.macro_configs):
                        if not self.is_running:
                            self.log_message.emit("用户取消操作")
                            break
                        
                        # 检查暂停状态
                        self.pause_event.wait()
                        if not self.is_running:
                            self.log_message.emit("用户取消操作")
                            break
                        
                        try:
                            macro_path = macro_config['macro_path']
                            module_name = macro_config['module_name']
                            self.log_message.emit(f"执行宏[{macro_index + 1}/{total_macros}]: {os.path.basename(macro_path)} - {module_name}")
                            self.execute_macro(sw_app, macro_path, module_name)
                            self.log_message.emit(f"宏[{macro_index + 1}/{total_macros}]执行完成")
                        except Exception as macro_err:
                            raise Exception(f"宏[{macro_index + 1}/{total_macros}]执行失败: {str(macro_err)}")
                    
                    if not self.is_running:
                        break
                    
                    # 执行宏后保存文档（遵循VBA逻辑）
                    self.log_message.emit(f"保存文件: {os.path.basename(file_path)}")
                    try:
                        # 获取活动文档（对标VBA: Set Part = swApp.ActiveDoc）
                        active_doc = sw_app.ActiveDoc
                        if active_doc is None:
                            self.log_message.emit(f"警告: 无活动文档")
                        else:
                            # 获取文档标题（对标VBA: Title = Part.GetTitle）
                            # 兼容属性和方法两种访问方式
                            doc_title = active_doc.GetTitle
                            if callable(doc_title):
                                doc_title = doc_title()
                            
                            # 保存文档（对标VBA: Part.Save）
                            active_doc.Save()
                            self.log_message.emit(f"保存完成: {doc_title}")
                            
                            # 释放对象引用（对标VBA: Set Part = Nothing）
                            active_doc = None
                            
                            # 使用标题关闭文档（对标VBA: swApp.CloseDoc Title）
                            sw_app.CloseDoc(doc_title)
                            self.log_message.emit(f"关闭完成: {doc_title}")
                    except Exception as save_err:
                        self.log_message.emit(f"保存关闭失败 {os.path.basename(file_path)}: {str(save_err)}")
                    
                    success_count += 1
                    self.log_message.emit(f"成功处理: {os.path.basename(file_path)}")
                    # 发送文件处理成功信号
                    self.file_status_update.emit(file_path, 'completed')
                    
                except Exception as e:
                    error_msg = f"处理失败 {os.path.basename(file_path)}: {str(e)}"
                    self.log_message.emit(error_msg)
                    error_messages.append(error_msg)
                    # 发送文件处理失败信号
                    self.file_status_update.emit(file_path, 'failed')
                    
                finally:
                    # 异常时尝试关闭文档（兜底逻辑）
                    if 'sw_doc' in locals() and sw_doc is not None:
                        try:
                            # 兼容属性和方法两种访问方式
                            doc_title = sw_doc.GetTitle
                            if callable(doc_title):
                                doc_title = doc_title()
                            sw_doc.Save()
                            sw_doc = None
                            sw_app.CloseDoc(doc_title)
                        except Exception:
                            pass
                
                # 更新进度
                progress = int((index + 1) / total_files * 100)
                self.progress_update.emit(progress, f"{index + 1}/{total_files}")
            
            if self.is_running:
                result_msg = f"批量执行完成！成功: {success_count}, 失败: {len(error_messages)}"
                self.log_message.emit(result_msg)
                self.task_complete.emit(len(error_messages) == 0, result_msg)
            
            pythoncom.CoUninitialize()
            
        except Exception as e:
            error_msg = f"致命错误: {str(e)}"
            self.log_message.emit(error_msg)
            self.task_complete.emit(False, error_msg)
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def execute_macro(self, sw_app, macro_path, module_name):
        """执行指定的宏模块"""
        # 模块名称格式应该是 "模块名.过程名"
        module_name = str(module_name) if module_name else ""
        if "." in module_name:
            parts = module_name.split(".", 1)  # 只按第一个点分割
            mod_name = parts[0]
            proc_name = parts[1] if len(parts) > 1 else "main"
        else:
            mod_name = module_name
            proc_name = "main"
        
        # 确保参数类型正确
        macro_path = str(macro_path)
        
        # 启动自动关闭弹窗线程（用于处理宏内部的 MsgBox 等弹窗）
        auto_close_event = threading.Event()
        auto_close_thread = threading.Thread(
            target=self._auto_close_sw_dialogs,
            args=(auto_close_event,),
            daemon=True
        )
        auto_close_thread.start()
        
        try:
            # RunMacro2(FileName, ModuleName, ProcedureName, Options, ErrorCode)
            # ErrorCode 是 [out] long 类型，需要 VT_BYREF 传递
            run_error = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            sw_app.RunMacro2(macro_path, mod_name, proc_name, 0, run_error)
        finally:
            # 停止自动关闭弹窗线程
            auto_close_event.set()
            auto_close_thread.join(timeout=2)

    def _auto_close_sw_dialogs(self, stop_event):
        """后台线程：自动关闭 SolidWorks 的弹窗（如 MsgBox）"""
        # Windows API 常量
        WM_CLOSE = 0x0010
        BM_CLICK = 0x00F5
        
        user32 = ctypes.windll.user32
        
        # SW 常见弹窗类名
        dialog_classes = [
            "#32770",       # 标准对话框
            "NUIDialog",    # SW 新版对话框
        ]
        
        # SW 主窗口标题关键词
        sw_titles = ["SOLIDWORKS", "SolidWorks"]
        
        while not stop_event.is_set():
            try:
                # 枚举所有顶级窗口
                def enum_callback(hwnd, lparam):
                    if stop_event.is_set():
                        return False
                    
                    class_name = ctypes.create_unicode_buffer(256)
                    user32.GetClassNameW(hwnd, class_name, 256)
                    cls = class_name.value
                    
                    # 检查是否为弹窗类
                    if cls not in dialog_classes:
                        return True
                    
                    # 获取窗口标题
                    title = ctypes.create_unicode_buffer(512)
                    user32.GetWindowTextW(hwnd, title, 512)
                    title_str = title.value
                    
                    if not title_str:
                        return True
                    
                    # 检查是否为 SW 相关弹窗
                    is_sw = any(sw in title_str for sw in sw_titles)
                    
                    # 获取父窗口检查
                    parent = user32.GetParent(hwnd)
                    if not is_sw and parent:
                        parent_title = ctypes.create_unicode_buffer(512)
                        user32.GetWindowTextW(parent, parent_title, 512)
                        is_sw = any(sw in parent_title.value for sw in sw_titles)
                    
                    if not is_sw:
                        return True
                    
                    # 检查窗口是否可见和启用
                    if not user32.IsWindowVisible(hwnd):
                        return True
                    
                    # 找到"确定"或"OK"按钮并点击
                    def find_button(btn_hwnd, btn_lparam):
                        btn_class = ctypes.create_unicode_buffer(256)
                        user32.GetClassNameW(btn_hwnd, btn_class, 256)
                        if btn_class.value == "Button":
                            btn_text = ctypes.create_unicode_buffer(256)
                            user32.GetWindowTextW(btn_hwnd, btn_text, 256)
                            text = btn_text.value
                            # 中文"确定"或英文"OK"
                            if text in ("确定", "OK", "是(&Y)", "Yes", "&Yes"):
                                user32.SendMessageW(btn_hwnd, BM_CLICK, 0, 0)
                                return False
                        return True
                    
                    # 枚举子窗口找按钮
                    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
                    user32.EnumChildWindows(hwnd, WNDENUMPROC(find_button), 0)
                    
                    return True
                
                WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
                user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
                
            except Exception:
                pass
            
            time.sleep(0.5)  # 每 0.5 秒检查一次

    def stop(self):
        """停止任务"""
        self.is_running = False
        self.pause_event.set()  # 确保线程能退出等待状态

    def pause(self):
        """暂停任务"""
        self.is_paused = True
        self.pause_event.clear()

    def resume(self):
        """继续任务"""
        self.is_paused = False
        self.pause_event.set()

    def _safe_close_document(self, sw_app, sw_doc, file_path):
        """
        安全关闭文档（简化版）
        :param sw_app: SolidWorks应用程序对象
        :param sw_doc: 文档对象
        :param file_path: 文件完整路径
        """
        file_name = os.path.basename(file_path)
        
        # 尝试多种关闭方式，但不抛出异常
        try:
            # 方式1: 通过文档对象关闭
            if sw_doc is not None:
                try:
                    sw_doc.QuitDoc()
                    return
                except Exception:
                    pass
            
            # 方式2: 通过应用程序关闭（标题）
            file_title = os.path.splitext(file_name)[0]
            try:
                sw_app.CloseDoc(file_title)
                return
            except Exception:
                pass
            
            # 方式3: 通过应用程序关闭（完整路径）
            try:
                sw_app.CloseDoc(file_path)
                return
            except Exception:
                pass
        except Exception:
            # 静默失败，不记录错误日志
            pass

class MacroLoaderThread(QThread):
    """宏模块加载线程"""
    modules_loaded = Signal(list, str)
    load_error = Signal(str)
    
    # 静态缓存：避免重复创建SW对象和重复解析宏
    _sw_app_cache = None
    _module_cache = {}

    def __init__(self, macro_path, parent=None):
        super().__init__(parent)
        self.macro_path = macro_path

    def run(self):
        """加载宏模块（优化版）"""
        try:
            # 检查缓存
            if self.macro_path in MacroLoaderThread._module_cache:
                self.modules_loaded.emit(MacroLoaderThread._module_cache[self.macro_path], "")
                return
            
            # 获取或创建SW对象
            sw_app = self._get_sw_app()
            if sw_app is None:
                self.load_error.emit("无法连接到SolidWorks")
                return
            
            # 调用GetMacroMethods获取模块列表
            try:
                modules_raw = sw_app.GetMacroMethods(self.macro_path, swAllMethods)
            except Exception as e:
                # COM对象可能已失效，尝试重新连接
                if "对象没有连接到服务器" in str(e) or -2147220995 in str(e):
                    self._reset_sw_app()
                    sw_app = self._get_sw_app()
                    if sw_app is None:
                        self.load_error.emit("无法连接到SolidWorks")
                        return
                    modules_raw = sw_app.GetMacroMethods(self.macro_path, swAllMethods)
                else:
                    raise
            
            # 处理返回结果
            module_list = self._parse_modules(modules_raw)
            
            # 更新缓存
            if module_list:
                MacroLoaderThread._module_cache[self.macro_path] = module_list
                self.modules_loaded.emit(module_list, "")
            else:
                self.modules_loaded.emit([], "宏程序中没有找到模块")
                
        except Exception as e:
            self.load_error.emit(f"加载宏模块失败: {str(e)}")

    def _get_sw_app(self):
        """获取或创建SolidWorks COM对象（复用缓存）"""
        try:
            if MacroLoaderThread._sw_app_cache is None:
                pythoncom.CoInitialize()
                MacroLoaderThread._sw_app_cache = win32com.client.Dispatch("SldWorks.Application")
            return MacroLoaderThread._sw_app_cache
        except Exception:
            return None
    
    def _reset_sw_app(self):
        """重置SW COM对象缓存，用于处理连接失效的情况"""
        try:
            # 尝试释放旧对象
            if MacroLoaderThread._sw_app_cache is not None:
                try:
                    # 尝试正常退出
                    MacroLoaderThread._sw_app_cache.ExitApp()
                except Exception:
                    pass
                # 释放COM对象
                del MacroLoaderThread._sw_app_cache
        except Exception:
            pass
        finally:
            MacroLoaderThread._sw_app_cache = None

    def _parse_modules(self, modules_raw):
        """快速解析宏模块列表"""
        if modules_raw is None:
            return []
            
        module_list = []
        try:
            # 尝试直接迭代（大多数情况下有效）
            for item in modules_raw:
                if isinstance(item, (list, tuple)):
                    for sub_item in item:
                        module_list.append(str(sub_item))
                else:
                    module_list.append(str(item))
        except TypeError:
            # 如果不能迭代，尝试转换为tuple
            try:
                raw_tuple = tuple(modules_raw)
                for item in raw_tuple:
                    if isinstance(item, (list, tuple)):
                        for sub_item in item:
                            module_list.append(str(sub_item))
                    else:
                        module_list.append(str(item))
            except Exception:
                # 单个值情况
                module_list.append(str(modules_raw))
        
        # 去重并保持顺序
        seen = set()
        result = []
        for m in module_list:
            if m not in seen:
                seen.add(m)
                result.append(m)
        return result

class MainWindow(QMainWindow):
    """主窗口类（支持多宏串联执行）"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SW批量宏命令脚本")
        self.setMinimumSize(800, 600)
        
        # 初始化变量
        self.worker_thread = None
        self.macro_units = []  # 存储宏程序选择单元 [(macro_edit, module_combo, browse_btn, delete_btn), ...]
        
        # 创建UI
        self.init_ui()
    
    def init_ui(self):
        """初始化UI界面（支持区域拉伸）"""
        # 创建全局滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.setCentralWidget(scroll_area)
        
        # 创建主垂直分割器（直接作为滚动区域的widget）
        main_splitter = QSplitter(Qt.Vertical)
        main_splitter.setHandleWidth(5)
        main_splitter.setChildrenCollapsible(False)
        main_splitter.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        scroll_area.setWidget(main_splitter)
        
        # 设置拉伸条样式，使用+++++符号让视觉更明显
        main_splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #d0d0d0;
                border: 1px solid #a0a0a0;
            }
            QSplitter::handle:vertical {
                height: 5px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #d0d0d0, stop:0.2 #e0e0e0, stop:0.4 #d0d0d0,
                    stop:0.6 #e0e0e0, stop:0.8 #d0d0d0, stop:1 #e0e0e0);
            }
            QSplitter::handle:vertical:hover {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #b0b0b0, stop:0.2 #c0c0c0, stop:0.4 #b0b0b0,
                    stop:0.6 #c0c0c0, stop:0.8 #b0b0b0, stop:1 #c0c0c0);
                border: 1px solid #808080;
            }
        """)
        
        # ========== 第一部分：文件夹选择 ==========
        folder_widget = QWidget()
        folder_layout = QVBoxLayout(folder_widget)
        folder_layout.setContentsMargins(10, 10, 10, 10)
        
        folder_group = QGroupBox("待处理文件夹")
        folder_group_layout = QHBoxLayout(folder_group)
        folder_group_layout.setSpacing(5)
        folder_group_layout.setContentsMargins(5, 5, 5, 5)
        
        self.folder_edit = QLineEdit()
        self.folder_edit.setPlaceholderText("请选择待处理的文件夹")
        self.folder_edit.setMinimumWidth(100)
        browse_btn = QPushButton("浏览")
        browse_btn.setFixedWidth(60)
        browse_btn.clicked.connect(self.browse_folder)
        
        folder_group_layout.addWidget(self.folder_edit)
        folder_group_layout.addWidget(browse_btn)
        folder_layout.addWidget(folder_group)
        
        main_splitter.addWidget(folder_widget)
        # 设置文件夹选择区域的固定大小（不可调整）
        folder_widget.setMinimumHeight(80)
        folder_widget.setMaximumHeight(80)
        
        # ========== 第二部分：文件列表（可拉伸） ==========
        file_list_widget = QWidget()
        file_list_layout = QVBoxLayout(file_list_widget)
        file_list_layout.setContentsMargins(10, 0, 10, 10)
        
        file_list_group = QGroupBox("文件列表")
        file_list_group_layout = QVBoxLayout(file_list_group)
        file_list_group_layout.setContentsMargins(5, 5, 5, 5)
        
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_list_widget.setSortingEnabled(True)
        self.file_list_widget.setSelectionBehavior(QListWidget.SelectRows)
        self.file_list_widget.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.file_list_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.file_status = {}
        
        btn_layout = QHBoxLayout()
        self.clear_selected_btn = QPushButton("清除所选")
        self.clear_part_btn = QPushButton("清除零件")
        self.clear_assembly_btn = QPushButton("清除装配体")
        self.clear_drawing_btn = QPushButton("清除工程图")
        self.clear_all_btn = QPushButton("清除所有")
        
        self.clear_selected_btn.clicked.connect(self.clear_selected)
        self.clear_part_btn.clicked.connect(lambda: self.clear_by_extension(".sldprt"))
        self.clear_assembly_btn.clicked.connect(lambda: self.clear_by_extension(".sldasm"))
        self.clear_drawing_btn.clicked.connect(lambda: self.clear_by_extension(".slddrw"))
        self.clear_all_btn.clicked.connect(self.clear_all)
        
        btn_layout.addWidget(self.clear_selected_btn)
        btn_layout.addWidget(self.clear_part_btn)
        btn_layout.addWidget(self.clear_assembly_btn)
        btn_layout.addWidget(self.clear_drawing_btn)
        btn_layout.addWidget(self.clear_all_btn)
        
        file_list_group_layout.addWidget(self.file_list_widget)
        file_list_group_layout.addLayout(btn_layout)
        file_list_layout.addWidget(file_list_group)
        
        main_splitter.addWidget(file_list_widget)
        # 设置文件列表区域的固定大小范围
        file_list_widget.setMinimumHeight(150)
        file_list_widget.setMaximumHeight(9999)
        file_list_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main_splitter.setStretchFactor(1, 1)  # 文件列表可拉伸
        
        # ========== 第三部分：宏程序选择（可拉伸） ==========
        macro_widget = QWidget()
        macro_layout = QVBoxLayout(macro_widget)
        macro_layout.setContentsMargins(10, 0, 10, 10)
        
        macro_group = QGroupBox("宏程序选择")
        macro_group_layout = QVBoxLayout(macro_group)
        macro_group_layout.setContentsMargins(5, 5, 5, 5)
        
        self.macro_scroll_area = QScrollArea()
        self.macro_scroll_area.setWidgetResizable(True)
        self.macro_scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # 取消垂直滑动条
        self.macro_scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        self.macro_container = QWidget()
        self.macro_layout = QVBoxLayout(self.macro_container)
        self.macro_layout.setSpacing(0)  # 宏程序单元之间紧贴
        self.macro_layout.setContentsMargins(0, 0, 0, 0)
        
        self.add_macro_unit()
        
        self.macro_scroll_area.setWidget(self.macro_container)
        macro_group_layout.addWidget(self.macro_scroll_area)
        
        add_macro_btn = QPushButton("新增宏")
        add_macro_btn.clicked.connect(self.add_macro_unit)
        macro_group_layout.addWidget(add_macro_btn)
        
        macro_layout.addWidget(macro_group)
        
        main_splitter.addWidget(macro_widget)
        # 设置宏程序选择区域的固定大小范围（初始高度150，每增加/删除一个宏增减100）
        # 1个宏: 150px, 2个宏: 250px, 3个宏: 350px, ...
        self.macro_widget = macro_widget  # 保存引用以便动态调整高度
        macro_widget.setMinimumHeight(150)  # 1个宏程序的高度
        macro_widget.setMaximumHeight(9999)
        macro_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main_splitter.setStretchFactor(2, 1)  # 宏程序选择区域可拉伸
        
        # ========== 第四部分：进度和耗时 ==========
        progress_widget = QWidget()
        progress_layout = QVBoxLayout(progress_widget)
        progress_layout.setContentsMargins(10, 0, 10, 10)
        
        # 进度条区域
        progress_group = QGroupBox("进度")
        progress_group_layout = QHBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("")  # 移除百分比显示
        
        progress_group_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(progress_group)
        
        # 运行耗时显示
        time_group = QGroupBox("运行耗时")
        time_group_layout = QHBoxLayout(time_group)
        time_group_layout.setContentsMargins(5, 5, 5, 5)
        
        self.elapsed_time_label = QLabel("00:00:00")
        self.elapsed_time_label.setStyleSheet("font-family: monospace; font-size: 16px;")
        self.elapsed_time_label.setAlignment(Qt.AlignCenter)
        self.elapsed_time_label.setFixedWidth(80)  # 适配时间格式长度 HH:MM:SS
        
        time_group_layout.addWidget(self.elapsed_time_label)
        progress_layout.addWidget(time_group)
        
        main_splitter.addWidget(progress_widget)
        # 设置进度和耗时区域的固定大小（不可调整）
        progress_widget.setMinimumHeight(120)
        progress_widget.setMaximumHeight(120)
        
        # ========== 第五部分：日志区域（可拉伸） ==========
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(10, 0, 10, 10)
        
        log_group = QGroupBox("处理日志")
        log_group_layout = QVBoxLayout(log_group)
        log_group_layout.setContentsMargins(5, 5, 5, 5)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        log_btn_layout = QHBoxLayout()
        self.save_log_btn = QPushButton("保存日志")
        self.clear_log_btn = QPushButton("清空日志")
        self.save_log_btn.clicked.connect(self.save_log)
        self.clear_log_btn.clicked.connect(self.clear_log)
        
        log_btn_layout.addStretch()
        log_btn_layout.addWidget(self.save_log_btn)
        log_btn_layout.addWidget(self.clear_log_btn)
        
        log_group_layout.addWidget(self.log_text)
        log_group_layout.addLayout(log_btn_layout)
        log_layout.addWidget(log_group)
        
        main_splitter.addWidget(log_widget)
        # 设置日志区域的固定大小范围
        log_widget.setMinimumHeight(150)
        log_widget.setMaximumHeight(9999)
        log_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        main_splitter.setStretchFactor(4, 1)  # 日志区域可拉伸
        
        # ========== 第六部分：执行按钮 ==========
        btn_widget = QWidget()
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setContentsMargins(10, 0, 10, 10)
        
        self.start_btn = QPushButton("开始")
        self.pause_btn = QPushButton("暂停")
        self.pause_btn.setEnabled(False)
        
        self.start_btn.clicked.connect(self.toggle_start_stop)
        self.pause_btn.clicked.connect(self.toggle_pause_resume)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.pause_btn)
        
        main_splitter.addWidget(btn_widget)
        # 设置按钮区域的固定大小（不可调整）
        btn_widget.setMinimumHeight(70)
        btn_widget.setMaximumHeight(70)
        
        # 设置各区域的初始大小（按顺序设置）
        main_splitter.setSizes([80, 250, 200, 120, 300, 70])  # 文件夹、文件列表、宏程序、进度、日志、按钮
        
        # 设置初始大小
        self.setMinimumSize(800, 700)
        self.setMaximumSize(16777215, 49999)  # 设置主界面最大高度为49999
        
        # 初始化耗时计时器
        self.start_time = None
        self.elapsed_time_timer = None
    
    def add_macro_unit(self):
        """添加一个宏程序选择单元"""
        unit_index = len(self.macro_units) + 1
        
        # 创建宏程序选择单元的布局
        unit_widget = QWidget()
        unit_layout = QVBoxLayout(unit_widget)
        unit_layout.setSpacing(0)
        unit_layout.setContentsMargins(0, 0, 0, 0)
        
        # 标题标签（紧贴地址框）
        title_label = QLabel(f"宏程序 {unit_index}")
        title_label.setStyleSheet("font-weight: bold; padding: 0; margin: 0; border: none;")
        title_label.setAlignment(Qt.AlignLeft)
        title_label.setContentsMargins(0, 0, 0, 0)
        unit_layout.addWidget(title_label)
        
        # 宏程序路径选择（紧贴标签）
        path_layout = QHBoxLayout()
        path_layout.setSpacing(3)
        path_layout.setContentsMargins(0, 0, 0, 0)
        
        macro_edit = QLineEdit()
        macro_edit.setPlaceholderText("请选择宏程序文件")
        macro_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        
        browse_btn = QPushButton("浏览宏程序")
        browse_btn.clicked.connect(lambda checked, edit=macro_edit: self.browse_macro(edit))
        
        delete_btn = QPushButton("删除宏")
        delete_btn.clicked.connect(lambda checked, widget=unit_widget: self.remove_macro_unit(widget))
        
        path_layout.addWidget(macro_edit)
        path_layout.addWidget(browse_btn)
        path_layout.addWidget(delete_btn)
        unit_layout.addLayout(path_layout)
        
        # 宏程序模块选择
        module_layout = QHBoxLayout()
        module_label = QLabel("宏程序模块:")
        module_combo = QComboBox()
        
        module_layout.addWidget(module_label)
        module_layout.addWidget(module_combo)
        unit_layout.addLayout(module_layout)
        
        # 添加到宏程序容器
        self.macro_layout.addWidget(unit_widget)
        
        # 保存到宏程序单元列表
        self.macro_units.append({
            'widget': unit_widget,
            'macro_edit': macro_edit,
            'module_combo': module_combo,
            'browse_btn': browse_btn,
            'delete_btn': delete_btn
        })
        
        # 动态调整宏选择区域高度（初始150，每增加一个宏增加100）
        # 1个宏: 150px, 2个宏: 250px, 3个宏: 350px, ...
        if hasattr(self, 'macro_widget'):
            macro_count = len(self.macro_units)
            new_height = 150 + (macro_count - 1) * 75  # 初始150，每增加一个增加100
            self.macro_widget.setMinimumHeight(new_height)
            self.macro_widget.setMaximumHeight(max(new_height, 9999))
    
    def remove_macro_unit(self, widget):
        """删除指定的宏程序选择单元"""
        # 找到并移除
        for i, unit in enumerate(self.macro_units):
            if unit['widget'] == widget:
                # 从布局中移除
                self.macro_layout.removeWidget(widget)
                widget.deleteLater()
                
                # 从列表中移除
                del self.macro_units[i]
                
                # 更新剩余宏程序的序号
                self.update_macro_unit_labels()
                break
        
        # 动态调整宏选择区域高度（删除一个宏减少100）
        # 1个宏: 150px, 2个宏: 250px, 3个宏: 350px, ...
        if hasattr(self, 'macro_widget'):
            macro_count = len(self.macro_units)
            new_height = 150 + (macro_count - 1) * 100  # 初始150，每减少一个减少100
            self.macro_widget.setMinimumHeight(new_height)
            self.macro_widget.setMaximumHeight(max(new_height, 9999))
        
        # 确保至少保留一个宏程序选择单元
        if len(self.macro_units) == 0:
            self.add_macro_unit()
    
    def update_macro_unit_labels(self):
        """更新宏程序单元的序号标签"""
        for i, unit in enumerate(self.macro_units):
            # 找到序号标签并更新
            layout = unit['widget'].layout()
            if layout and layout.count() > 0:
                title_label = layout.itemAt(0).widget()
                if title_label and isinstance(title_label, QLabel):
                    title_label.setText(f"宏程序 {i + 1}")
    
    def browse_folder(self):
        """浏览选择文件夹"""
        from PySide6.QtWidgets import QFileDialog
        folder_path = QFileDialog.getExistingDirectory(
            self, "选择待处理文件夹", "", QFileDialog.ShowDirsOnly
        )
        
        if folder_path:
            self.folder_edit.setText(folder_path)
            self.load_files(folder_path)
    
    def load_files(self, folder_path):
        """加载文件夹中的SolidWorks文件"""
        self.file_list_widget.clear()
        self.file_status.clear()
        extensions = [".sldprt", ".sldasm", ".slddrw"]
        
        try:
            # 排序：先按扩展名，再按文件名（保持字母顺序A-Z）
            all_files = []
            for filename in os.listdir(folder_path):
                if filename.startswith("~$"):
                    continue
                ext = os.path.splitext(filename)[1].lower()
                if ext in extensions:
                    all_files.append(filename)
            
            # 按文件名字母顺序排序
            all_files.sort(key=lambda x: x.lower())
            
            # 暂时禁用排序，避免addItem返回None
            old_sorting = self.file_list_widget.isSortingEnabled()
            self.file_list_widget.setSortingEnabled(False)
            
            try:
                for filename in all_files:
                    full_path = os.path.join(folder_path, filename)
                    # 创建 QListWidgetItem 对象，确保不为 None
                    item = QListWidgetItem(filename)
                    item.setData(Qt.UserRole, full_path)
                    self.file_list_widget.addItem(item)
                    self.file_status[full_path] = None
            finally:
                # 恢复排序设置
                self.file_list_widget.setSortingEnabled(old_sorting)
            
            self.log(f"已加载 {self.file_list_widget.count()} 个文件")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"加载文件失败: {str(e)}")
    
    def browse_macro(self, macro_edit):
        """浏览选择宏程序文件（支持多宏选择，带唯一性校验）"""
        from PySide6.QtWidgets import QFileDialog
        macro_path, _ = QFileDialog.getOpenFileName(
            self, "选择宏程序文件", "", "宏程序 (*.swp)"
        )
        
        if macro_path:
            # 唯一性校验
            if self.is_macro_duplicate(macro_path, macro_edit):
                QMessageBox.warning(self, "警告", "请选择其他宏")
                return
            
            # 清空之前的模块列表
            for unit in self.macro_units:
                if unit['macro_edit'] == macro_edit:
                    unit['module_combo'].clear()
                    break
            
            macro_edit.setText(macro_path)
            self.load_macro_modules(macro_path, macro_edit)
    
    def is_macro_duplicate(self, macro_path, current_edit):
        """检查宏程序文件是否重复"""
        macro_name = os.path.basename(macro_path).lower()
        
        for unit in self.macro_units:
            # 跳过当前正在编辑的单元
            if unit['macro_edit'] == current_edit:
                continue
            
            existing_path = unit['macro_edit'].text().strip()
            if existing_path:
                existing_name = os.path.basename(existing_path).lower()
                if existing_name == macro_name:
                    return True
        
        return False
    
    def load_macro_modules(self, macro_path, macro_edit):
        """加载宏程序中的模块列表（使用后台线程避免卡顿）"""
        if not HAS_WIN32COM:
            QMessageBox.warning(self, "错误", "未安装pywin32库")
            return
        
        # 使用后台线程加载宏模块
        self.macro_loader = MacroLoaderThread(macro_path)
        self.macro_loader.modules_loaded.connect(lambda modules, msg, edit=macro_edit: self.on_modules_loaded(modules, msg, edit))
        self.macro_loader.load_error.connect(self.on_macro_load_error)
        self.macro_loader.start()
    
    def on_modules_loaded(self, modules, message, macro_edit):
        """宏模块加载完成"""
        # 找到对应的module_combo
        module_combo = None
        for unit in self.macro_units:
            if unit['macro_edit'] == macro_edit:
                module_combo = unit['module_combo']
                break
        
        if module_combo is None:
            return
        
        module_combo.clear()
        if modules:
            for module in modules:
                # 确保 module 是字符串类型
                module_str = str(module) if not isinstance(module, str) else module
                module_combo.addItem(module_str)
            # 默认选择第一个，或者包含MAIN的
            for i in range(module_combo.count()):
                item_text = module_combo.itemText(i)
                if "MAIN" in item_text.upper():
                    module_combo.setCurrentIndex(i)
                    break
            else:
                if module_combo.count() > 0:
                    module_combo.setCurrentIndex(0)
            self.log(f"已加载 {len(modules)} 个宏模块")
        else:
            QMessageBox.information(self, "提示", message or "宏程序中没有找到模块")
    
    def on_macro_load_error(self, error_msg):
        """宏模块加载错误"""
        QMessageBox.warning(self, "错误", error_msg)
    
    def clear_selected(self):
        """清除所选文件"""
        for item in reversed(self.file_list_widget.selectedItems()):
            # 从状态字典中移除
            file_path = item.data(Qt.UserRole)
            if file_path and file_path in self.file_status:
                del self.file_status[file_path]
            row = self.file_list_widget.row(item)
            self.file_list_widget.takeItem(row)
    
    def clear_by_extension(self, extension):
        """按扩展名清除文件"""
        extension = extension.lower()
        i = self.file_list_widget.count() - 1
        while i >= 0:
            item = self.file_list_widget.item(i)
            if item.text().lower().endswith(extension):
                # 从状态字典中移除
                file_path = item.data(Qt.UserRole)
                if file_path and file_path in self.file_status:
                    del self.file_status[file_path]
                self.file_list_widget.takeItem(i)
            i -= 1
    
    def clear_all(self):
        """清除所有文件"""
        self.file_list_widget.clear()
        self.file_status.clear()
    
    def set_file_status(self, file_path, status):
        """
        设置文件处理状态
        :param file_path: 文件完整路径
        :param status: None=未处理, 'completed'=完成, 'failed'=失败
        """
        if file_path in self.file_status:
            self.file_status[file_path] = status
            self.update_file_item_visual(file_path, status)
    
    def update_file_item_visual(self, file_path, status):
        """
        更新文件列表项的视觉显示
        :param file_path: 文件完整路径
        :param status: None=未处理, 'completed'=完成, 'failed'=失败
        """
        # 查找对应的item
        for i in range(self.file_list_widget.count()):
            item = self.file_list_widget.item(i)
            if item.data(Qt.UserRole) == file_path:
                # 设置背景色
                if status == 'completed':
                    # 处理完成：绿色背景
                    item.setBackground(Qt.green)
                    item.setForeground(Qt.black)
                elif status == 'failed':
                    # 处理失败：红色背景
                    item.setBackground(Qt.red)
                    item.setForeground(Qt.white)
                else:
                    # 未处理或重置：使用默认样式
                    item.setBackground(Qt.transparent)
                    item.setForeground(Qt.black)
                break
    
    def get_selected_files(self):
        """获取选中的文件列表（完整路径）"""
        selected_files = []
        for item in self.file_list_widget.selectedItems():
            file_path = item.data(Qt.UserRole)
            if file_path:
                selected_files.append(file_path)
        return selected_files
    
    def save_log(self):
        """保存日志到文件"""
        from PySide6.QtWidgets import QFileDialog
        log_text = self.log_text.toPlainText()
        if not log_text.strip():
            QMessageBox.warning(self, "警告", "日志内容为空")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存日志", "", "文本文件 (*.txt);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(log_text)
                QMessageBox.information(self, "成功", "日志保存成功")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"保存日志失败: {str(e)}")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def log(self, message):
        """添加日志消息"""
        self.log_text.append(message)
        # 滚动到底部
        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.log_text.setTextCursor(cursor)
    
    def toggle_start_stop(self):
        """开始/中止任务切换（支持多宏串联执行）"""
        # 如果正在运行或暂停中，中止任务
        if self.worker_thread and (self.worker_thread.isRunning() or hasattr(self.worker_thread, 'is_paused') and self.worker_thread.is_paused):
            self.worker_thread.stop()
            self.start_btn.setText("开始")
            self.pause_btn.setEnabled(False)
            self.pause_btn.setText("暂停")
            self.stop_elapsed_timer()
            self.log("正在中止任务...")
            return
        
        # 开始新任务
        # 验证输入
        file_count = self.file_list_widget.count()
        if file_count == 0:
            QMessageBox.warning(self, "警告", "文件列表为空")
            return
        
        # 收集宏程序配置
        macro_configs = []
        for unit in self.macro_units:
            macro_path = unit['macro_edit'].text().strip()
            module_name = unit['module_combo'].currentText()
            
            if not macro_path:
                QMessageBox.warning(self, "警告", "请选择所有宏程序文件")
                return
            
            if not module_name:
                QMessageBox.warning(self, "警告", "请选择所有宏程序模块")
                return
            
            macro_configs.append({
                'macro_path': macro_path,
                'module_name': module_name
            })
        
        if len(macro_configs) == 0:
            QMessageBox.warning(self, "警告", "请至少添加一个宏程序")
            return
        
        # 获取文件列表（使用存储的完整路径）
        file_list = []
        for i in range(file_count):
            item = self.file_list_widget.item(i)
            file_path = item.data(Qt.UserRole)
            if file_path:
                file_list.append(file_path)
            else:
                file_list.append(item.text())
        
        # 重置所有文件状态
        for fp in file_list:
            self.set_file_status(fp, None)
        
        # 开始任务
        self.start_btn.setText("中止")
        self.pause_btn.setEnabled(True)
        self.pause_btn.setText("暂停")
        self.progress_bar.setValue(0)
        self.log_text.clear()
        self.elapsed_time_label.setText("00:00:00")
        
        # 重置并启动耗时计时器
        self.start_elapsed_timer()
        
        # 记录将要执行的宏程序列表
        self.log(f"即将执行 {len(macro_configs)} 个宏程序:")
        for i, config in enumerate(macro_configs):
            self.log(f"  [{i + 1}] {os.path.basename(config['macro_path'])} - {config['module_name']}")
        
        self.worker_thread = WorkerThread(file_list, macro_configs)
        self.worker_thread.progress_update.connect(self.update_progress)
        self.worker_thread.task_complete.connect(self.task_complete)
        self.worker_thread.log_message.connect(self.log)
        self.worker_thread.file_status_update.connect(self.set_file_status)
        self.worker_thread.start()
    
    def start_elapsed_timer(self):
        """启动耗时计时器"""
        self.start_time = time.time()
        self.elapsed_time_label.setText("00:00:00")
        
        if self.elapsed_time_timer:
            self.elapsed_time_timer.stop()
        
        self.elapsed_time_timer = QTimer()
        self.elapsed_time_timer.timeout.connect(self.update_elapsed_time)
        self.elapsed_time_timer.start(1000)  # 每秒更新一次
    
    def stop_elapsed_timer(self):
        """停止耗时计时器"""
        if self.elapsed_time_timer:
            self.elapsed_time_timer.stop()
            self.elapsed_time_timer = None
    
    def update_elapsed_time(self):
        """更新耗时显示"""
        if self.start_time is None:
            return
        
        elapsed = int(time.time() - self.start_time)
        hours = elapsed // 3600
        minutes = (elapsed % 3600) // 60
        seconds = elapsed % 60
        self.elapsed_time_label.setText(f"{hours:02d}:{minutes:02d}:{seconds:02d}")
    
    def update_progress(self, value, text):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def toggle_pause_resume(self):
        """暂停/继续任务切换"""
        if self.worker_thread and self.worker_thread.isRunning():
            if self.pause_btn.text() == "暂停":
                self.worker_thread.pause()
                self.pause_btn.setText("继续")
                self.log("任务已暂停")
            else:
                self.worker_thread.resume()
                self.pause_btn.setText("暂停")
                self.log("任务继续执行")
    
    def task_complete(self, success, message):
        """任务完成处理"""
        self.start_btn.setText("开始")
        self.pause_btn.setEnabled(False)
        self.pause_btn.setText("暂停")
        self.stop_elapsed_timer()
        
        if success:
            QMessageBox.information(self, "完成", message)
        else:
            QMessageBox.warning(self, "完成", message)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 设置风格
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())