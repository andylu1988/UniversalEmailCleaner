import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import json
import os
import sys
import threading
import requests
import time
import csv
import subprocess
from datetime import datetime, timedelta
import logging
import traceback

# Try importing dependencies
try:
    from azure.identity import InteractiveBrowserCredential
except ImportError:
    pass

EXCHANGELIB_ERROR = None
try:
    from exchangelib import Account, Credentials, Configuration, DELEGATE, IMPERSONATION, Message, Mailbox, EWSDateTime, CalendarItem
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    from exchangelib.properties import FieldPath
    from exchangelib.fields import ExtendedPropertyField
    # Ignore SSL warnings for self-signed certs if needed
    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
except ImportError as e:
    EXCHANGELIB_ERROR = str(e)
    # Define dummy classes to prevent NameError if used before check
    class Credentials: pass
    class Account: pass
    class Configuration: pass
    class Message: pass
    class Mailbox: pass
    class EWSDateTime: pass
    class CalendarItem: pass
    DELEGATE = None
    IMPERSONATION = None


def guess_calendar_item_type(item):
    """Guess CalendarItemType with fallbacks when exchangelib mapping is missing."""
    # 1) Prefer server-provided CalendarItemType
    cit = getattr(item, "calendar_item_type", None)
    if cit:
        cit_str = getattr(cit, "value", None) or str(cit)
        for v in ("RecurringMaster", "Occurrence", "Exception", "Single"):
            if v.lower() in cit_str.lower():
                return v

    # 2) Fallbacks based on recurrence-related fields
    if getattr(item, "recurrence", None):
        return "RecurringMaster"

    if getattr(item, "recurrence_id", None) or getattr(item, "original_start", None):
        return "Occurrence"

    if getattr(item, "recurring_master_id", None):
        return "Occurrence"

    # 3) If IsRecurring=True but lacking other hints, treat as master for safety
    if getattr(item, "is_recurring", False):
        return "RecurringMaster"

    # Default to Single
    return "Single"


def translate_pattern_type(ptype_name):
    mapping = {
        'DailyPattern': '按天',
        'WeeklyPattern': '按周',
        'AbsoluteMonthlyPattern': '按月(固定)',
        'RelativeMonthlyPattern': '按月(相对)',
        'AbsoluteYearlyPattern': '按年(固定)',
        'RelativeYearlyPattern': '按年(相对)',
        'RegeneratingPattern': '重新生成'
    }
    return mapping.get(ptype_name, ptype_name)


def get_pattern_details(pattern_obj):
    """
    Extract detailed recurrence pattern information.
    Returns a string like "按周: 星期=周一, 周三, 间隔=1"
    """
    if not pattern_obj:
        return ""
    
    raw_type = pattern_obj.__class__.__name__
    pattern_type = translate_pattern_type(raw_type)
    details = []
    
    # Helper for weekdays
    weekday_map = {
        'Mon': '周一', 'Tue': '周二', 'Wed': '周三', 'Thu': '周四', 'Fri': '周五', 'Sat': '周六', 'Sun': '周日',
        'Monday': '周一', 'Tuesday': '周二', 'Wednesday': '周三', 'Thursday': '周四', 'Friday': '周五', 'Saturday': '周六', 'Sunday': '周日'
    }
    
    # Extract common pattern attributes
    if hasattr(pattern_obj, 'interval'):
        details.append(f"间隔={pattern_obj.interval}")
    
    if hasattr(pattern_obj, 'days_of_week'):
        dow = pattern_obj.days_of_week
        if dow:
            if isinstance(dow, (list, tuple)):
                days_str = ", ".join(weekday_map.get(str(d), str(d)) for d in dow)
            else:
                days_str = weekday_map.get(str(dow), str(dow))
            details.append(f"星期={days_str}")
    
    if hasattr(pattern_obj, 'day_of_month'):
        details.append(f"日期={pattern_obj.day_of_month}日")
    
    if hasattr(pattern_obj, 'first_day_of_week'):
        fd = str(pattern_obj.first_day_of_week)
        details.append(f"周首日={weekday_map.get(fd, fd)}")
    
    if hasattr(pattern_obj, 'month'):
        details.append(f"月份={pattern_obj.month}月")
    
    if hasattr(pattern_obj, 'day_of_week_index'):
        # First, Second, Third, Fourth, Last
        idx_map = {'First': '第一个', 'Second': '第二个', 'Third': '第三个', 'Fourth': '第四个', 'Last': '最后一个'}
        idx = str(pattern_obj.day_of_week_index)
        details.append(f"索引={idx_map.get(idx, idx)}")
    
    details_str = ", ".join(details) if details else ""
    return f"{pattern_type}: {details_str}" if details_str else pattern_type


def get_recurrence_duration(recurrence_obj):
    """
    Extract recurrence duration information.
    Returns format like:
      - "无限期" for no end date
      - "结束于: 2025-12-31" if has end date
      - "共 10 次" if limited by count
    """
    if not recurrence_obj:
        return ""
    
    try:
        # Try to find the boundary object
        # In exchangelib, Recurrence.boundary holds the EndDate/NoEnd/Numbered recurrence
        boundary = getattr(recurrence_obj, 'boundary', recurrence_obj)
        
        # Debug logging
        logging.debug(f"Recurrence Check: BoundaryType={type(boundary).__name__}")
        
        # 1. Check for End Date (attribute is 'end' in exchangelib, but sometimes 'end_date' in other contexts)
        end_date = getattr(boundary, 'end', None) or getattr(boundary, 'end_date', None) or getattr(recurrence_obj, 'end_date', None)
        if end_date:
            return f"结束于: {end_date}"
            
        # 2. Check for Number of Occurrences
        number = getattr(boundary, 'number', None) or getattr(recurrence_obj, 'number', None) or getattr(recurrence_obj, 'max_occurrences', None)
        if number:
            return f"共 {number} 次"

        # 3. Check for No End
        # Check class name or no_end attribute
        b_type = boundary.__class__.__name__
        if 'NoEnd' in b_type:
            return "无限期"
            
        if getattr(boundary, 'no_end', False) or getattr(recurrence_obj, 'no_end', False):
            return "无限期"
    
        # 4. Fallback: Inspect all attributes
        if hasattr(boundary, '__dict__'):
            for k, v in boundary.__dict__.items():
                if k in ('end', 'end_date') and v:
                    return f"结束于: {v}"
                if k in ('number', 'max_occurrences') and v and isinstance(v, int):
                    return f"共 {v} 次"

    except Exception as e:
        logging.error(f"Error extracting recurrence duration: {e}")
        
    return "未知"

    return "未知"

    # 3. Check for No End
    # Check class name or no_end attribute
    b_type = boundary.__class__.__name__
    if b_type == 'NoEndRecurrence':
        return "无限期"
        
    if getattr(boundary, 'no_end', False) or getattr(recurrence_obj, 'no_end', False):
        return "无限期"
    
    return "未知"


def is_endless_recurring(item_type, recurrence_obj):
    """
    Check if this is a RecurringMaster with endless recurrence.
    Returns "True" if master + endless, otherwise "N/A"
    """
    if item_type != "RecurringMaster":
        return "N/A"
    
    if not recurrence_obj:
        return "N/A"
    
    boundary = getattr(recurrence_obj, 'boundary', recurrence_obj)
    
    # Check explicit NoEnd
    if boundary.__class__.__name__ == 'NoEndRecurrence':
        return "True"
    if getattr(boundary, 'no_end', False) or getattr(recurrence_obj, 'no_end', False):
        return "True"
        
    # Check for EndDate or Number (Definite End)
    end_date = getattr(boundary, 'end_date', None) or getattr(recurrence_obj, 'end_date', None)
    if end_date:
        return "False"
        
    number = getattr(boundary, 'number', None) or getattr(recurrence_obj, 'number', None) or getattr(recurrence_obj, 'max_occurrences', None)
    if number:
        return "False"

    return "False"
    
    if not (has_end_date or has_max_occurrences or has_no_end):
        return "True"
    
    return "N/A"


class Logger:
    def __init__(self, log_area, log_file_path):
        self.log_area = log_area
        self.log_file_path = log_file_path
        self.level = "NORMAL" # NORMAL or ADVANCED

    def set_level(self, level):
        self.level = level

    def log(self, message, level="INFO", is_advanced=False):
        # If message is advanced but current level is NORMAL, skip
        if is_advanced and self.level != "ADVANCED":
            return

        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        full_msg = f"[{timestamp}] [{level}] {message}"
        
        # GUI Update (Thread safe)
        def _update():
            self.log_area.config(state='normal')
            self.log_area.insert(tk.END, full_msg + "\n")
            self.log_area.see(tk.END)
            self.log_area.config(state='disabled')
        
        if self.log_area:
            self.log_area.after(0, _update)
        
        # File Write
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(full_msg + "\n")
        except:
            pass

    def log_to_file_only(self, message):
        """Writes directly to file, skipping GUI. Useful for large debug dumps."""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        full_msg = f"[{timestamp}] [DEBUG_DATA] {message}"
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(full_msg + "\n")
        except:
            pass

class EwsTraceAdapter(NoVerifyHTTPAdapter):
    logger = None
    log_responses = True  # Default to True, can be disabled for "Advanced" mode
    
    def send(self, request, *args, **kwargs):
        # Resolve stream argument - NTLM auth needs stream=True internally
        stream = kwargs.get('stream', False)
        if not stream and len(args) > 0:
            stream = args[0]

        if self.logger:
            try:
                tid = threading.get_ident()
                now = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%SZ")
                
                # Log Request Headers
                headers_str = f"{request.method} {request.path_url} HTTP/1.1\n"
                headers_str += "\n".join(f"{k}: {v}" for k, v in request.headers.items())
                
                trace_header = f'<Trace Tag="EwsRequestHttpHeaders" Tid="{tid}" Time="{now}">\n{headers_str}\n</Trace>'
                self.logger.debug(trace_header)
                
                # Log Request Body
                if request.body:
                    body_to_log = None
                    if isinstance(request.body, (str, bytes)):
                        body_to_log = request.body
                        if isinstance(body_to_log, bytes):
                            body_to_log = body_to_log.decode('utf-8', errors='replace')
                    else:
                        body_to_log = f"[Body type {type(request.body)} - Not logged]"

                    trace_body = f'<Trace Tag="EwsRequest" Tid="{tid}" Time="{now}" Version="1.0">\n{body_to_log}\n</Trace>'
                    self.logger.debug(trace_body)
            except Exception:
                pass

        # CRITICAL: Call parent with original arguments to preserve NTLM auth behavior
        response = super().send(request, *args, **kwargs)
        
        # Capture response body immediately
        if self.log_responses:
            try:
                # Only capture XML or Text responses to avoid binary blobs
                content_type = response.headers.get('Content-Type', '')
                if 'xml' in content_type or 'text' in content_type:
                    # Force read content (this caches it in response.content)
                    content = response.content 
                    
                    # Write immediately to file
                    docs_dir = os.path.join(os.path.expanduser("~"), "Documents", "UniversalEmailCleaner")
                    os.makedirs(docs_dir, exist_ok=True)
                    log_path = os.path.join(docs_dir, "ews_getitem_responses.log")
                    
                    with open(log_path, "a", encoding="utf-8") as f:
                        f.write(f"\n\n{'='*80}\n")
                        f.write(f"Time: {datetime.now()}\n")
                        f.write(f"URL: {request.url}\n")
                        f.write(f"Status: {response.status_code}\n")
                        f.write(f"{'='*80}\n")
                        try:
                            f.write(content.decode('utf-8', errors='replace'))
                        except:
                            f.write("<Binary or undecodable content>")
                        f.write(f"\n{'='*80}\n")

            except Exception:
                pass
            
        return response


class UniversalEmailCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("通用邮件清理工具 (Graph API & EWS)")
        self.root.geometry("1100x900")
        
        style = ttk.Style()
        style.theme_use('clam')
        
        # --- Paths & Config ---
        self.documents_dir = os.path.join(os.path.expanduser("~"), "Documents", "UniversalEmailCleaner")
        if not os.path.exists(self.documents_dir):
            os.makedirs(self.documents_dir)
            
        self.log_file_path = os.path.join(self.documents_dir, "app.log")
        self.config_file_path = os.path.join(self.documents_dir, "config.json")
        self.reports_dir = os.path.join(self.documents_dir, "Reports")
        if not os.path.exists(self.reports_dir):
            os.makedirs(self.reports_dir)

        # 菜单栏
        menubar = tk.Menu(root)
        
        # 工具菜单
        tools_menu = tk.Menu(menubar, tearoff=0)
        
        # 日志配置子菜单
        log_menu = tk.Menu(tools_menu, tearoff=0)
        self.log_level_var = tk.StringVar(value="Normal") # Normal, Advanced, Expert
        
        def on_log_level_change():
            val = self.log_level_var.get()
            if val == "Expert":
                confirm = messagebox.askyesno("警告", "日志排错专用，日志量会很大且包含敏感信息，慎选！\n\n确认开启专家模式吗？")
                if not confirm:
                    self.log_level_var.set("Normal")
                    return
            # Sync with UI combobox if it exists (it will be created later, so we bind variable)
            
        log_menu.add_radiobutton(label="默认 (Default)", variable=self.log_level_var, value="Normal", command=on_log_level_change)
        log_menu.add_radiobutton(label="高级 (Advanced - 仅记录 EWS 请求)", variable=self.log_level_var, value="Advanced", command=on_log_level_change)
        log_menu.add_radiobutton(label="专家 (Expert - 记录 EWS 请求和响应)", variable=self.log_level_var, value="Expert", command=on_log_level_change)
        
        tools_menu.add_cascade(label="日志配置 (Log Level)", menu=log_menu)
        menubar.add_cascade(label="工具 (Tools)", menu=tools_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="版本历史 (Version History)", command=self.show_history)
        help_menu.add_command(label="关于 (About)", command=self.show_about)
        menubar.add_cascade(label="帮助 (Help)", menu=help_menu)
        root.config(menu=menubar)

        # --- Variables ---
        # Graph Config
        self.graph_auth_mode_var = tk.StringVar(value="Auto") # Auto (Cert) or Manual (Secret)
        self.app_id_var = tk.StringVar()
        self.tenant_id_var = tk.StringVar()
        self.thumbprint_var = tk.StringVar()
        self.client_secret_var = tk.StringVar()
        self.graph_env_var = tk.StringVar(value="Global")

        # EWS Config
        self.ews_server_var = tk.StringVar()
        self.ews_user_var = tk.StringVar()
        self.ews_pass_var = tk.StringVar()
        self.ews_auth_type_var = tk.StringVar(value="Impersonation") # Impersonation or Delegate
        self.ews_use_autodiscover = tk.BooleanVar(value=True)

        # Cleanup Config
        self.source_type_var = tk.StringVar(value="Graph") # Graph or EWS
        self.csv_path_var = tk.StringVar()
        self.report_only_var = tk.BooleanVar(value=True)
        # self.log_level_var is already defined in menu setup
        
        # Cleanup Target
        self.cleanup_target_var = tk.StringVar(value="Email") # Email or Meeting
        self.meeting_scope_var = tk.StringVar(value="All") # All, Single, Series
        self.meeting_only_cancelled_var = tk.BooleanVar(value=False)

        # Criteria
        self.criteria_msg_id = tk.StringVar()
        self.criteria_subject = tk.StringVar()
        self.criteria_sender = tk.StringVar()
        self.criteria_body = tk.StringVar()
        self.criteria_start_date = tk.StringVar()
        self.criteria_end_date = tk.StringVar()
        self.criteria_item_class = tk.StringVar(value="IPM.Note")

        # --- UI Layout ---
        main_frame = ttk.Frame(root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)

        # Tabs
        self.tab_connection = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_connection, text="1. 连接配置")
        
        self.tab_cleanup = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_cleanup, text="2. 任务配置")

        # Log Area
        log_frame = ttk.LabelFrame(main_frame, text="运行日志")
        log_frame.pack(fill="both", expand=True, pady=(10, 0))
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', font=("Consolas", 9))
        self.log_area.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.logger = Logger(self.log_area, self.log_file_path)

        # Links
        link_frame = ttk.Frame(log_frame)
        link_frame.pack(fill="x", padx=5)
        tk.Label(link_frame, text=f"日志文件: {self.log_file_path}", fg="blue", cursor="hand2").pack(side="left")
        self.report_link_lbl = tk.Label(link_frame, text="", fg="blue", cursor="hand2")
        self.report_link_lbl.pack(side="left", padx=20)

        # Build Tabs
        self.build_connection_tab()
        self.build_cleanup_tab()

        self.load_config()

    def log(self, msg, level="INFO", is_advanced=False):
        self.logger.log(msg, level, is_advanced)

    def update_report_link(self, path):
        def _update():
            self.report_link_lbl.config(text=f"最新报告: {path}")
            self.report_link_lbl.bind("<Button-1>", lambda e: os.startfile(path) if os.path.exists(path) else None)
        self.root.after(0, _update)

    def show_history(self):
        history_window = tk.Toplevel(self.root)
        history_window.title("版本历史")
        history_window.geometry("600x400")
        
        txt = scrolledtext.ScrolledText(history_window, padx=10, pady=10)
        txt.pack(fill="both", expand=True)
        
        # 尝试读取 CHANGELOG.md
        changelog_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CHANGELOG.md")
        if not os.path.exists(changelog_path):
            # 如果是打包后的环境，尝试在临时目录找
            if getattr(sys, 'frozen', False):
                 changelog_path = os.path.join(sys._MEIPASS, "CHANGELOG.md")

        if os.path.exists(changelog_path):
            with open(changelog_path, 'r', encoding='utf-8') as f:
                content = f.read()
                txt.insert(tk.END, content)
        else:
            txt.insert(tk.END, "未找到版本记录文件。")
            
        txt.config(state='disabled')

    def show_about(self):
        messagebox.showinfo("关于", "通用邮件清理工具 (Universal Email Cleaner) v1.3.2\n\n支持 Microsoft Graph API 和 Exchange Web Services (EWS)。\n用于批量清理或生成邮件报告。")

    # --- Config Management ---
    def load_config(self):
        if os.path.exists(self.config_file_path):
            try:
                with open(self.config_file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Graph
                    self.graph_auth_mode_var.set(config.get('graph_auth_mode', 'Auto'))
                    self.app_id_var.set(config.get('app_id', ''))
                    self.tenant_id_var.set(config.get('tenant_id', ''))
                    self.thumbprint_var.set(config.get('thumbprint', ''))
                    self.client_secret_var.set(config.get('client_secret', ''))
                    self.graph_env_var.set(config.get('graph_env', 'Global'))
                    # EWS
                    self.ews_server_var.set(config.get('ews_server', ''))
                    self.ews_user_var.set(config.get('ews_user', ''))
                    self.ews_use_autodiscover.set(config.get('ews_autodiscover', True))
                    self.ews_auth_type_var.set(config.get('ews_auth_type', 'Impersonation'))
                    # Common
                    self.source_type_var.set(config.get('source_type', 'Graph'))
                    self.csv_path_var.set(config.get('csv_path', ''))
                    self.log(">>> 配置已加载。")
            except Exception as e:
                self.log(f"X 加载配置失败: {e}", "ERROR")

    def save_config(self):
        config = {
            'graph_auth_mode': self.graph_auth_mode_var.get(),
            'app_id': self.app_id_var.get(),
            'tenant_id': self.tenant_id_var.get(),
            'thumbprint': self.thumbprint_var.get(),
            'client_secret': self.client_secret_var.get(),
            'graph_env': self.graph_env_var.get(),
            'ews_server': self.ews_server_var.get(),
            'ews_user': self.ews_user_var.get(),
            'ews_autodiscover': self.ews_use_autodiscover.get(),
            'ews_auth_type': self.ews_auth_type_var.get(),
            'source_type': self.source_type_var.get(),
            'csv_path': self.csv_path_var.get()
        }
        try:
            with open(self.config_file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
            self.log(">>> 配置已保存。")
        except Exception as e:
            self.log(f"X 保存配置失败: {e}", "ERROR")

    # --- Tab 1: Connection Setup ---
    def build_connection_tab(self):
        main_frame = ttk.Frame(self.tab_connection, padding=10)
        main_frame.pack(fill="both", expand=True)

        # 1. Connection Type Selection
        type_frame = ttk.LabelFrame(main_frame, text="连接模式选择")
        type_frame.pack(fill="x", pady=5)
        
        ttk.Radiobutton(type_frame, text="EWS (Exchange Web Services)", variable=self.source_type_var, value="EWS", command=self.toggle_connection_ui).pack(side="left", padx=20, pady=10)
        ttk.Radiobutton(type_frame, text="Microsoft Graph API", variable=self.source_type_var, value="Graph", command=self.toggle_connection_ui).pack(side="left", padx=20, pady=10)

        # 2. EWS Configuration Frame
        self.ews_frame = ttk.LabelFrame(main_frame, text="EWS 配置 (Exchange 2010-2019 / Online)")
        self.ews_frame.pack(fill="x", pady=5, ipady=5)
        
        ews_grid = ttk.Frame(self.ews_frame)
        ews_grid.pack(anchor="w", padx=10, pady=5)

        # Server
        ttk.Label(ews_grid, text="EWS 服务器:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_ews_server = ttk.Entry(ews_grid, textvariable=self.ews_server_var, width=40)
        self.entry_ews_server.grid(row=0, column=1, padx=5, pady=5)
        
        self.chk_ews_auto = ttk.Checkbutton(ews_grid, text="使用自动发现 (Autodiscover)", variable=self.ews_use_autodiscover, 
                                   command=lambda: self.entry_ews_server.config(state='disabled' if self.ews_use_autodiscover.get() else 'normal'))
        self.chk_ews_auto.grid(row=0, column=2, padx=5)
        
        # Credentials
        ttk.Label(ews_grid, text="管理员账号 (UPN):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(ews_grid, textvariable=self.ews_user_var, width=40).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(ews_grid, text="管理员密码:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        ttk.Entry(ews_grid, textvariable=self.ews_pass_var, show="*", width=40).grid(row=2, column=1, padx=5, pady=5)

        # Auth Type
        ttk.Label(ews_grid, text="访问类型:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        auth_frame = ttk.Frame(ews_grid)
        auth_frame.grid(row=3, column=1, sticky="w")
        ttk.Radiobutton(auth_frame, text="模拟 (Impersonation)", variable=self.ews_auth_type_var, value="Impersonation").pack(side="left", padx=5)
        ttk.Radiobutton(auth_frame, text="代理 (Delegate)", variable=self.ews_auth_type_var, value="Delegate").pack(side="left", padx=5)

        ttk.Button(self.ews_frame, text="测试 EWS 连接", command=self.test_ews_connection).pack(anchor="w", padx=10, pady=10)

        # 3. Graph Configuration Frame
        self.graph_frame = ttk.LabelFrame(main_frame, text="Graph API 配置 (Exchange Online)")
        self.graph_frame.pack(fill="x", pady=5, ipady=5)
        
        # Graph Common Settings (Environment & Mode)
        graph_common_frame = ttk.Frame(self.graph_frame)
        graph_common_frame.pack(anchor="w", padx=10, pady=5)
        
        # Environment Selection
        ttk.Label(graph_common_frame, text="环境:").pack(side="left")
        ttk.Radiobutton(graph_common_frame, text="全球版 (Global)", variable=self.graph_env_var, value="Global").pack(side="left", padx=10)
        ttk.Radiobutton(graph_common_frame, text="世纪互联 (China)", variable=self.graph_env_var, value="China").pack(side="left", padx=10)
        
        ttk.Label(graph_common_frame, text="|  配置方式:").pack(side="left", padx=10)
        ttk.Radiobutton(graph_common_frame, text="自动配置 (证书认证)", variable=self.graph_auth_mode_var, value="Auto", command=self.toggle_graph_ui).pack(side="left", padx=10)
        ttk.Radiobutton(graph_common_frame, text="手动配置 (Client Secret)", variable=self.graph_auth_mode_var, value="Manual", command=self.toggle_graph_ui).pack(side="left", padx=10)

        # Graph Auto Frame
        self.graph_auto_frame = ttk.Frame(self.graph_frame)
        self.graph_auto_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(self.graph_auto_frame, text="一键初始化 (创建 App & 证书)", command=self.start_graph_setup_thread).pack(side="left", padx=0)
        ttk.Button(self.graph_auto_frame, text="删除 App", command=self.start_delete_app_thread).pack(side="left", padx=5)

        # Graph Manual Frame
        self.graph_manual_frame = ttk.Frame(self.graph_frame)
        self.graph_manual_frame.pack(fill="x", padx=10, pady=5)
        
        manual_grid = ttk.Frame(self.graph_manual_frame)
        manual_grid.pack(anchor="w")
        
        ttk.Label(manual_grid, text="Tenant ID:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(manual_grid, textvariable=self.tenant_id_var, width=40).grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(manual_grid, text="App ID:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(manual_grid, textvariable=self.app_id_var, width=40).grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(manual_grid, text="Client Secret:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(manual_grid, textvariable=self.client_secret_var, width=40, show="*").grid(row=2, column=1, padx=5, pady=2)

        # Initial Toggle
        self.toggle_connection_ui()

    def toggle_connection_ui(self):
        mode = self.source_type_var.get()
        if mode == "EWS":
            self._enable_frame(self.ews_frame)
            self._disable_frame(self.graph_frame)
        else:
            self._disable_frame(self.ews_frame)
            self._enable_frame(self.graph_frame)
            self.toggle_graph_ui() # Re-apply graph internal state

    def toggle_graph_ui(self):
        if self.source_type_var.get() != "Graph":
            return # Don't mess if Graph is disabled
            
        mode = self.graph_auth_mode_var.get()
        if mode == "Auto":
            self.graph_auto_frame.pack(fill="x", padx=10, pady=5)
            self.graph_manual_frame.pack_forget()
        else:
            self.graph_auto_frame.pack_forget()
            self.graph_manual_frame.pack(fill="x", padx=10, pady=5)

    def _enable_frame(self, frame):
        for child in frame.winfo_children():
            try:
                child.configure(state='normal')
            except:
                pass
            # Recursively enable
            if isinstance(child, (ttk.Frame, ttk.LabelFrame)):
                self._enable_frame(child)

    def _disable_frame(self, frame):
        for child in frame.winfo_children():
            try:
                child.configure(state='disabled')
            except:
                pass
            # Recursively disable
            if isinstance(child, (ttk.Frame, ttk.LabelFrame)):
                self._disable_frame(child)

    def start_graph_setup_thread(self):
        threading.Thread(target=self.run_graph_setup, daemon=True).start()

    def start_delete_app_thread(self):
        if not self.app_id_var.get():
            messagebox.showerror("错误", "未找到 App ID，无法执行删除。")
            return
        if messagebox.askyesno("确认删除", f"确定要删除 Azure AD 应用 ({self.app_id_var.get()}) 吗？\n这将清除云端配置，且不可恢复！"):
            threading.Thread(target=self.run_delete_app, daemon=True).start()

    def run_graph_setup(self):
        try:
            env = self.graph_env_var.get()
            if env == "China":
                authority_host = "https://login.chinacloudapi.cn"
                graph_endpoint = "https://microsoftgraph.chinacloudapi.cn"
                scope = "https://microsoftgraph.chinacloudapi.cn/.default"
            else:
                authority_host = "https://login.microsoftonline.com"
                graph_endpoint = "https://graph.microsoft.com"
                scope = "https://graph.microsoft.com/.default"

            self.log(f">>> 正在启动 Azure 登录 ({env})...")
            from azure.identity import InteractiveBrowserCredential
            credential = InteractiveBrowserCredential(authority=authority_host)
            token = credential.get_token(scope)
            headers = {"Authorization": f"Bearer {token.token}", "Content-Type": "application/json"}
            
            # Get Tenant Info
            resp = requests.get(f"{graph_endpoint}/v1.0/organization", headers=headers)
            if resp.status_code != 200: raise Exception(f"获取租户失败: {resp.text}")
            org_info = resp.json()['value'][0]
            tenant_id = org_info['id']
            self.tenant_id_var.set(tenant_id)
            self.log(f"√ 租户 ID: {tenant_id}")

            # Generate Cert
            self.log(">>> 生成自签名证书...")
            ps_script = """
            $cert = New-SelfSignedCertificate -DnsName "UniversalEmailCleaner-Cert" -CertStoreLocation "cert:\\CurrentUser\\My" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter (Get-Date).AddYears(2)
            $thumbprint = $cert.Thumbprint
            $certContent = [System.Convert]::ToBase64String($cert.GetRawCertData())
            $result = @{Thumbprint=$thumbprint; Base64=$certContent}
            $result | ConvertTo-Json -Compress
            """
            cert_json = self.run_powershell_script(ps_script)
            cert_data = json.loads(cert_json)
            thumbprint = cert_data['Thumbprint']
            cert_blob = cert_data['Base64']
            self.thumbprint_var.set(thumbprint)
            self.log(f"√ 证书生成成功: {thumbprint}")

            # Create App
            self.log(">>> 创建 Azure AD 应用程序...")
            app_body = {
                "displayName": "UniversalEmailCleaner-App",
                "signInAudience": "AzureADMyOrg",
                "keyCredentials": [{"type": "AsymmetricX509Cert", "usage": "Verify", "key": cert_blob, "displayName": "Auto-Cert"}]
            }
            resp = requests.post(f"{graph_endpoint}/v1.0/applications", headers=headers, json=app_body)
            if resp.status_code != 201: raise Exception(f"创建 App 失败: {resp.text}")
            app_id = resp.json()['appId']
            self.app_id_var.set(app_id)
            self.log(f"√ App 创建成功: {app_id}")

            # Create SP
            time.sleep(5)
            sp_body = {"appId": app_id}
            resp = requests.post(f"{graph_endpoint}/v1.0/servicePrincipals", headers=headers, json=sp_body)
            if resp.status_code == 201:
                sp_id = resp.json()['id']
            else:
                resp = requests.get(f"{graph_endpoint}/v1.0/servicePrincipals?$filter=appId eq '{app_id}'", headers=headers)
                sp_id = resp.json()['value'][0]['id']
            self.log(f"√ 服务主体就绪: {sp_id}")

            # Grant Permissions
            self.log(">>> 正在授予 API 权限...")
            resp = requests.get(f"{graph_endpoint}/v1.0/servicePrincipals?$filter=appId eq '00000003-0000-0000-c000-000000000000'", headers=headers)
            graph_sp = resp.json()['value'][0]
            graph_sp_id = graph_sp['id']

            roles_to_add = ["Mail.ReadWrite", "User.Read.All", "Calendars.ReadWrite"]
            for role_name in roles_to_add:
                role_id = next((r['id'] for r in graph_sp['appRoles'] if r['value'] == role_name), None)
                if role_id:
                    assign_body = {"principalId": sp_id, "resourceId": graph_sp_id, "appRoleId": role_id}
                    r = requests.post(f"{graph_endpoint}/v1.0/servicePrincipals/{sp_id}/appRoleAssignments", headers=headers, json=assign_body)
                    if r.status_code in [201, 409]:
                        self.log(f"√ 权限 {role_name} 授予成功")
                    else:
                        self.log(f"X 权限 {role_name} 授予失败: {r.text}")
            
            self.log(">>> 初始化完成！")
            self.save_config()
            messagebox.showinfo("成功", "初始化完成！\nApp ID 和 证书指纹 已自动填入。")

        except Exception as e:
            self.log(f"X 错误: {e}", "ERROR")
            messagebox.showerror("错误", str(e))

    def run_delete_app(self):
        try:
            app_id = self.app_id_var.get()
            env = self.graph_env_var.get()
            
            if env == "China":
                authority_host = "https://login.chinacloudapi.cn"
                graph_endpoint = "https://microsoftgraph.chinacloudapi.cn"
                scope = "https://microsoftgraph.chinacloudapi.cn/.default"
            else:
                authority_host = "https://login.microsoftonline.com"
                graph_endpoint = "https://graph.microsoft.com"
                scope = "https://graph.microsoft.com/.default"

            self.log(f">>> 正在启动 Azure 登录以删除 App...")
            from azure.identity import InteractiveBrowserCredential
            credential = InteractiveBrowserCredential(authority=authority_host)
            token = credential.get_token(scope)
            headers = {"Authorization": f"Bearer {token.token}", "Content-Type": "application/json"}

            # Find App Object ID
            resp = requests.get(f"{graph_endpoint}/v1.0/applications?$filter=appId eq '{app_id}'", headers=headers)
            if resp.status_code == 200 and resp.json()['value']:
                obj_id = resp.json()['value'][0]['id']
                requests.delete(f"{graph_endpoint}/v1.0/applications/{obj_id}", headers=headers)
                self.log(f"√ App {app_id} 已删除")
                self.app_id_var.set("")
                self.tenant_id_var.set("")
                self.thumbprint_var.set("")
                self.save_config()
                messagebox.showinfo("成功", "App 已删除。")
            else:
                self.log(f"X 未找到 App {app_id}")
                messagebox.showwarning("警告", "未找到该 App，可能已被删除。")

        except Exception as e:
            self.log(f"X 删除失败: {e}", "ERROR")
            messagebox.showerror("错误", str(e))

    def test_ews_connection(self):
        if EXCHANGELIB_ERROR:
            self.log(f"EWS 模块加载失败: {EXCHANGELIB_ERROR}", level="ERROR")
            messagebox.showerror("错误", f"无法加载 EWS 模块 (exchangelib)。\n错误信息: {EXCHANGELIB_ERROR}")
            return
        threading.Thread(target=self._test_ews, daemon=True).start()

    def _clean_server_address(self, server_input):
        if not server_input: return server_input
        # Remove protocol
        if server_input.lower().startswith("http://"):
            server_input = server_input[7:]
        elif server_input.lower().startswith("https://"):
            server_input = server_input[8:]
        
        # Remove path
        if "/" in server_input:
            server_input = server_input.split("/")[0]
            
        return server_input.strip()

    def _normalize_date_input(self, date_str):
        """
        Normalize date input to YYYY-MM-DD format.
        Supports: YYYY/MM/DD, YYYYMMDD, YYYY-MM-DD
        """
        if not date_str:
            return None
        
        date_str = date_str.strip()
        try:
            # Try YYYY-MM-DD
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass
        
        try:
            # Try YYYY/MM/DD
            dt = datetime.strptime(date_str, "%Y/%m/%d")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass
            
        try:
            # Try YYYYMMDD
            dt = datetime.strptime(date_str, "%Y%m%d")
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass
            
        return date_str # Return original if parse fails

    def _test_ews(self):
        try:
            self.log(">>> 正在测试 EWS 连接...")
            user = self.ews_user_var.get()
            pwd = self.ews_pass_var.get()
            server = self._clean_server_address(self.ews_server_var.get())
            use_auto = self.ews_use_autodiscover.get()

            if not user or not pwd:
                raise Exception("需要用户名和密码。")

            credentials = Credentials(user, pwd)
            
            if use_auto:
                self.log("Using Autodiscover...")
                account = Account(primary_smtp_address=user, credentials=credentials, autodiscover=True)
            else:
                if not server: raise Exception("Server URL required if Autodiscover is off.")
                self.log(f"Connecting to server: {server}")
                config = Configuration(server=server, credentials=credentials)
                account = Account(primary_smtp_address=user, config=config, autodiscover=False)

            self.log(f"√ Connection Successful! Server: {account.protocol.service_endpoint}")
            self.save_config()
        except Exception as e:
            self.log(f"X Connection Failed: {e}", "ERROR")

    # --- Tab 3: Cleanup ---
    def build_cleanup_tab(self):
        frame = ttk.Frame(self.tab_cleanup, padding=10)
        frame.pack(fill="both", expand=True)

        # Source Selection
        src_frame = ttk.LabelFrame(frame, text="源系统 & 目标")
        src_frame.pack(fill="x", pady=5)
        
        ttk.Label(src_frame, text="源系统:").pack(side="left", padx=5)
        ttk.Radiobutton(src_frame, text="Graph API", variable=self.source_type_var, value="Graph").pack(side="left", padx=5)
        ttk.Radiobutton(src_frame, text="Exchange EWS", variable=self.source_type_var, value="EWS").pack(side="left", padx=5)
        
        ttk.Label(src_frame, text="| 目标用户 CSV:").pack(side="left", padx=5)
        ttk.Entry(src_frame, textvariable=self.csv_path_var, width=50).pack(side="left", padx=5)
        ttk.Button(src_frame, text="浏览...", command=lambda: self.csv_path_var.set(filedialog.askopenfilename(filetypes=[("CSV", "*.csv")]))).pack(side="left")

        # Target Selection
        target_frame = ttk.LabelFrame(frame, text="清理对象类型")
        target_frame.pack(fill="x", pady=5)
        ttk.Radiobutton(target_frame, text="邮件 (Email)", variable=self.cleanup_target_var, value="Email", command=self.update_ui_for_target).pack(side="left", padx=10)
        ttk.Radiobutton(target_frame, text="会议 (Meeting)", variable=self.cleanup_target_var, value="Meeting", command=self.update_ui_for_target).pack(side="left", padx=10)
        
        # Meeting Options
        self.meeting_opt_frame = ttk.LabelFrame(frame, text="会议特定选项")
        # Pack later if needed or pack and hide
        
        ttk.Label(self.meeting_opt_frame, text="循环类型:").pack(side="left", padx=5)
        ttk.Combobox(self.meeting_opt_frame, textvariable=self.meeting_scope_var, values=["所有 (All)", "仅单次 (Single Instance)", "仅系列主会议 (Series Master)"], state="readonly", width=25).pack(side="left", padx=5)
        
        ttk.Checkbutton(self.meeting_opt_frame, text="仅处理已取消 (IsCancelled Only)", variable=self.meeting_only_cancelled_var).pack(side="left", padx=15)

        # Criteria
        self.filter_frame = ttk.LabelFrame(frame, text="搜索条件 (留空则忽略)")
        self.filter_frame.pack(fill="x", pady=5)
        
        grid_opts = {'padx': 5, 'pady': 2, 'sticky': 'w'}
        
        ttk.Label(self.filter_frame, text="Message ID:").grid(row=0, column=0, **grid_opts)
        ttk.Entry(self.filter_frame, textvariable=self.criteria_msg_id, width=30).grid(row=0, column=1, **grid_opts)
        
        self.lbl_subject = ttk.Label(self.filter_frame, text="主题包含:")
        self.lbl_subject.grid(row=1, column=0, **grid_opts)
        ttk.Entry(self.filter_frame, textvariable=self.criteria_subject, width=30).grid(row=1, column=1, **grid_opts)
        
        self.lbl_sender = ttk.Label(self.filter_frame, text="发件人地址:")
        self.lbl_sender.grid(row=1, column=2, **grid_opts)
        ttk.Entry(self.filter_frame, textvariable=self.criteria_sender, width=30).grid(row=1, column=3, **grid_opts)

        ttk.Label(self.filter_frame, text="开始日期 (YYYY-MM-DD):").grid(row=2, column=0, **grid_opts)
        ttk.Entry(self.filter_frame, textvariable=self.criteria_start_date, width=30).grid(row=2, column=1, **grid_opts)
        
        ttk.Label(self.filter_frame, text="结束日期 (YYYY-MM-DD):").grid(row=2, column=2, **grid_opts)
        ttk.Entry(self.filter_frame, textvariable=self.criteria_end_date, width=30).grid(row=2, column=3, **grid_opts)

        self.lbl_body = ttk.Label(self.filter_frame, text="正文包含:")
        self.lbl_body.grid(row=3, column=0, **grid_opts)
        ttk.Entry(self.filter_frame, textvariable=self.criteria_body, width=80).grid(row=3, column=1, columnspan=3, **grid_opts)

        self.update_ui_for_target() # Init state

        # Options
        opt_frame = ttk.LabelFrame(frame, text="执行选项")
        opt_frame.pack(fill="x", pady=5)
        
        ttk.Checkbutton(opt_frame, text="仅报告 (不删除)", variable=self.report_only_var).pack(side="left", padx=10)
        
        ttk.Label(opt_frame, text="| 日志级别:").pack(side="left", padx=5)
        
        def on_log_level_click():
            val = self.log_level_var.get()
            if val == "Expert":
                confirm = messagebox.askyesno("警告", "日志排错专用，日志量会很大且包含敏感信息，慎选！\n\n确认开启专家模式吗？")
                if not confirm:
                    self.log_level_var.set("Normal")

        ttk.Radiobutton(opt_frame, text="默认 (Default)", variable=self.log_level_var, value="Normal", command=on_log_level_click).pack(side="left", padx=5)
        ttk.Radiobutton(opt_frame, text="高级 (Advanced)", variable=self.log_level_var, value="Advanced", command=on_log_level_click).pack(side="left", padx=5)
        ttk.Radiobutton(opt_frame, text="专家 (Expert)", variable=self.log_level_var, value="Expert", command=on_log_level_click).pack(side="left", padx=5)

        # Start
        ttk.Button(frame, text="开始清理任务", command=self.start_cleanup_thread).pack(pady=10, ipadx=20, ipady=5)

    def update_ui_for_target(self):
        target = self.cleanup_target_var.get()
        if target == "Meeting":
            self.meeting_opt_frame.pack(fill="x", pady=5, after=self.filter_frame) # Pack below filter or above? Let's put it above filter
            self.meeting_opt_frame.pack(fill="x", pady=5, before=self.filter_frame)
            
            self.lbl_subject.config(text="会议标题包含:")
            self.lbl_sender.config(text="组织者地址:")
            self.lbl_body.config(text="会议内容包含:")
        else:
            self.meeting_opt_frame.pack_forget()
            self.lbl_subject.config(text="邮件主题包含:")
            self.lbl_sender.config(text="发件人地址:")
            self.lbl_body.config(text="邮件正文包含:")

    def start_cleanup_thread(self):
        if not self.csv_path_var.get():
            messagebox.showerror("错误", "请选择 CSV 文件。")
            return
        
        # Validation
        source = self.source_type_var.get()
        if source == "Graph":
            mode = self.graph_auth_mode_var.get()
            if mode == "Auto":
                if not self.app_id_var.get() or not self.tenant_id_var.get() or not self.thumbprint_var.get():
                    messagebox.showwarning("配置缺失", "您选择了 Graph API (自动/证书) 模式，但未配置 App ID, Tenant ID 或 Thumbprint。\n请前往 '1. 连接配置' 标签页进行配置。")
                    self.notebook.select(self.tab_connection)
                    return
            else: # Manual
                if not self.app_id_var.get() or not self.tenant_id_var.get() or not self.client_secret_var.get():
                    messagebox.showwarning("配置缺失", "您选择了 Graph API (手动/Secret) 模式，但未配置 App ID, Tenant ID 或 Client Secret。\n请前往 '1. 连接配置' 标签页进行配置。")
                    self.notebook.select(self.tab_connection)
                    return

        elif source == "EWS":
            if not self.ews_user_var.get() or not self.ews_pass_var.get():
                messagebox.showwarning("配置缺失", "您选择了 EWS 模式，但未配置用户名或密码。\n请前往 '1. 连接配置' 标签页进行配置。")
                self.notebook.select(self.tab_connection)
                return
            if not self.ews_use_autodiscover.get() and not self.ews_server_var.get():
                messagebox.showwarning("配置缺失", "您选择了 EWS 模式且未启用自动发现，但未配置服务器地址。\n请前往 '1. 连接配置' 标签页进行配置。")
                self.notebook.select(self.tab_connection)
                return

        self.logger.set_level(self.log_level_var.get().upper())
        self.save_config()
        
        threading.Thread(target=self.run_cleanup, daemon=True).start()

    def run_cleanup(self):
        source = self.source_type_var.get()
        if source == "Graph":
            self.run_graph_cleanup()
        else:
            self.run_ews_cleanup()

    # --- Helper Methods ---
    def run_powershell_script(self, script):
        wrapped_script = f"""
        $ErrorActionPreference = 'Stop'
        try {{
            {script}
        }} catch {{
            Write-Error $_
            exit 1
        }}
        """
        command = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", wrapped_script]
        creation_flags = 0x08000000 if sys.platform == 'win32' else 0
        process = subprocess.run(command, capture_output=True, text=True, creationflags=creation_flags)
        if process.returncode != 0:
            raise Exception(f"PowerShell Error: {process.stderr}")
        return process.stdout.strip()

    def get_token_from_secret(self, tenant_id, client_id, client_secret, env):
        authority_host = "https://login.chinacloudapi.cn" if env == "China" else "https://login.microsoftonline.com"
        scope = "https://microsoftgraph.chinacloudapi.cn/.default" if env == "China" else "https://graph.microsoft.com/.default"
        token_url = f"{authority_host}/{tenant_id}/oauth2/v2.0/token"
        
        data = {
            'client_id': client_id,
            'scope': scope,
            'client_secret': client_secret,
            'grant_type': 'client_credentials'
        }
        
        resp = requests.post(token_url, data=data)
        if resp.status_code == 200:
            return resp.json().get('access_token')
        else:
            raise Exception(f"获取 Token 失败: {resp.text}")

    def get_token_from_cert(self, tenant_id, client_id, thumbprint, env):
        authority = "https://login.chinacloudapi.cn" if env == "China" else "https://login.microsoftonline.com"
        scope = "https://microsoftgraph.chinacloudapi.cn/.default" if env == "China" else "https://graph.microsoft.com/.default"
        
        ps_template = r"""
        $ErrorActionPreference = 'Stop'
        try {
            $Thumbprint = "__THUMBPRINT__"
            $TenantId = "__TENANTID__"
            $ClientId = "__CLIENTID__"
            $Scope = "__SCOPE__"
            $Authority = "__AUTHORITY__/$TenantId/oauth2/v2.0/token"

            $Cert = Get-Item "Cert:\CurrentUser\My\$Thumbprint"
            if (-not $Cert) { throw "Certificate with thumbprint $Thumbprint not found in CurrentUser\My" }
            
            function ConvertTo-Base64UrlString([byte[]]$Bytes) {
                [Convert]::ToBase64String($Bytes).Split('=')[0].Replace('+', '-').Replace('/', '_')
            }

            $HeaderTemplate = '{{ "alg": "RS256", "x5t": "{0}", "typ": "JWT" }}'
            $HeaderJson = $HeaderTemplate -f [Convert]::ToBase64String($Cert.GetCertHash())
            $Header = [System.Text.Encoding]::UTF8.GetBytes($HeaderJson)
            $HeaderStr = ConvertTo-Base64UrlString $Header

            $Now = [Math]::Floor([decimal](Get-Date).ToUniversalTime().Subtract([datetime]'1970-01-01').TotalSeconds)
            $Exp = $Now + 300
            
            $PayloadTemplate = '{{ "aud": "{0}", "exp": {1}, "iss": "{2}", "jti": "{3}", "nbf": {4}, "sub": "{2}" }}'
            $PayloadJson = $PayloadTemplate -f $Authority, $Exp, $ClientId, [Guid]::NewGuid(), $Now
            $Payload = [System.Text.Encoding]::UTF8.GetBytes($PayloadJson)
            $PayloadStr = ConvertTo-Base64UrlString $Payload

            $ToSign = [System.Text.Encoding]::UTF8.GetBytes("$HeaderStr.$PayloadStr")
            
            $RSACrypto = $Cert.PrivateKey
            if (-not $RSACrypto) { throw "Private key not accessible for certificate" }

            $Signature = $null
            if ($RSACrypto.GetType().Name -match "RSACryptoServiceProvider") {
                $Signature = $RSACrypto.SignData($ToSign, "SHA256")
            } else {
                $Signature = $RSACrypto.SignData($ToSign, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            }

            $SignatureStr = ConvertTo-Base64UrlString $Signature
            $JWT = "$HeaderStr.$PayloadStr.$SignatureStr"

            $Body = @{
                client_id = $ClientId
                scope = $Scope
                client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                client_assertion = $JWT
                grant_type = "client_credentials"
            }

            $Response = Invoke-RestMethod -Method Post -Uri $Authority -Body $Body
            Write-Output $Response.access_token
        } catch {
            Write-Error $_
            exit 1
        }
        """
        
        script = ps_template.replace("__THUMBPRINT__", thumbprint)\
                            .replace("__TENANTID__", tenant_id)\
                            .replace("__CLIENTID__", client_id)\
                            .replace("__SCOPE__", scope)\
                            .replace("__AUTHORITY__", authority)
                            
        return self.run_powershell_script(script)

    # --- Graph Logic ---
    def run_graph_cleanup(self):
        try:
            app_id = self.app_id_var.get()
            tenant_id = self.tenant_id_var.get()
            thumbprint = self.thumbprint_var.get()
            client_secret = self.client_secret_var.get()
            env = self.graph_env_var.get()
            auth_mode = self.graph_auth_mode_var.get()
            
            graph_endpoint = "https://microsoftgraph.chinacloudapi.cn" if env == "China" else "https://graph.microsoft.com"

            self.log(">>> 正在获取 Access Token...")
            
            if auth_mode == "Auto":
                token = self.get_token_from_cert(tenant_id, app_id, thumbprint, env)
            else:
                token = self.get_token_from_secret(tenant_id, app_id, client_secret, env)
                
            if not token: raise Exception("获取 Token 失败")
            
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            self.log("√ Token 获取成功")

            # Read CSV
            users = []
            with open(self.csv_path_var.get(), 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # 兼容带 BOM 或不带 BOM 的 key
                    key = next((k for k in row.keys() if 'UserPrincipalName' in k), None)
                    if key and row[key]:
                        users.append(row[key].strip())
            
            self.log(f">>> 找到 {len(users)} 个用户")

            # Report File
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            report_path = os.path.join(self.reports_dir, f"Graph_Report_{timestamp}.csv")
            self.update_report_link(report_path)
            
            target_type = self.cleanup_target_var.get()
            
            with open(report_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = ['UserPrincipalName', 'ItemId', 'Subject', 'Sender/Organizer', 'Time', 'Type', 'Action', 'Status', 'Details']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()

                # Build Filter
                filters = []
                
                # Common Filters
                if self.criteria_subject.get(): filters.append(f"contains(subject, '{self.criteria_subject.get()}')")
                
                start_date = self.criteria_start_date.get().strip().replace('/', '-')
                end_date = self.criteria_end_date.get().strip().replace('/', '-')
                
                # Target Specific Logic
                if target_type == "Email":
                    resource = "messages"
                    if self.criteria_msg_id.get(): filters.append(f"internetMessageId eq '{self.criteria_msg_id.get()}'")
                    if self.criteria_sender.get(): filters.append(f"from/emailAddress/address eq '{self.criteria_sender.get()}'")
                    if start_date: filters.append(f"receivedDateTime ge {start_date}T00:00:00Z")
                    if end_date: filters.append(f"receivedDateTime le {end_date}T23:59:59Z")
                else: # Meeting
                    resource = "events" # or calendar/events
                    if self.criteria_sender.get(): filters.append(f"organizer/emailAddress/address eq '{self.criteria_sender.get()}'")
                    if start_date: filters.append(f"start/dateTime ge '{start_date}T00:00:00'")
                    if end_date: filters.append(f"end/dateTime le '{end_date}T23:59:59'")
                    
                    # Meeting Specifics
                    if self.meeting_only_cancelled_var.get():
                        filters.append("isCancelled eq true")
                    
                    scope = self.meeting_scope_var.get()
                    if "Single" in scope:
                        filters.append("type eq 'singleInstance'")
                    elif "Series" in scope:
                        filters.append("type eq 'seriesMaster'")
                    # If All, no type filter

                filter_str = " and ".join(filters)
                body_keyword = self.criteria_body.get()

                for user in users:
                    self.log(f"--- 正在处理: {user} ---")
                    try:
                        url = f"{graph_endpoint}/v1.0/users/{user}/{resource}"
                        
                        if target_type == "Email":
                            params = {"$top": 100, "$select": "id,subject,from,receivedDateTime,createdDateTime,body"}
                        else:
                            params = {"$top": 100, "$select": "id,subject,organizer,start,type,isCancelled,body"}

                        if filter_str: params["$filter"] = filter_str
                        
                        if body_keyword:
                            params["$search"] = f'"body:{body_keyword}"'
                            headers["ConsistencyLevel"] = "eventual"
                        
                        while url:
                            if self.log_level_var.get() == "Advanced":
                                self.logger.log_to_file_only(f"GRAPH REQ: GET {url}")
                                self.logger.log_to_file_only(f"HEADERS: {json.dumps(headers, default=str)}")
                                if params: self.logger.log_to_file_only(f"PARAMS: {json.dumps(params, default=str)}")

                            self.log(f"请求: GET {url} | 参数: {params}", is_advanced=True)
                            resp = requests.get(url, headers=headers, params=params if "users" in url and "?" not in url else None) # Simple check to avoid double params
                            
                            if self.log_level_var.get() == "Advanced":
                                self.logger.log_to_file_only(f"GRAPH RESP: {resp.status_code}")
                                self.logger.log_to_file_only(f"HEADERS: {json.dumps(dict(resp.headers), default=str)}")
                                self.logger.log_to_file_only(f"BODY: {resp.text}")
                            
                            if resp.status_code != 200:
                                self.log(f"  X 查询失败: {resp.text}", "ERROR")
                                self.log(f"响应: {resp.text}", is_advanced=True)
                                writer.writerow({'UserPrincipalName': user, 'Status': 'Error', 'Details': resp.text})
                                break
                            
                            data = resp.json()
                            items = data.get('value', [])
                            
                            if not items:
                                self.log("  未找到匹配项。")
                                break

                            for item in items:
                                should_delete = True
                                if body_keyword and "$search" not in params:
                                    content = item.get('body', {}).get('content', '')
                                    if body_keyword.lower() not in content.lower():
                                        should_delete = False

                                if should_delete:
                                    item_id = item['id']
                                    subject = item.get('subject', '无主题')
                                    
                                    if target_type == "Email":
                                        sender = item.get('from', {}).get('emailAddress', {}).get('address', '未知')
                                        time_val = item.get('receivedDateTime')
                                        item_type = "Email"
                                    else:
                                        sender = item.get('organizer', {}).get('emailAddress', {}).get('address', '未知')
                                        time_val = item.get('start', {}).get('dateTime')
                                        item_type = item.get('type', 'Event')
                                        if item.get('isCancelled'): item_type += " (Cancelled)"

                                    row_data = {
                                        'UserPrincipalName': user,
                                        'ItemId': item_id,
                                        'Subject': subject,
                                        'Sender/Organizer': sender,
                                        'Time': time_val,
                                        'Type': item_type,
                                        'Action': 'ReportOnly' if report_only else 'Delete',
                                        'Status': 'Pending',
                                        'Details': ''
                                    }

                                    if report_only:
                                        self.log(f"  [报告] 发现: {subject} ({item_type})")
                                        row_data['Status'] = 'Skipped'
                                        row_data['Details'] = '仅报告模式'
                                    else:
                                        self.log(f"  正在删除: {subject}")
                                        del_url = f"{graph_endpoint}/v1.0/users/{user}/{resource}/{item_id}"
                                        
                                        if self.log_level_var.get() == "Advanced":
                                            self.logger.log_to_file_only(f"GRAPH REQ: DELETE {del_url}")
                                            self.logger.log_to_file_only(f"HEADERS: {json.dumps(headers, default=str)}")

                                        self.log(f"请求: DELETE {del_url}", is_advanced=True)
                                        del_resp = requests.delete(del_url, headers=headers)
                                        
                                        if self.log_level_var.get() == "Advanced":
                                            self.logger.log_to_file_only(f"GRAPH RESP: {del_resp.status_code}")
                                            self.logger.log_to_file_only(f"BODY: {del_resp.text}")
                                        
                                        if del_resp.status_code == 204:
                                            self.log("    √ 已删除")
                                            row_data['Status'] = 'Success'
                                        else:
                                            self.log(f"    X 删除失败: {del_resp.status_code}", "ERROR")
                                            self.log(f"响应: {del_resp.text}", is_advanced=True)
                                            row_data['Status'] = 'Failed'
                                            row_data['Details'] = f"状态码: {del_resp.status_code}"
                                    
                                    writer.writerow(row_data)
                                    csvfile.flush()

                            url = data.get('@odata.nextLink')
                            # Reset params for next link as they are usually included
                            params = None 
                            
                    except Exception as ue:
                        self.log(f"  X 处理用户出错: {ue}", "ERROR")
                        writer.writerow({'UserPrincipalName': user, 'Status': 'Error', 'Details': str(ue)})
                            
                    except Exception as ue:
                        self.log(f"  X 处理用户出错: {ue}", "ERROR")
                        writer.writerow({'UserPrincipalName': user, 'Status': 'Error', 'Details': str(ue)})

            self.log(f">>> 任务完成! 报告: {report_path}")
            msg_title = "完成"
            if self.report_only_var.get():
                msg_body = f"扫描生成报告任务完成。\n报告: {report_path}"
            else:
                msg_body = f"清理任务已完成。\n报告: {report_path}"
                
            messagebox.showinfo(msg_title, msg_body)

        except Exception as e:
            self.log(f"X 运行时错误: {e}", "ERROR")
            messagebox.showerror("错误", str(e))
        finally:
            pass

    # --- EWS Logic ---
    def run_ews_cleanup(self):
        if EXCHANGELIB_ERROR:
            self.log(f"EWS 模块加载失败: {EXCHANGELIB_ERROR}", level="ERROR")
            messagebox.showerror("错误", f"无法加载 EWS 模块 (exchangelib)。\n错误信息: {EXCHANGELIB_ERROR}")
            return

        # Configure Advanced Logging for EWS
        ews_log_handlers = []
        log_level = self.log_level_var.get()
        
        if log_level in ("Advanced", "Expert"):
            try:
                # Create Handler
                file_handler = logging.FileHandler(self.log_file_path, encoding='utf-8')
                # We use a simple formatter because the TraceAdapter handles the XML formatting
                file_handler.setFormatter(logging.Formatter('%(message)s')) 
                ews_log_handlers.append(file_handler)

                # Setup Trace Logger
                trace_logger = logging.getLogger("EWS_TRACE")
                trace_logger.setLevel(logging.DEBUG)
                trace_logger.addHandler(file_handler)
                
                # Inject Adapter
                EwsTraceAdapter.logger = trace_logger
                EwsTraceAdapter.log_responses = (log_level == "Expert")
                BaseProtocol.HTTP_ADAPTER_CLS = EwsTraceAdapter
                
                # Check permission for response log if Expert
                if log_level == "Expert":
                    try:
                        docs_dir = os.path.join(os.path.expanduser("~"), "Documents", "UniversalEmailCleaner")
                        os.makedirs(docs_dir, exist_ok=True)
                        test_path = os.path.join(docs_dir, "ews_getitem_responses.log")
                        with open(test_path, "a", encoding="utf-8") as f:
                            pass
                        self.log(f"EWS 响应日志将写入: {test_path}", is_advanced=True)
                    except Exception as e:
                        self.log(f"警告: 无法写入响应日志文件: {e}", "ERROR")
                    
            except Exception as e:
                self.log(f"无法启用 EWS 调试日志: {e}", "ERROR")
        else:
            # Reset to default if not advanced/expert
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

        try:
            self.log(">>> 开始 EWS 清理...")
            
            # 1. Connect
            user = self.ews_user_var.get()
            pwd = self.ews_pass_var.get()
            server = self._clean_server_address(self.ews_server_var.get())
            use_auto = self.ews_use_autodiscover.get()
            auth_type = self.ews_auth_type_var.get()

            creds = Credentials(user, pwd)
            config = None
            if not use_auto:
                self.log(f"Connecting to server: {server}")
                config = Configuration(server=server, credentials=creds)

            # 2. Read CSV
            users = []
            with open(self.csv_path_var.get(), 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # 兼容带 BOM 或不带 BOM 的 key
                    key = next((k for k in row.keys() if 'UserPrincipalName' in k), None)
                    if key and row[key]:
                        users.append(row[key].strip())
            
            self.log(f"在 CSV 中找到 {len(users)} 个用户。")

            # 3. Report File
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            report_path = os.path.join(self.reports_dir, f"EWS_Report_{timestamp}.csv")
            self.update_report_link(report_path)

            # Determine headers based on target type
            target_type = self.cleanup_target_var.get()
            if target_type == "Meeting":
                fieldnames = [
                    'UserPrincipalName', 'Subject', 'Type', 'MeetingGOID', 'CleanGOID', 
                    'Organizer', 'Attendees', 'Start', 'End', 'UserRole', 
                    'IsCancelled', 'ResponseStatus', 'RecurrencePattern', 'PatternDetails', 'RecurrenceDuration', 'IsEndless',
                    'Action', 'Status', 'Details'
                ]
            else:
                fieldnames = ['UserPrincipalName', 'MessageId', 'Subject', 'Sender', 'Received', 'Action', 'Status', 'Details']

            with open(report_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()

                for target_email in users:
                    self.log(f"--- 正在处理: {target_email} ---")
                    try:
                        # Impersonation Setup
                        if auth_type == "Impersonation":
                            if use_auto:
                                account = Account(primary_smtp_address=target_email, credentials=creds, autodiscover=True, access_type=IMPERSONATION)
                            else:
                                account = Account(primary_smtp_address=target_email, config=config, autodiscover=False, access_type=IMPERSONATION)
                        else:
                            # Delegate / Direct
                            if use_auto:
                                account = Account(primary_smtp_address=target_email, credentials=creds, autodiscover=True, access_type=DELEGATE)
                            else:
                                account = Account(primary_smtp_address=target_email, config=config, autodiscover=False, access_type=DELEGATE)

                        self.log(f"已连接到邮箱: {target_email}", is_advanced=True)

                        # Build Query
                        
                        # Date Parsing (Pre-calculation)
                        start_date_str = self._normalize_date_input(self.criteria_start_date.get())
                        end_date_str = self._normalize_date_input(self.criteria_end_date.get())
                        
                        start_dt = None
                        end_dt = None
                        
                        if start_date_str:
                            self.criteria_start_date.set(start_date_str)
                            dt = datetime.strptime(start_date_str, "%Y-%m-%d")
                            start_dt = EWSDateTime.from_datetime(dt).replace(tzinfo=account.default_timezone)
                            
                        if end_date_str:
                            self.criteria_end_date.set(end_date_str)
                            dt = datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1)
                            end_dt = EWSDateTime.from_datetime(dt).replace(tzinfo=account.default_timezone)

                        # Recurrence Cache
                        recurrence_cache = {}
                        # Track UIDs we need to fetch masters for
                        uids_to_fetch = set()

                        if target_type == "Email":
                            qs = account.inbox.all()
                            if start_dt:
                                qs = qs.filter(datetime_received__gte=start_dt)
                            if end_dt:
                                qs = qs.filter(datetime_received__lt=end_dt)
                            if self.criteria_sender.get():
                                qs = qs.filter(sender__icontains=self.criteria_sender.get())
                        else:
                            # Meeting Logic with CalendarView
                            if start_dt or end_dt:
                                # View requires both start and end. Default if missing.
                                # Default range: 1900 to 2100 if not specified
                                view_start = start_dt if start_dt else EWSDateTime(1900, 1, 1, tzinfo=account.default_timezone)
                                view_end = end_dt if end_dt else EWSDateTime(2100, 1, 1, tzinfo=account.default_timezone)
                                
                                self.log(f"使用日历视图 (CalendarView) 展开循环会议: {view_start} -> {view_end}", is_advanced=True)
                                # Some fields cause InvalidField errors in CalendarView.only().
                                # Use plain view() and enrich per-item via GetItem.
                                qs = account.calendar.view(start=view_start, end=view_end)
                               
                            else:
                                self.log("未指定日期范围，使用普通查询 (不展开循环会议实例)", is_advanced=True)
                                qs = account.calendar.all()

                            if self.criteria_sender.get():
                                qs = qs.filter(organizer__icontains=self.criteria_sender.get())
                            
                            # Meeting Specifics
                            if self.meeting_only_cancelled_var.get():
                                qs = qs.filter(is_cancelled=True)

                        # Common Filters
                        # Note: CalendarView does not support server-side filtering (Restrictions).
                        # If using CalendarView (Meeting with date range), we must apply filters client-side.
                        is_calendar_view = (target_type == "Meeting" and (start_dt or end_dt))

                        if not is_calendar_view:
                            if self.criteria_msg_id.get():
                                qs = qs.filter(message_id=self.criteria_msg_id.get())
                            if self.criteria_subject.get():
                                qs = qs.filter(subject__icontains=self.criteria_subject.get())
                        
                        # Body (Client-side filter for EWS usually, or use full text search if supported)
                        # EWS 'body__icontains' can be slow or restricted.
                        
                        self.log(f"正在查询 EWS...", is_advanced=True)
                        
                        # Note: Avoid using only() here because some field paths are
                        # not available in CalendarView / folder contexts and raise errors.
                        # We'll access fields safely via getattr when present.
                        if target_type == "Meeting":
                            pass

                        items = list(qs) # Execute query
                        if not items:
                            self.log("未找到项目。")
                        
                        for item in items:
                            # Client Side Filters
                            if target_type == "Meeting":
                                # 1. Filter by Subject (if CalendarView)
                                if is_calendar_view and self.criteria_subject.get():
                                    if self.criteria_subject.get().lower() not in (item.subject or "").lower():
                                        continue
                                
                                # 2. Filter by Organizer (if CalendarView)
                                if is_calendar_view and self.criteria_sender.get():
                                    organizer_email = item.organizer.email_address if item.organizer else ""
                                    if self.criteria_sender.get().lower() not in organizer_email.lower():
                                        continue

                                # 3. Filter by IsCancelled (if CalendarView)
                                if is_calendar_view and self.meeting_only_cancelled_var.get():
                                    if not item.is_cancelled:
                                        continue

                                scope = self.meeting_scope_var.get()
                                inferred_type = guess_calendar_item_type(item)

                                # Apply scope filter using inferred type
                                if "Single" in scope:
                                    if inferred_type != 'Single': continue
                                elif "Series" in scope:
                                    if inferred_type not in ('RecurringMaster', 'Occurrence', 'Exception'): continue
                                # If "All", pass all items

                            # Body check if needed
                            if self.criteria_body.get():
                                if self.criteria_body.get().lower() not in (item.body or "").lower():
                                    continue

                            # Enrich meeting item details (GetItem) to ensure key fields are populated
                            if target_type == "Meeting":
                                try:
                                    # CalendarView() often returns shallow items. Re-fetch the full item to ensure key fields exist.
                                    _id = getattr(item, 'id', None) or getattr(item, 'item_id', None)
                                    _ck = getattr(item, 'changekey', None) or getattr(item, 'change_key', None) or getattr(item, 'changeKey', None)
                                    if _id:
                                        # Fetch full item by id/changekey, then refresh to load full properties
                                        full_item = account.calendar.get(id=_id, changekey=_ck) if _ck else account.calendar.get(id=_id)
                                        try:
                                            full_item.refresh()  # Trigger GetItem to load additional fields
                                        except Exception:
                                            pass

                                        if full_item:
                                            item = full_item
                                        # Advanced debug: dump key fields including guess result and raw fields
                                        if self.log_level_var.get() == "Advanced":
                                            guess_type = guess_calendar_item_type(item)
                                            ci_type = getattr(item, 'calendar_item_type', None)
                                            ci_uid = getattr(item, 'uid', None)
                                            ci_master = getattr(item, 'recurring_master_id', None)
                                            is_rec = getattr(item, 'is_recurring', None)
                                            rec = getattr(item, 'recurrence', None)
                                            rec_id = getattr(item, 'recurrence_id', None)
                                            orig_start = getattr(item, 'original_start', None)
                                            py_type = type(item)
                                            class_name = item.__class__.__name__
                                            item_class = getattr(item, 'item_class', None)
                                            self.log(
                                                f"[EWS] ItemFields -> "
                                                f"GuessType={guess_type} "
                                                f"RawType={ci_type} "
                                                f"UID={ci_uid} "
                                                f"IsRecurring={is_rec} "
                                                f"RecMasterId={ci_master} "
                                                f"RecurrenceId={rec_id} "
                                                f"OriginalStart={orig_start} "
                                                f"HasRecurrence={rec is not None} "
                                                f"PyType={py_type} "
                                                f"ClassName={class_name} "
                                                f"ItemClass={item_class}",
                                                is_advanced=True
                                            )
                                except Exception as _e:
                                    # Continue with shallow item if enrichment fails
                                    self.log(f"  无法获取完整项属性 (GetItem): {_e}", is_advanced=True)

                            # Extract attributes safely (CalendarItem vs Message)
                            item_id = getattr(item, 'id', None) or (item.item_id if hasattr(item, 'item_id') else getattr(item, 'message_id', 'Unknown ID'))
                            subject = item.subject
                            
                            row = {}
                            if target_type == "Meeting":
                                # (1) Master vs Instance
                                has_recurrence = getattr(item, 'recurrence', None) is not None
                                has_instance_markers = (getattr(item, 'original_start', None) is not None) or (getattr(item, 'recurrence_id', None) is not None)
                                
                                if has_recurrence:
                                    m_type = "RecurringMaster"
                                elif has_instance_markers:
                                    m_type = "Instance"
                                else:
                                    # Fall back to guess if neither markers are present
                                    m_type = guess_calendar_item_type(item)

                                # Get UID and RecurringMasterId
                                m_uid = getattr(item, 'uid', '')
                                m_recurring_master_id = getattr(item, 'recurring_master_id', None)
                                
                                # GOID / CleanGOID
                                m_goid = m_uid
                                m_clean_goid = m_goid 

                                # Organizer
                                m_organizer = item.organizer.email_address if item.organizer else 'Unknown'
                                
                                # Attendees
                                m_attendees = []
                                if item.required_attendees:
                                    m_attendees.extend([a.mailbox.email_address for a in item.required_attendees if a.mailbox])
                                if item.optional_attendees:
                                    m_attendees.extend([a.mailbox.email_address for a in item.optional_attendees if a.mailbox])
                                m_attendees_str = "; ".join(m_attendees)
                                
                                # Start / End
                                m_start = getattr(item, 'start', '')
                                m_end = getattr(item, 'end', '')
                                
                                # User Role
                                # Check if target_email is organizer
                                m_role = 'Attendee'
                                if m_organizer.lower() == target_email.lower():
                                    m_role = 'Organizer'
                                
                                # IsCancelled
                                m_is_cancelled = getattr(item, 'is_cancelled', False)
                                
                                # Response Status
                                m_response_status = getattr(item, 'my_response_type', 'Unknown')
                                
                                # (2) If instance, determine Occurrence vs Exception
                                original_start = getattr(item, 'original_start', None)
                                if m_type == "Instance":
                                    # Definite exception when start != original_start
                                    start_val = getattr(item, 'start', None)
                                    if start_val and original_start and start_val != original_start:
                                        m_type = "Exception"
                                    else:
                                        # Try to fetch master and compare subject/location/attendees
                                        master_item = None
                                        try:
                                            # Prefer recurring_master_id
                                            master_id = m_recurring_master_id
                                            master_ck = None
                                            if master_id is not None:
                                                if hasattr(master_id, "id"):
                                                    master_ck = getattr(master_id, "changekey", None)
                                                    master_id = master_id.id
                                                master_item = account.calendar.get(id=master_id, changekey=master_ck) if master_ck else account.calendar.get(id=master_id)
                                            # Fallback by UID
                                            if master_item is None and m_uid:
                                                for m in account.calendar.all():
                                                    if getattr(m, 'uid', '') == m_uid and getattr(m, 'recurrence', None):
                                                        master_item = m
                                                        break
                                        except Exception as e:
                                            self.log(f"  查找主项用于比较失败: {e}", is_advanced=True)

                                        if master_item:
                                            try:
                                                subj_diff = (item.subject or '') != (master_item.subject or '')
                                                loc_diff = (getattr(item, 'location', None) or '') != (getattr(master_item, 'location', None) or '')
                                                # Compare attendees (basic string set compare)
                                                def attendees_set(it):
                                                    s = set()
                                                    if getattr(it, 'required_attendees', None):
                                                        s.update([a.mailbox.email_address for a in it.required_attendees if a.mailbox])
                                                    if getattr(it, 'optional_attendees', None):
                                                        s.update([a.mailbox.email_address for a in it.optional_attendees if a.mailbox])
                                                    return s
                                                att_diff = attendees_set(item) != attendees_set(master_item)
                                                if subj_diff or loc_diff or att_diff:
                                                    m_type = "Exception"
                                                else:
                                                    m_type = "Occurrence"
                                            except Exception:
                                                m_type = "Instance-Unknown"
                                        else:
                                            m_type = "Instance-Unknown"

                                # (3) Recurrence Pattern
                                m_recurrence = ""
                                m_pattern_details = ""
                                m_recurrence_duration = ""
                                m_is_endless = "N/A"
                                
                                if m_type == "RecurringMaster" and getattr(item, 'recurrence', None):
                                    pat = getattr(item.recurrence, 'pattern', None)
                                    if pat:
                                        m_recurrence = translate_pattern_type(pat.__class__.__name__)
                                        m_pattern_details = get_pattern_details(pat)
                                    m_recurrence_duration = get_recurrence_duration(item.recurrence)
                                    m_is_endless = is_endless_recurring(m_type, item.recurrence)
                                elif m_type in ("Occurrence", "Exception", "Instance", "Instance-Unknown"):
                                    # Instance: try find master via RecurringMasterId, then UID; cache it
                                    master_item_for_pattern = None
                                    try:
                                        # Check cache first
                                        if m_uid and m_uid in recurrence_cache:
                                            m_recurrence = recurrence_cache[m_uid]
                                            if self.log_level_var.get() == "Advanced":
                                                self.log(f"  使用缓存的 Pattern: {m_recurrence}", is_advanced=True)
                                        else:
                                            master_id = m_recurring_master_id
                                            master_ck = None
                                            if master_id is not None:
                                                if hasattr(master_id, "id"):
                                                    master_ck = getattr(master_id, "changekey", None)
                                                    master_id = master_id.id
                                                try:
                                                    master_item_for_pattern = account.calendar.get(id=master_id, changekey=master_ck) if master_ck else account.calendar.get(id=master_id)
                                                    if self.log_level_var.get() == "Advanced":
                                                        self.log(f"  通过 RecurringMasterId 获得主项", is_advanced=True)
                                                except Exception as e:
                                                    if self.log_level_var.get() == "Advanced":
                                                        self.log(f"  通过 RecurringMasterId 查找失败: {e}", is_advanced=True)
                                            
                                            if master_item_for_pattern is None and m_uid:
                                                try:
                                                    if self.log_level_var.get() == "Advanced":
                                                        self.log(f"  尝试通过 UID 查找主项...", is_advanced=True)
                                                    for m in account.calendar.all():
                                                        if getattr(m, 'uid', '') == m_uid and getattr(m, 'recurrence', None):
                                                            master_item_for_pattern = m
                                                            if self.log_level_var.get() == "Advanced":
                                                                self.log(f"  通过 UID 找到主项", is_advanced=True)
                                                            break
                                                except Exception as e:
                                                    if self.log_level_var.get() == "Advanced":
                                                        self.log(f"  通过 UID 查找主项失败: {e}", is_advanced=True)
                                            
                                            if master_item_for_pattern and getattr(master_item_for_pattern, 'recurrence', None):
                                                pat = getattr(master_item_for_pattern.recurrence, 'pattern', None)
                                                if pat:
                                                    m_recurrence = translate_pattern_type(pat.__class__.__name__)
                                                    m_pattern_details = get_pattern_details(pat)
                                                    if m_uid:
                                                        recurrence_cache[m_uid] = m_recurrence
                                                    if self.log_level_var.get() == "Advanced":
                                                        self.log(f"  获得 Pattern: {m_recurrence}", is_advanced=True)
                                                m_recurrence_duration = get_recurrence_duration(master_item_for_pattern.recurrence)
                                    except Exception as e:
                                        self.log(f"  获取主项用于计算 Pattern 失败: {e}", is_advanced=True)

                                row = {
                                    'UserPrincipalName': target_email,
                                    'Subject': subject,
                                    'Type': m_type,
                                    'MeetingGOID': m_goid,
                                    'CleanGOID': m_clean_goid,
                                    'Organizer': m_organizer,
                                    'Attendees': m_attendees_str,
                                    'Start': m_start,
                                    'End': m_end,
                                    'UserRole': m_role,
                                    'IsCancelled': m_is_cancelled,
                                    'ResponseStatus': m_response_status,
                                    'RecurrencePattern': m_recurrence,
                                    'PatternDetails': m_pattern_details,
                                    'RecurrenceDuration': m_recurrence_duration,
                                    'IsEndless': m_is_endless,
                                    'Action': 'Report' if self.report_only_var.get() else 'Delete',
                                    'Status': 'Pending',
                                    'Details': ''
                                }
                            else:
                                # Email Fields
                                sender_val = item.sender.email_address if item.sender else 'Unknown'
                                received_val = getattr(item, 'datetime_received', 'Unknown')
                                row = {
                                    'UserPrincipalName': target_email,
                                    'MessageId': item_id,
                                    'Subject': subject,
                                    'Sender': sender_val,
                                    'Received': received_val,
                                    'Action': 'Report' if self.report_only_var.get() else 'Delete',
                                    'Status': 'Pending',
                                    'Details': ''
                                }

                            if self.report_only_var.get():
                                self.log(f"  [报告] 发现: {item.subject}")
                                row['Status'] = 'Skipped'
                            else:
                                self.log(f"  正在删除: {item.subject}")
                                item.delete()
                                row['Status'] = 'Success'
                            
                            writer.writerow(row)
                            csvfile.flush()

                    except Exception as e:
                        self.log(f"  处理用户 {target_email} 出错: {e}", "ERROR")
                        self.log(f"  Traceback: {traceback.format_exc()}", is_advanced=True)
                        writer.writerow({'UserPrincipalName': target_email, 'Status': 'Error', 'Details': str(e)})

            self.log(f">>> 任务完成。报告: {report_path}")
            
            msg_title = "完成"
            if self.report_only_var.get():
                msg_body = "扫描生成报告任务完成。"
            else:
                msg_body = "清理任务已完成。"
                
            messagebox.showinfo(msg_title, msg_body)

        except Exception as e:
            self.log(f"EWS 运行时错误: {e}", "ERROR")
        finally:
            if ews_log_handlers:
                try:
                    for h in ews_log_handlers:
                        logging.getLogger("EWS_TRACE").removeHandler(h)
                        h.close()
                except:
                    pass
            # Reset Adapter
            BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
            EwsTraceAdapter.logger = None

if __name__ == "__main__":
    root = tk.Tk()
    app = UniversalEmailCleanerApp(root)
    root.mainloop()