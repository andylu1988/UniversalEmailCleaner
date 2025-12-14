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
import ctypes
from concurrent.futures import ThreadPoolExecutor
import calendar
import webbrowser
import base64
import io

APP_VERSION = "v1.5.4"
GITHUB_PROJECT_URL = "https://github.com/andylu1988/UniversalEmailCleaner"
GITHUB_PROFILE_URL = "https://github.com/andylu1988"

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

# DPI Awareness
try:
    # Try to set Per-Monitor DPI Awareness V2 (Windows 10 Creators Update and newer)
    ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
except Exception:
    try:
        # Fallback to Per-Monitor DPI Aware (Windows 8.1 and newer)
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            # Fallback to System DPI Aware (Windows Vista and newer)
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

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

class DateEntry(ttk.Frame):
    def __init__(self, master, textvariable, mode_var=None, other_date_var=None, **kwargs):
        super().__init__(master, **kwargs)
        self.variable = textvariable
        self.mode_var = mode_var
        self.other_date_var = other_date_var
        self.entry = ttk.Entry(self, textvariable=self.variable, width=15)
        self.entry.pack(side="left", fill="x", expand=True)
        # Bind click event to open calendar
        self.entry.bind("<Button-1>", lambda e: self.open_calendar())
        
        self.btn = ttk.Button(self, text="📅", width=3, command=self.open_calendar)
        self.btn.pack(side="left", padx=(2, 0))

    def open_calendar(self):
        # Check if calendar is already open
        if hasattr(self, 'top') and self.top.winfo_exists():
            self.top.lift()
            return

        self.top = tk.Toplevel(self)
        self.top.title("选择日期")
        self.top.geometry("280x280")
        self.top.grab_set()
        
        # Center popup
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        self.top.geometry(f"+{x}+{y}")

        cal_frame = ttk.Frame(self.top)
        cal_frame.pack(fill="both", expand=True, padx=5, pady=5)

        now = datetime.now()
        current_date = now
        try:
            if self.variable.get():
                current_date = datetime.strptime(self.variable.get(), "%Y-%m-%d")
        except:
            pass

        # For Meeting mode, use the other side date as reference for the +/- 2 year selectable window
        self._meeting_ref_date = None
        if self.mode_var and self.mode_var.get() == "Meeting":
            ref_raw = None
            try:
                if self.other_date_var and self.other_date_var.get():
                    ref_raw = self.other_date_var.get()
            except Exception:
                ref_raw = None

            if ref_raw:
                try:
                    self._meeting_ref_date = datetime.strptime(ref_raw, "%Y-%m-%d").date()
                except Exception:
                    self._meeting_ref_date = None

            if self._meeting_ref_date is None:
                try:
                    if self.variable.get():
                        self._meeting_ref_date = datetime.strptime(self.variable.get(), "%Y-%m-%d").date()
                except Exception:
                    self._meeting_ref_date = None

            if self._meeting_ref_date is None:
                self._meeting_ref_date = datetime.now().date()
        
        self.cal_year = tk.IntVar(value=current_date.year)
        self.cal_month = tk.IntVar(value=current_date.month)
        
        # Header with Year/Month navigation + quick selectors
        header = ttk.Frame(cal_frame)
        header.pack(fill="x", pady=5)
        
        # Nav
        ttk.Button(header, text="<<", width=3, command=lambda: self.change_year(-1, cal_grid)).pack(side="left")
        ttk.Button(header, text="<", width=2, command=lambda: self.change_month(-1, cal_grid)).pack(side="left")

        # Quick selectors
        year_values = None
        if self.mode_var and self.mode_var.get() == "Meeting" and getattr(self, "_meeting_ref_date", None):
            y0 = self._meeting_ref_date.year
            year_values = [str(y) for y in range(y0 - 2, y0 + 3)]
        else:
            y0 = current_date.year
            year_values = [str(y) for y in range(y0 - 10, y0 + 11)]

        self.year_cb = ttk.Combobox(header, values=year_values, width=6, state="readonly")
        self.year_cb.set(str(self.cal_year.get()))
        self.year_cb.pack(side="left", padx=(6, 2))

        self.month_cb = ttk.Combobox(header, values=[f"{m:02d}" for m in range(1, 13)], width=4, state="readonly")
        self.month_cb.set(f"{self.cal_month.get():02d}")
        self.month_cb.pack(side="left", padx=(2, 6))

        def _on_year_month_change(_evt=None):
            try:
                y = int(self.year_cb.get())
                m = int(self.month_cb.get())
            except Exception:
                return
            self.cal_year.set(y)
            self.cal_month.set(m)
            self.render_calendar(cal_grid, self.top)

        self.year_cb.bind("<<ComboboxSelected>>", _on_year_month_change)
        self.month_cb.bind("<<ComboboxSelected>>", _on_year_month_change)

        ttk.Button(header, text=">", width=2, command=lambda: self.change_month(1, cal_grid)).pack(side="left")
        ttk.Button(header, text=">>", width=3, command=lambda: self.change_year(1, cal_grid)).pack(side="left")

        # Grid
        cal_grid = ttk.Frame(cal_frame)
        cal_grid.pack(fill="both", expand=True)
        
        self.render_calendar(cal_grid, self.top)

    def change_year(self, delta, grid_frame):
        y = self.cal_year.get() + delta
        self.cal_year.set(y)
        if hasattr(self, "year_cb"):
            self.year_cb.set(str(y))
        self.render_calendar(grid_frame, grid_frame.winfo_toplevel())

    def change_month(self, delta, grid_frame):
        m = self.cal_month.get() + delta
        y = self.cal_year.get()
        if m < 1:
            m = 12
            y -= 1
        elif m > 12:
            m = 1
            y += 1
        self.cal_year.set(y)
        self.cal_month.set(m)
        if hasattr(self, "year_cb"):
            self.year_cb.set(str(y))
        if hasattr(self, "month_cb"):
            self.month_cb.set(f"{m:02d}")
        self.render_calendar(grid_frame, grid_frame.winfo_toplevel())

    def render_calendar(self, frame, top):
        for widget in frame.winfo_children():
            widget.destroy()
            
        days = ["一", "二", "三", "四", "五", "六", "日"]
        for i, d in enumerate(days):
            ttk.Label(frame, text=d, anchor="center").grid(row=0, column=i, sticky="nsew")
            
        cal = calendar.monthcalendar(self.cal_year.get(), self.cal_month.get())
        
        # Determine valid date range if in Meeting mode
        min_date = None
        max_date = None
        if self.mode_var and self.mode_var.get() == "Meeting":
            ref_d = getattr(self, "_meeting_ref_date", None) or datetime.now().date()
            min_date = ref_d - timedelta(days=365*2)
            max_date = ref_d + timedelta(days=365*2)

        for r, week in enumerate(cal):
            for c, day in enumerate(week):
                if day != 0:
                    state = "normal"
                    if min_date and max_date:
                        try:
                            current_d = datetime(self.cal_year.get(), self.cal_month.get(), day).date()
                            if not (min_date <= current_d <= max_date):
                                state = "disabled"
                        except ValueError:
                            pass

                    btn = tk.Button(frame, text=str(day), relief="flat", state=state,
                                    command=lambda d=day: self.select_date(d, top))
                    btn.grid(row=r+1, column=c, sticky="nsew", padx=1, pady=1)
        
        for i in range(7):
            frame.columnconfigure(i, weight=1)

    def select_date(self, day, top):
        date_str = f"{self.cal_year.get()}-{self.cal_month.get():02d}-{day:02d}"
        self.variable.set(date_str)
        top.destroy()


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
    def __init__(self, log_area, log_dir):
        self.log_area = log_area
        self.log_dir = log_dir
        self.level = "NORMAL" # NORMAL or ADVANCED
        self.file_lock = threading.Lock()

    def _get_log_file_path(self):
        date_str = datetime.now().strftime("%Y-%m-%d")
        return os.path.join(self.log_dir, f"app_{date_str}.log")

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
            if self.log_area:
                self.log_area.config(state='normal')
                self.log_area.insert(tk.END, full_msg + "\n")
                self.log_area.see(tk.END)
                self.log_area.config(state='disabled')
        
        if self.log_area:
            self.log_area.after(0, _update)
        
        # File Write
        try:
            log_path = self._get_log_file_path()
            with self.file_lock:
                with open(log_path, "a", encoding="utf-8") as f:
                    f.write(full_msg + "\n")
        except:
            pass

    def log_to_file_only(self, message):
        """Writes directly to file, skipping GUI. Useful for large debug dumps."""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        full_msg = f"[{timestamp}] [DEBUG_DATA] {message}"
        try:
            log_path = self._get_log_file_path()
            with self.file_lock:
                with open(log_path, "a", encoding="utf-8") as f:
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
        self.root.title(f"通用邮件清理工具 {APP_VERSION} (Graph API & EWS)")
        self.root.geometry("1100x900")
        self.root.minsize(900, 700)
        
        # Auto-scale font based on DPI if possible, or just use a slightly larger default
        default_font_size = 10
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('.', font=('Segoe UI', default_font_size))
        style.configure('Treeview', font=('Segoe UI', default_font_size))
        style.configure('TButton', font=('Segoe UI', default_font_size))
        
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
        
        # EWS Autodiscover Menu
        tools_menu.add_command(label="使用自动发现刷新 EWS 配置 (Refresh EWS Config)", command=self.refresh_ews_config)
        tools_menu.add_separator()

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
        
        # Log Toolbar
        log_toolbar = ttk.Frame(log_frame)
        log_toolbar.pack(fill="x", padx=5, pady=2)
        
        self.log_visible = True
        def toggle_log():
            current_height = self.root.winfo_height()
            if self.log_visible:
                self.log_area.pack_forget()
                self.btn_toggle_log.config(text="显示日志 (Show Log)")
                
                # Shrink window to avoid empty space
                # Assuming log area is roughly 200px
                new_height = max(600, current_height - 200)
                self.root.geometry(f"{self.root.winfo_width()}x{new_height}")
                
                # Stop log frame from expanding
                log_frame.pack_configure(expand=False)
                
                self.log_visible = False
            else:
                # Pack before link_frame to ensure it stays above links
                self.log_area.pack(fill="both", expand=True, padx=5, pady=5, before=self.link_frame)
                self.btn_toggle_log.config(text="隐藏日志 (Hide Log)")
                
                # Restore window height
                new_height = current_height + 200
                self.root.geometry(f"{self.root.winfo_width()}x{new_height}")
                
                # Allow log frame to expand
                log_frame.pack_configure(expand=True)
                
                self.log_visible = True
                
        self.btn_toggle_log = ttk.Button(log_toolbar, text="隐藏日志 (Hide Log)", command=toggle_log, width=20)
        self.btn_toggle_log.pack(side="right")

        self.log_area = scrolledtext.ScrolledText(log_frame, height=12, state='disabled', font=("Consolas", 10))
        self.log_area.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.logger = Logger(self.log_area, self.documents_dir)

        # Links
        self.link_frame = ttk.Frame(log_frame)
        self.link_frame.pack(fill="x", padx=5)
        
        log_dir = os.path.dirname(self.log_file_path)
        self.log_link_lbl = tk.Label(self.link_frame, text=f"日志目录: {log_dir}", fg="blue", cursor="hand2")
        self.log_link_lbl.pack(side="left")
        self.log_link_lbl.bind("<Button-1>", lambda e: os.startfile(log_dir) if os.path.exists(log_dir) else None)
        
        self.report_link_lbl = tk.Label(self.link_frame, text="", fg="blue", cursor="hand2")
        self.report_link_lbl.pack(side="left", padx=20)

        # Build Tabs
        self.build_connection_tab()
        self.build_cleanup_tab()

        self.load_config()
        
        # Ensure UI state matches config
        self.toggle_connection_ui()

    def refresh_ews_config(self):
        if not self.ews_user_var.get() or not self.ews_pass_var.get():
            messagebox.showwarning("提示", "请先在连接配置中填写 EWS 管理员账号和密码。")
            return
        
        self.ews_use_autodiscover.set(True)
        self.test_ews_connection()

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
        about = tk.Toplevel(self.root)
        about.title("关于")
        about.resizable(False, False)
        about.geometry("520x260")
        about.transient(self.root)
        about.grab_set()

        outer = ttk.Frame(about, padding=12)
        outer.pack(fill="both", expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x")

        avatar_lbl = ttk.Label(top)
        avatar_lbl.pack(side="left", padx=(0, 12))

        text_col = ttk.Frame(top)
        text_col.pack(side="left", fill="both", expand=True)

        ttk.Label(text_col, text=f"通用邮件清理工具 (Universal Email Cleaner) {APP_VERSION}", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(text_col, text="支持 Microsoft Graph API 和 Exchange Web Services (EWS)。\n用于批量清理或生成邮件报告。", justify="left").pack(anchor="w", pady=(6, 8))

        link = tk.Label(text_col, text=GITHUB_PROJECT_URL, fg="#1a73e8", cursor="hand2")
        try:
            link.configure(font=("Segoe UI", 10, "underline"))
        except Exception:
            pass
        link.pack(anchor="w")
        link.bind("<Button-1>", lambda _e: webbrowser.open(GITHUB_PROJECT_URL))

        ttk.Label(text_col, text=f"GitHub: {GITHUB_PROFILE_URL}").pack(anchor="w", pady=(6, 0))

        # Load avatar from avatar_b64.txt (preferred) and show it if Pillow is available
        def _try_load_avatar_b64():
            base_dir = os.path.dirname(os.path.abspath(__file__))
            candidates = [
                os.path.join(base_dir, "avatar_b64.txt"),
            ]
            if getattr(sys, 'frozen', False):
                candidates.append(os.path.join(sys._MEIPASS, "avatar_b64.txt"))

            for p in candidates:
                try:
                    if os.path.exists(p):
                        with open(p, "r", encoding="utf-8") as f:
                            return f.read().strip()
                except Exception:
                    continue
            return None

        avatar_b64 = _try_load_avatar_b64()
        if avatar_b64 and Image is not None and ImageTk is not None:
            try:
                raw = base64.b64decode(avatar_b64)
                img = Image.open(io.BytesIO(raw)).convert("RGBA")
                img = img.resize((88, 88))
                about._avatar_img = ImageTk.PhotoImage(img)
                avatar_lbl.configure(image=about._avatar_img)
            except Exception:
                pass

        btns = ttk.Frame(outer)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="关闭", command=about.destroy).pack(side="right")

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
                    self.source_type_var.set(config.get('source_type', 'EWS')) # Default to EWS if not set
                    self.csv_path_var.set(config.get('csv_path', ''))
                    self.log(">>> 配置已加载。")
            except Exception as e:
                self.log(f"X 加载配置失败: {e}", "ERROR")
        else:
            # No config file, set default to EWS
            self.source_type_var.set("EWS")
            self.toggle_connection_ui()

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
        self.ews_frame = ttk.LabelFrame(main_frame, text="EWS 配置 (Exchange On-Premise)")
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
        graph_common_frame.pack(fill="x", padx=10, pady=5)
        
        # Environment Selection
        env_frame = ttk.Frame(graph_common_frame)
        env_frame.pack(side="left")
        ttk.Label(env_frame, text="环境:").pack(side="left")
        ttk.Radiobutton(env_frame, text="全球版 (Global)", variable=self.graph_env_var, value="Global").pack(side="left", padx=5)
        ttk.Radiobutton(env_frame, text="世纪互联 (China)", variable=self.graph_env_var, value="China").pack(side="left", padx=5)
        
        ttk.Separator(graph_common_frame, orient="vertical").pack(side="left", fill="y", padx=10)
        
        # Mode Selection
        mode_frame = ttk.Frame(graph_common_frame)
        mode_frame.pack(side="left")
        ttk.Label(mode_frame, text="配置方式:").pack(side="left")
        ttk.Radiobutton(mode_frame, text="自动配置 (证书)", variable=self.graph_auth_mode_var, value="Auto", command=self.toggle_graph_ui).pack(side="left", padx=5)
        ttk.Radiobutton(mode_frame, text="手动配置 (Secret)", variable=self.graph_auth_mode_var, value="Manual", command=self.toggle_graph_ui).pack(side="left", padx=5)

        # Graph Auto Frame
        self.graph_auto_frame = ttk.Frame(self.graph_frame)
        # Note: We don't pack it here, toggle_graph_ui will handle it
        
        ttk.Button(self.graph_auto_frame, text="一键初始化 (创建 App & 证书)", command=self.start_graph_setup_thread).pack(side="left", padx=0)
        ttk.Button(self.graph_auto_frame, text="删除 App", command=self.start_delete_app_thread).pack(side="left", padx=5)

        # Graph Manual Frame
        self.graph_manual_frame = ttk.Frame(self.graph_frame)
        # Note: We don't pack it here, toggle_graph_ui will handle it
        
        manual_grid = ttk.Frame(self.graph_manual_frame)
        manual_grid.pack(anchor="w", fill="x", expand=True)
        
        grid_opts = {'padx': 5, 'pady': 5, 'sticky': 'w'}
        
        ttk.Label(manual_grid, text="租户 ID (Tenant ID):").grid(row=0, column=0, **grid_opts)
        ttk.Entry(manual_grid, textvariable=self.tenant_id_var, width=50).grid(row=0, column=1, **grid_opts)
        
        ttk.Label(manual_grid, text="客户端 ID (App ID):").grid(row=1, column=0, **grid_opts)
        ttk.Entry(manual_grid, textvariable=self.app_id_var, width=50).grid(row=1, column=1, **grid_opts)
        
        ttk.Label(manual_grid, text="客户端密钥 (Client Secret):").grid(row=2, column=0, **grid_opts)
        ttk.Entry(manual_grid, textvariable=self.client_secret_var, width=50, show="*").grid(row=2, column=1, **grid_opts)

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
            self.graph_manual_frame.pack_forget()
            self.graph_auto_frame.pack(fill="x", padx=10, pady=5)
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
                # Update server var if successful
                if account.protocol.service_endpoint:
                    self.ews_server_var.set(account.protocol.service_endpoint)
                    self.log(f"Autodiscover found server: {account.protocol.service_endpoint}")
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
        self.start_date_entry = DateEntry(self.filter_frame, textvariable=self.criteria_start_date, mode_var=self.cleanup_target_var, other_date_var=self.criteria_end_date)
        self.start_date_entry.grid(row=2, column=1, **grid_opts)
        
        ttk.Label(self.filter_frame, text="结束日期 (YYYY-MM-DD):").grid(row=2, column=2, **grid_opts)
        self.end_date_entry = DateEntry(self.filter_frame, textvariable=self.criteria_end_date, mode_var=self.cleanup_target_var, other_date_var=self.criteria_start_date)
        self.end_date_entry.grid(row=2, column=3, **grid_opts)

        self.meeting_date_hint_label = ttk.Label(
            self.filter_frame,
            text="提示：Graph 会议必须填写开始/结束日期（用于展开循环会议 occurrence/exception）；EWS 不填写日期则不展开循环会议实例。"
        )
        self.meeting_date_hint_label.grid(row=3, column=0, columnspan=4, padx=5, pady=(2, 0), sticky='w')

        self.lbl_body = ttk.Label(self.filter_frame, text="正文包含:")
        self.lbl_body.grid(row=4, column=0, **grid_opts)
        ttk.Entry(self.filter_frame, textvariable=self.criteria_body, width=80).grid(row=4, column=1, columnspan=3, **grid_opts)

        self.update_ui_for_target() # Init state

        # Options
        opt_frame = ttk.LabelFrame(frame, text="执行选项")
        opt_frame.pack(fill="x", pady=5)
        
        self.btn_start_text = tk.StringVar(value="开始扫描 (Start Scan)")
        
        def on_report_only_change():
            if self.report_only_var.get():
                self.btn_start_text.set("开始扫描 (Start Scan)")
            else:
                self.btn_start_text.set("开始清理 (Start Clean)")
                messagebox.showwarning("警告", "您已取消 '仅报告' 模式！\n\n接下来的操作将 **永久删除** 数据！\n请务必确认 CSV 和 筛选条件 正确！")

        ttk.Checkbutton(opt_frame, text="仅报告 (不删除)", variable=self.report_only_var, command=on_report_only_change).pack(side="left", padx=10)
        
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
        ttk.Button(frame, textvariable=self.btn_start_text, command=self.start_cleanup_thread).pack(pady=10, ipadx=20, ipady=5)

    def update_ui_for_target(self):
        target = self.cleanup_target_var.get()
        if target == "Meeting":
            self.meeting_opt_frame.pack(fill="x", pady=5, after=self.filter_frame) # Pack below filter or above? Let's put it above filter
            self.meeting_opt_frame.pack(fill="x", pady=5, before=self.filter_frame)
            
            self.lbl_subject.config(text="会议标题包含:")
            self.lbl_sender.config(text="组织者地址:")
            self.lbl_body.config(text="会议内容包含:")
            if hasattr(self, 'meeting_date_hint_label'):
                self.meeting_date_hint_label.grid()
        else:
            self.meeting_opt_frame.pack_forget()
            self.lbl_subject.config(text="邮件主题包含:")
            self.lbl_sender.config(text="发件人地址:")
            self.lbl_body.config(text="邮件正文包含:")
            if hasattr(self, 'meeting_date_hint_label'):
                self.meeting_date_hint_label.grid_remove()

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

        # Date Range Validation for Meetings
        if self.cleanup_target_var.get() == "Meeting":
            start_str = self._normalize_date_input(self.criteria_start_date.get())
            end_str = self._normalize_date_input(self.criteria_end_date.get())

            # Graph calendarView requires both start and end
            if self.source_type_var.get() == "Graph":
                if not start_str or not end_str:
                    messagebox.showwarning("配置缺失", "Graph 模式下会议查询需要【开始日期】和【结束日期】（用于展开循环会议 occurrence/exception）。")
                    return

            if start_str and end_str:
                try:
                    s_dt = datetime.strptime(start_str, "%Y-%m-%d")
                    e_dt = datetime.strptime(end_str, "%Y-%m-%d")
                    if (e_dt - s_dt).days > 730: # Approx 2 years
                        messagebox.showwarning("日期范围过大", "会议清理的时间跨度不能超过 2 年。")
                        return
                except Exception:
                    pass

        # Double Confirmation for Deletion
        if not self.report_only_var.get():
            confirm1 = messagebox.askyesno("高风险操作确认", "您当前处于【删除模式】！\n\n程序将 **永久删除** 匹配的邮件/会议，且 **无法恢复**。\n\n是否确认继续？")
            if not confirm1:
                return
            
            confirm2 = messagebox.askyesno("最终确认", "请再次确认：\n\n1. 您已备份重要数据。\n2. 您已确认 CSV 用户列表无误。\n3. 您已确认筛选条件无误。\n\n点击 '是' 将立即开始删除操作！")
            if not confirm2:
                return

        self.logger.set_level(self.log_level_var.get().upper())
        self.save_config()
        
        threading.Thread(target=self.run_cleanup, daemon=True).start()

    def run_cleanup(self):
        self.log("-" * 60)
        self.log(f"任务开始: {datetime.now()}")
        self.log(f"模式: {'仅报告 (Report Only)' if self.report_only_var.get() else '删除 (DELETE)'}")
        self.log("-" * 60)

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

    def process_single_user_graph(self, user, graph_endpoint, headers, resource, delete_resource, target_type, filter_str, body_keyword,
                                  report_only, writer, csv_lock, calendar_view_start=None, calendar_view_end=None):
        self.log(f"--- 正在处理: {user} ---")
        try:
            # Meetings: prefer calendarView to expand recurrence into occurrence/exception within a date range
            if target_type == "Meeting" and resource == "calendarView":
                if not calendar_view_start or not calendar_view_end:
                    raise Exception("Graph calendarView requires startDateTime and endDateTime")
                url = f"{graph_endpoint}/v1.0/users/{user}/calendarView"
                params = {
                    "startDateTime": calendar_view_start,
                    "endDateTime": calendar_view_end,
                    "$top": 100,
                    "$select": "id,subject,organizer,attendees,start,end,type,isCancelled,iCalUId,seriesMasterId,bodyPreview",
                    # Try to read GOID via MAPI extended property PidLidGlobalObjectId (PSETID_Meeting, Id 0x0003)
                    "$expand": "singleValueExtendedProperties($filter=id eq 'Binary {6ED8DA90-450B-101B-98DA-00AA003F1305} Id 0x0003')",
                }
            else:
                url = f"{graph_endpoint}/v1.0/users/{user}/{resource}"
                if target_type == "Email":
                    params = {"$top": 100, "$select": "id,subject,from,receivedDateTime,createdDateTime,body"}
                else:
                    params = {"$top": 100, "$select": "id,subject,organizer,attendees,start,end,type,isCancelled,iCalUId,seriesMasterId,bodyPreview"}

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
                    with csv_lock:
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
                            start_val = item.get('start', {}).get('dateTime')
                            end_val = item.get('end', {}).get('dateTime')
                            item_type = item.get('type', 'Event')

                            attendees = item.get('attendees', []) or []
                            attendee_emails = []
                            for a in attendees:
                                addr = (a.get('emailAddress') or {}).get('address')
                                if addr:
                                    attendee_emails.append(addr)

                            response_status = item.get('responseStatus', {}).get('response', '')
                            is_cancelled = bool(item.get('isCancelled'))
                            ical_uid = item.get('iCalUId', '')
                            series_master_id = item.get('seriesMasterId', '')

                            goid_b64 = ''
                            try:
                                props = item.get('singleValueExtendedProperties') or []
                                if props:
                                    # Graph returns base64 for Binary extended properties
                                    goid_b64 = props[0].get('value', '') or ''
                            except Exception:
                                goid_b64 = ''

                            row_data = {
                                'UserPrincipalName': user,
                                'Subject': subject,
                                'Type': item_type,
                                'MeetingGOID': goid_b64,
                                'CleanGOID': (goid_b64 or ical_uid or item_id),
                                'iCalUId': ical_uid,
                                'SeriesMasterId': series_master_id,
                                'Organizer': sender,
                                'Attendees': ';'.join(attendee_emails),
                                'Start': start_val,
                                'End': end_val,
                                'UserRole': '',
                                'IsCancelled': is_cancelled,
                                'ResponseStatus': response_status,
                                'RecurrencePattern': '',
                                'PatternDetails': '',
                                'RecurrenceDuration': '',
                                'IsEndless': '',
                                'Action': 'ReportOnly' if report_only else 'Delete',
                                'Status': 'Pending',
                                'Details': ''
                            }

                            if is_cancelled:
                                row_data['Type'] = f"{row_data['Type']} (Cancelled)"

                            # Best-effort: if this is an occurrence/exception and has seriesMasterId,
                            # pull recurrence from master (cache per user) to align with EWS report.
                            try:
                                if (item_type in ('occurrence', 'exception') or 'occurrence' in str(item_type).lower() or 'exception' in str(item_type).lower()) and series_master_id:
                                    if not hasattr(self, '_graph_master_cache'):
                                        self._graph_master_cache = {}
                                    user_cache = self._graph_master_cache.setdefault(user, {})
                                    if series_master_id not in user_cache:
                                        master_url = f"{graph_endpoint}/v1.0/users/{user}/events/{series_master_id}"
                                        master_params = {
                                            "$select": "id,type,recurrence,iCalUId",
                                        }
                                        m_resp = requests.get(master_url, headers=headers, params=master_params)
                                        if m_resp.status_code == 200:
                                            user_cache[series_master_id] = m_resp.json()
                                        else:
                                            user_cache[series_master_id] = None
                                    master_obj = user_cache.get(series_master_id)
                                    if master_obj and master_obj.get('recurrence'):
                                        rec = master_obj.get('recurrence')
                                        pattern = (rec.get('pattern') or {})
                                        rng = (rec.get('range') or {})
                                        row_data['RecurrencePattern'] = pattern.get('type', '')
                                        row_data['PatternDetails'] = json.dumps(pattern, ensure_ascii=False)
                                        row_data['RecurrenceDuration'] = json.dumps(rng, ensure_ascii=False)
                                        row_data['IsEndless'] = (rng.get('type') == 'noEnd')
                            except Exception:
                                pass

                        if target_type == "Email":
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
                            del_url = f"{graph_endpoint}/v1.0/users/{user}/{delete_resource}/{item_id}"
                            
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
                        
                        with csv_lock:
                            writer.writerow(row_data)
                            # csvfile.flush() # Flush handled by main loop or context manager

                url = data.get('@odata.nextLink')
                # Reset params for next link as they are usually included
                params = None 
                
        except Exception as ue:
            self.log(f"  X 处理用户出错: {ue}", "ERROR")
            with csv_lock:
                writer.writerow({'UserPrincipalName': user, 'Status': 'Error', 'Details': str(ue)})

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
                if target_type == "Meeting":
                    fieldnames = [
                        'UserPrincipalName', 'Subject', 'Type', 'MeetingGOID', 'CleanGOID',
                        'iCalUId', 'SeriesMasterId',
                        'Organizer', 'Attendees', 'Start', 'End', 'UserRole',
                        'IsCancelled', 'ResponseStatus', 'RecurrencePattern', 'PatternDetails', 'RecurrenceDuration', 'IsEndless',
                        'Action', 'Status', 'Details'
                    ]
                else:
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
                    delete_resource = "messages"
                    if self.criteria_msg_id.get(): filters.append(f"internetMessageId eq '{self.criteria_msg_id.get()}'")
                    if self.criteria_sender.get(): filters.append(f"from/emailAddress/address eq '{self.criteria_sender.get()}'")
                    if start_date: filters.append(f"receivedDateTime ge {start_date}T00:00:00Z")
                    if end_date: filters.append(f"receivedDateTime le {end_date}T23:59:59Z")
                else: # Meeting
                    resource = "calendarView"
                    delete_resource = "events"
                    if self.criteria_sender.get(): filters.append(f"organizer/emailAddress/address eq '{self.criteria_sender.get()}'")
                    # calendarView uses startDateTime/endDateTime query params for time window; keep filter for additional criteria only
                    
                    # Meeting Specifics
                    if self.meeting_only_cancelled_var.get():
                        filters.append("isCancelled eq true")
                    
                    scope = self.meeting_scope_var.get()
                    if "Single" in scope:
                        filters.append("type eq 'singleInstance'")
                    elif "Series" in scope:
                        # Include expanded instances too (occurrence/exception) for EWS-like scan
                        filters.append("type eq 'seriesMaster' or type eq 'occurrence' or type eq 'exception'")
                    # If All, no type filter

                filter_str = " and ".join(filters)
                body_keyword = self.criteria_body.get()

                calendar_view_start = None
                calendar_view_end = None
                if target_type == "Meeting":
                    # Use UTC ISO; calendarView requires both
                    calendar_view_start = f"{start_date}T00:00:00Z" if start_date else None
                    calendar_view_end = f"{end_date}T23:59:59Z" if end_date else None

                csv_lock = threading.Lock()
                report_only = self.report_only_var.get()
                
                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = []
                    for user in users:
                        futures.append(executor.submit(
                            self.process_single_user_graph, 
                            user, graph_endpoint, headers, resource, delete_resource, target_type, filter_str, body_keyword,
                            report_only, writer, csv_lock, calendar_view_start, calendar_view_end
                        ))
                    
                    # Wait for all to complete
                    for future in futures:
                        try:
                            future.result()
                        except Exception as e:
                            self.log(f"Task Error: {e}", "ERROR")

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

    def process_single_user_ews(self, target_email, creds, config, auth_type, use_auto, target_type, 
                                start_date_str, end_date_str, criteria_sender, criteria_msg_id, 
                                criteria_subject, criteria_body, meeting_only_cancelled, meeting_scope, 
                                report_only, writer, csv_lock, log_level):
        try:
            self.log(f"--- 正在处理: {target_email} ---")
            
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

            # Date Parsing
            start_dt = None
            end_dt = None
            
            if start_date_str:
                dt = datetime.strptime(start_date_str, "%Y-%m-%d")
                start_dt = EWSDateTime.from_datetime(dt).replace(tzinfo=account.default_timezone)
                
            if end_date_str:
                dt = datetime.strptime(end_date_str, "%Y-%m-%d") + timedelta(days=1)
                end_dt = EWSDateTime.from_datetime(dt).replace(tzinfo=account.default_timezone)

            # Recurrence Cache
            recurrence_cache = {}

            # Optimization: Set page size for faster retrieval
            page_size = 100

            if target_type == "Email":
                qs = account.inbox.all().order_by('-datetime_received')
                qs.page_size = page_size
                if start_dt:
                    qs = qs.filter(datetime_received__gte=start_dt)
                if end_dt:
                    qs = qs.filter(datetime_received__lt=end_dt)
                if criteria_sender:
                    qs = qs.filter(sender__icontains=criteria_sender)
            else:
                # Meeting Logic with CalendarView
                if start_dt or end_dt:
                    # View requires both start and end. Default if missing.
                    view_start = start_dt if start_dt else EWSDateTime(1900, 1, 1, tzinfo=account.default_timezone)
                    view_end = end_dt if end_dt else EWSDateTime(2100, 1, 1, tzinfo=account.default_timezone)
                    
                    self.log(f"使用日历视图 (CalendarView) 展开循环会议: {view_start} -> {view_end}", is_advanced=True)
                    qs = account.calendar.view(start=view_start, end=view_end)
                    # CalendarView does not support page_size in the same way as QuerySet, but we can try
                else:
                    self.log("未指定日期范围，使用普通查询 (不展开循环会议实例)", is_advanced=True)
                    qs = account.calendar.all()
                    qs.page_size = page_size

                if criteria_sender:
                    qs = qs.filter(organizer__icontains=criteria_sender)
                
                if meeting_only_cancelled:
                    qs = qs.filter(is_cancelled=True)

            # Common Filters
            is_calendar_view = (target_type == "Meeting" and (start_dt or end_dt))

            if not is_calendar_view:
                if criteria_msg_id:
                    qs = qs.filter(message_id=criteria_msg_id)
                if criteria_subject:
                    qs = qs.filter(subject__icontains=criteria_subject)
            
            self.log(f"正在查询 EWS...", is_advanced=True)
            
            # Optimization: Use only() to fetch required fields if possible, but CalendarView is tricky.
            # For Email, we can optimize.
            if target_type == "Email":
                # We need: id, subject, sender, datetime_received, body (if filtered)
                # Note: 'body' can be heavy. Only fetch if needed.
                fields = ['id', 'changekey', 'subject', 'sender', 'datetime_received']
                if criteria_body:
                    fields.append('body')
                qs = qs.only(*fields)

            items = list(qs) # Execute query
            if not items:
                self.log(f"用户 {target_email} 未找到项目。")
            
            for item in items:
                # Client Side Filters
                if target_type == "Meeting":
                    # 1. Filter by Subject (if CalendarView)
                    if is_calendar_view and criteria_subject:
                        if criteria_subject.lower() not in (item.subject or "").lower():
                            continue
                    
                    # 2. Filter by Organizer (if CalendarView)
                    if is_calendar_view and criteria_sender:
                        organizer_email = item.organizer.email_address if item.organizer else ""
                        if criteria_sender.lower() not in organizer_email.lower():
                            continue

                    # 3. Filter by IsCancelled (if CalendarView)
                    if is_calendar_view and meeting_only_cancelled:
                        if not item.is_cancelled:
                            continue

                    inferred_type = guess_calendar_item_type(item)

                    # Apply scope filter using inferred type
                    if "Single" in meeting_scope:
                        if inferred_type != 'Single': continue
                    elif "Series" in meeting_scope:
                        if inferred_type not in ('RecurringMaster', 'Occurrence', 'Exception'): continue

                # Body check
                if criteria_body:
                    if criteria_body.lower() not in (item.body or "").lower():
                        continue

                # Enrich meeting item details
                if target_type == "Meeting":
                    try:
                        _id = getattr(item, 'id', None) or getattr(item, 'item_id', None)
                        _ck = getattr(item, 'changekey', None) or getattr(item, 'change_key', None) or getattr(item, 'changeKey', None)
                        if _id:
                            full_item = account.calendar.get(id=_id, changekey=_ck) if _ck else account.calendar.get(id=_id)
                            try:
                                full_item.refresh()
                            except Exception:
                                pass
                            if full_item:
                                item = full_item
                    except Exception as _e:
                        self.log(f"  无法获取完整项属性 (GetItem): {_e}", is_advanced=True)

                # Extract attributes
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
                        m_type = guess_calendar_item_type(item)

                    m_uid = getattr(item, 'uid', '')
                    m_recurring_master_id = getattr(item, 'recurring_master_id', None)
                    m_goid = m_uid
                    m_clean_goid = m_goid 
                    m_organizer = item.organizer.email_address if item.organizer else 'Unknown'
                    
                    m_attendees = []
                    if item.required_attendees:
                        m_attendees.extend([a.mailbox.email_address for a in item.required_attendees if a.mailbox])
                    if item.optional_attendees:
                        m_attendees.extend([a.mailbox.email_address for a in item.optional_attendees if a.mailbox])
                    m_attendees_str = "; ".join(m_attendees)
                    
                    m_start = getattr(item, 'start', '')
                    m_end = getattr(item, 'end', '')
                    
                    m_role = 'Attendee'
                    if m_organizer.lower() == target_email.lower():
                        m_role = 'Organizer'
                    
                    m_is_cancelled = getattr(item, 'is_cancelled', False)
                    m_response_status = getattr(item, 'my_response_type', 'Unknown')
                    
                    # (2) If instance, determine Occurrence vs Exception
                    original_start = getattr(item, 'original_start', None)
                    if m_type == "Instance":
                        start_val = getattr(item, 'start', None)
                        if start_val and original_start and start_val != original_start:
                            m_type = "Exception"
                        else:
                            master_item = None
                            try:
                                master_id = m_recurring_master_id
                                master_ck = None
                                if master_id is not None:
                                    if hasattr(master_id, "id"):
                                        master_ck = getattr(master_id, "changekey", None)
                                        master_id = master_id.id
                                    master_item = account.calendar.get(id=master_id, changekey=master_ck) if master_ck else account.calendar.get(id=master_id)
                                if master_item is None and m_uid:
                                    for m in account.calendar.all():
                                        if getattr(m, 'uid', '') == m_uid and getattr(m, 'recurrence', None):
                                            master_item = m
                                            break
                            except Exception:
                                pass

                            if master_item:
                                try:
                                    subj_diff = (item.subject or '') != (master_item.subject or '')
                                    loc_diff = (getattr(item, 'location', None) or '') != (getattr(master_item, 'location', None) or '')
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
                        master_item_for_pattern = None
                        try:
                            if m_uid and m_uid in recurrence_cache:
                                m_recurrence = recurrence_cache[m_uid]
                            else:
                                master_id = m_recurring_master_id
                                master_ck = None
                                if master_id is not None:
                                    if hasattr(master_id, "id"):
                                        master_ck = getattr(master_id, "changekey", None)
                                        master_id = master_id.id
                                    try:
                                        master_item_for_pattern = account.calendar.get(id=master_id, changekey=master_ck) if master_ck else account.calendar.get(id=master_id)
                                    except Exception:
                                        pass
                                
                                if master_item_for_pattern is None and m_uid:
                                    try:
                                        for m in account.calendar.all():
                                            if getattr(m, 'uid', '') == m_uid and getattr(m, 'recurrence', None):
                                                master_item_for_pattern = m
                                                break
                                    except Exception:
                                        pass
                                
                                if master_item_for_pattern and getattr(master_item_for_pattern, 'recurrence', None):
                                    pat = getattr(master_item_for_pattern.recurrence, 'pattern', None)
                                    if pat:
                                        m_recurrence = translate_pattern_type(pat.__class__.__name__)
                                        m_pattern_details = get_pattern_details(pat)
                                        if m_uid:
                                            recurrence_cache[m_uid] = m_recurrence
                                    m_recurrence_duration = get_recurrence_duration(master_item_for_pattern.recurrence)
                        except Exception:
                            pass

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
                        'Action': 'Report' if report_only else 'Delete',
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
                        'Action': 'Report' if report_only else 'Delete',
                        'Status': 'Pending',
                        'Details': ''
                    }

                if report_only:
                    self.log(f"  [报告] 发现: {item.subject}")
                    row['Status'] = 'Skipped'
                else:
                    self.log(f"  正在删除: {item.subject}")
                    item.delete()
                    row['Status'] = 'Success'
                
                with csv_lock:
                    writer.writerow(row)
                    # csvfile.flush()

        except Exception as e:
            self.log(f"  处理用户 {target_email} 出错: {e}", "ERROR")
            self.log(f"  Traceback: {traceback.format_exc()}", is_advanced=True)
            with csv_lock:
                writer.writerow({'UserPrincipalName': target_email, 'Status': 'Error', 'Details': str(e)})

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

                # Extract variables for threads
                start_date_str = self._normalize_date_input(self.criteria_start_date.get())
                end_date_str = self._normalize_date_input(self.criteria_end_date.get())
                
                # Update UI vars once if needed
                if start_date_str: self.criteria_start_date.set(start_date_str)
                if end_date_str: self.criteria_end_date.set(end_date_str)

                criteria_sender = self.criteria_sender.get()
                criteria_msg_id = self.criteria_msg_id.get()
                criteria_subject = self.criteria_subject.get()
                criteria_body = self.criteria_body.get()
                meeting_only_cancelled = self.meeting_only_cancelled_var.get()
                meeting_scope = self.meeting_scope_var.get()
                report_only = self.report_only_var.get()
                log_level = self.log_level_var.get()
                
                csv_lock = threading.Lock()

                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = []
                    for target_email in users:
                        futures.append(executor.submit(
                            self.process_single_user_ews,
                            target_email, creds, config, auth_type, use_auto, target_type,
                            start_date_str, end_date_str, criteria_sender, criteria_msg_id,
                            criteria_subject, criteria_body, meeting_only_cancelled, meeting_scope,
                            report_only, writer, csv_lock, log_level
                        ))
                    
                    for future in futures:
                        try:
                            future.result()
                        except Exception as e:
                            self.log(f"Task Error: {e}", "ERROR")

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