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
from ctypes import wintypes
from concurrent.futures import ThreadPoolExecutor
import calendar
import webbrowser
import base64
import io
import random
from requests.adapters import HTTPAdapter

APP_VERSION = "v1.11.2"

# Use a stable AppUserModelID on Windows. If this changes per version, Windows may keep
# showing a cached/pinned icon from an older shortcut.
WINDOWS_APP_USER_MODEL_ID = "UniversalEmailCleaner"
GITHUB_PROJECT_URL = "https://github.com/andylu1988/UniversalEmailCleaner"
GITHUB_PROFILE_URL = "https://github.com/andylu1988"


_thread_local = threading.local()


def _get_pooled_session() -> requests.Session:
    sess = getattr(_thread_local, 'session', None)
    if sess is not None:
        return sess

    sess = requests.Session()
    adapter = HTTPAdapter(pool_connections=64, pool_maxsize=64, max_retries=0)
    sess.mount('https://', adapter)
    sess.mount('http://', adapter)
    _thread_local.session = sess
    return sess


def _dpapi_protect_text(plain_text: str) -> str | None:
    if sys.platform != 'win32':
        return None
    try:
        class DATA_BLOB(ctypes.Structure):
            _fields_ = [
                ("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_byte)),
            ]

        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        CRYPTPROTECT_UI_FORBIDDEN = 0x01

        data = plain_text.encode('utf-8')
        in_blob = DATA_BLOB(len(data), ctypes.cast(ctypes.create_string_buffer(data), ctypes.POINTER(ctypes.c_byte)))
        out_blob = DATA_BLOB()

        # Encrypt for current user (no LOCAL_MACHINE flag)
        if not crypt32.CryptProtectData(
            ctypes.byref(in_blob),
            "UniversalEmailCleaner Graph Token",
            None,
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            ctypes.byref(out_blob),
        ):
            return None

        try:
            protected_bytes = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        finally:
            kernel32.LocalFree(out_blob.pbData)

        return base64.b64encode(protected_bytes).decode('ascii')
    except Exception:
        return None


def _dpapi_unprotect_text(protected_b64: str) -> str | None:
    if sys.platform != 'win32':
        return None
    try:
        class DATA_BLOB(ctypes.Structure):
            _fields_ = [
                ("cbData", wintypes.DWORD),
                ("pbData", ctypes.POINTER(ctypes.c_byte)),
            ]

        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32
        CRYPTPROTECT_UI_FORBIDDEN = 0x01

        protected_bytes = base64.b64decode(protected_b64)
        in_blob = DATA_BLOB(len(protected_bytes), ctypes.cast(ctypes.create_string_buffer(protected_bytes), ctypes.POINTER(ctypes.c_byte)))
        out_blob = DATA_BLOB()

        if not crypt32.CryptUnprotectData(
            ctypes.byref(in_blob),
            None,
            None,
            None,
            None,
            CRYPTPROTECT_UI_FORBIDDEN,
            ctypes.byref(out_blob),
        ):
            return None

        try:
            plain_bytes = ctypes.string_at(out_blob.pbData, out_blob.cbData)
        finally:
            kernel32.LocalFree(out_blob.pbData)

        return plain_bytes.decode('utf-8', errors='strict')
    except Exception:
        return None


def resource_path(relative_path: str) -> str:
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, relative_path)


def _win32_force_window_icon(hwnd: int, ico_path: str) -> None:
    """Force-set the window/taskbar icon on Windows via WM_SETICON.

    Tk's iconbitmap/iconphoto is sometimes ignored by the taskbar on some Windows builds.
    """
    if sys.platform != 'win32':
        return
    try:
        user32 = ctypes.windll.user32
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x0010
        LR_DEFAULTSIZE = 0x0040
        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1

        # Navigate up to the real top-level window (Tk frame HWND is a child)
        GA_ROOT = 2
        real_hwnd = user32.GetAncestor(hwnd, GA_ROOT) or hwnd

        hicon_big = user32.LoadImageW(None, ico_path, IMAGE_ICON, 48, 48, LR_LOADFROMFILE)
        hicon_small = user32.LoadImageW(None, ico_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        if not hicon_big:
            hicon_big = user32.LoadImageW(None, ico_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
        if not hicon_small:
            hicon_small = hicon_big
        if hicon_big:
            user32.SendMessageW(real_hwnd, WM_SETICON, ICON_BIG, hicon_big)
        if hicon_small:
            user32.SendMessageW(real_hwnd, WM_SETICON, ICON_SMALL, hicon_small)
    except Exception:
        pass

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
    from exchangelib import Account, Credentials, Configuration, DELEGATE, IMPERSONATION, Message, Mailbox, EWSDateTime, CalendarItem, NTLM, BASIC
    from exchangelib import ItemId as EwsItemId
    try:
        from exchangelib.items import DeleteType
    except Exception:
        DeleteType = None
    try:
        from exchangelib import OAuth2Credentials, OAuth2LegacyCredentials
    except ImportError:
        OAuth2Credentials = None
        OAuth2LegacyCredentials = None
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
    DeleteType = None
    OAuth2Credentials = None
    OAuth2LegacyCredentials = None
    DELEGATE = None
    IMPERSONATION = None
    NTLM = "NTLM"
    BASIC = "basic"

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

        try:
            self.top.iconbitmap(resource_path("graph-mail-delete.ico"))
        except Exception:
            pass
        
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


def _graph_weekday_cn(day):
    mapping = {
        'monday': '周一', 'tuesday': '周二', 'wednesday': '周三',
        'thursday': '周四', 'friday': '周五', 'saturday': '周六', 'sunday': '周日'
    }
    if not day:
        return ""
    return mapping.get(str(day).lower(), str(day))


def format_graph_recurrence_pattern(pattern):
    """Return (pattern_name, pattern_details) as human-readable strings."""
    if not isinstance(pattern, dict) or not pattern:
        return "", ""

    p_type = (pattern.get('type') or '').strip()
    interval = pattern.get('interval')

    type_cn = {
        'daily': '按天',
        'weekly': '按周',
        'absolutemonthly': '按月(固定)',
        'relativemonthly': '按月(相对)',
        'absoluteyearly': '按年(固定)',
        'relativeyearly': '按年(相对)',
    }.get(p_type.lower(), p_type)

    details = []
    if interval:
        details.append(f"间隔={interval}")

    days = pattern.get('daysOfWeek') or []
    if days:
        days_cn = ",".join(_graph_weekday_cn(d) for d in days)
        details.append(f"星期={days_cn}")

    if pattern.get('dayOfMonth'):
        details.append(f"日期={pattern.get('dayOfMonth')}日")

    if pattern.get('month'):
        details.append(f"月份={pattern.get('month')}月")

    if pattern.get('index'):
        idx_map = {
            'first': '第一个', 'second': '第二个', 'third': '第三个', 'fourth': '第四个', 'last': '最后一个'
        }
        details.append(f"索引={idx_map.get(str(pattern.get('index')).lower(), pattern.get('index'))}")

    details_str = ", ".join(details)
    return type_cn, (f"{type_cn}: {details_str}" if details_str else type_cn)


def format_graph_recurrence_range(rng):
    """Return (duration_str, is_endless_str) as human-readable strings."""
    if not isinstance(rng, dict) or not rng:
        return "", ""

    r_type = (rng.get('type') or '').strip().lower()
    start_date = rng.get('startDate')
    end_date = rng.get('endDate')
    number = rng.get('numberOfOccurrences')
    tz = rng.get('recurrenceTimeZone')

    parts = []
    if start_date:
        parts.append(f"开始: {start_date}")

    is_endless = ""
    if r_type == 'noend':
        parts.append("无限期")
        is_endless = "True"
    elif r_type == 'enddate':
        if end_date:
            parts.append(f"结束: {end_date}")
        is_endless = "False"
    elif r_type == 'numbered':
        if number:
            parts.append(f"共 {number} 次")
        is_endless = "False"
    else:
        if end_date:
            parts.append(f"结束: {end_date}")

    if tz:
        parts.append(f"时区: {tz}")

    return "; ".join(parts), is_endless


def decode_graph_goid_base64_to_hex(goid_b64):
    """Graph MAPI Binary extended properties return base64; convert to hex for easier comparison."""
    if not goid_b64:
        return ""
    try:
        raw = base64.b64decode(goid_b64)
        return raw.hex().upper()
    except Exception:
        return ""


def redact_sensitive_headers(headers, save_authorization=False):
    """Mask sensitive auth material before writing debug logs."""
    if not isinstance(headers, dict) or not headers:
        return {}

    redacted = {}
    for k, v in headers.items():
        key_lower = (k or "").lower()
        if key_lower == "authorization":
            if save_authorization:
                redacted[k] = v
            else:
                if isinstance(v, str) and v.lower().startswith("bearer "):
                    redacted[k] = "Bearer ***"
                else:
                    redacted[k] = "***"
        else:
            redacted[k] = v
    return redacted


def format_graph_meeting_response_status(user_email, user_role, organizer_email, attendees, item_response_status):
    """Organizer => attendee responses; Attendee => self responseStatus."""
    role = (user_role or "").strip().lower()
    user_email_l = (user_email or "").strip().lower()
    organizer_email_l = (organizer_email or "").strip().lower()

    if role == "organizer":
        parts = []
        for a in (attendees or []):
            addr = ((a.get("emailAddress") or {}).get("address") or "").strip()
            if not addr:
                continue
            addr_l = addr.lower()
            if addr_l in (user_email_l, organizer_email_l):
                continue
            resp = ((a.get("status") or {}).get("response") or "").strip()
            parts.append(f"{addr}:{resp}" if resp else f"{addr}:")
        return ";".join(parts)

    # attendee (or unknown): Graph event.responseStatus is the current user's status
    return ((item_response_status or {}).get("response") or "").strip()


class Logger:
    def __init__(self, log_area, log_dir):
        self.log_area = log_area
        self.log_dir = log_dir
        self.level = "NORMAL" # NORMAL / ADVANCED / EXPERT
        self.file_lock = threading.Lock()

    def _level_rank(self, level):
        mapping = {"NORMAL": 0, "ADVANCED": 1, "EXPERT": 2}
        return mapping.get((level or "").upper(), 0)

    def _get_log_file_path(self, kind="app"):
        date_str = datetime.now().strftime("%Y-%m-%d")
        if kind == "advanced":
            return os.path.join(self.log_dir, f"app_advanced_{date_str}.log")
        if kind == "expert":
            return os.path.join(self.log_dir, f"app_expert_{date_str}.log")
        return os.path.join(self.log_dir, f"app_{date_str}.log")

    def get_current_debug_log_path(self):
        if self.level == "ADVANCED":
            return self._get_log_file_path("advanced")
        if self.level == "EXPERT":
            return self._get_log_file_path("expert")
        return ""

    def set_level(self, level):
        # Accept GUI values: Normal/Advanced/Expert
        val = (level or "NORMAL").upper()
        if val in ("NORMAL", "ADVANCED", "EXPERT"):
            self.level = val
        else:
            self.level = "NORMAL"

    def log(self, message, level="INFO", is_advanced=False):
        # If message is advanced but current level is NORMAL, skip
        if is_advanced and self._level_rank(self.level) < self._level_rank("ADVANCED"):
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
            with self.file_lock:
                # Always write normal log
                with open(self._get_log_file_path("app"), "a", encoding="utf-8") as f:
                    f.write(full_msg + "\n")

                # Advanced/Expert go to their own debug logs (separate from normal)
                if is_advanced:
                    dbg_path = self.get_current_debug_log_path()
                    if dbg_path:
                        with open(dbg_path, "a", encoding="utf-8") as f:
                            f.write(full_msg + "\n")
        except Exception:
            pass

    def log_to_file_only(self, message, min_level="ADVANCED"):
        """Writes directly to debug file (advanced/expert), skipping GUI."""
        if self._level_rank(self.level) < self._level_rank(min_level):
            return

        dbg_path = self.get_current_debug_log_path()
        if not dbg_path:
            return

        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        full_msg = f"[{timestamp}] [DEBUG_DATA] {message}"
        try:
            with self.file_lock:
                with open(dbg_path, "a", encoding="utf-8") as f:
                    f.write(full_msg + "\n")
        except Exception:
            pass

class EwsTraceAdapter(NoVerifyHTTPAdapter):
    logger = None
    log_responses = True  # Default to True, can be disabled for "Advanced" mode
    response_log_path = None
    
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
                    log_path = self.response_log_path
                    if not log_path:
                        docs_dir = os.path.join(os.path.expanduser("~"), "Documents", "UniversalEmailCleaner")
                        os.makedirs(docs_dir, exist_ok=True)
                        date_str = datetime.now().strftime("%Y-%m-%d")
                        log_path = os.path.join(docs_dir, f"ews_getitem_responses_expert_{date_str}.log")
                    
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

        # Improve Windows taskbar icon behavior by setting AppUserModelID
        # and explicitly setting a Tk iconphoto (some Windows builds ignore iconbitmap for taskbar).
        try:
            if sys.platform == 'win32':
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOWS_APP_USER_MODEL_ID)
        except Exception:
            pass

        try:
            self.root.iconbitmap(resource_path("graph-mail-delete.ico"))
        except Exception:
            pass

        try:
            if Image is not None and ImageTk is not None:
                ico_path = resource_path("graph-mail-delete.ico")
                img = Image.open(ico_path).convert("RGBA")
                img = img.resize((64, 64))
                self._app_icon_photo = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, self._app_icon_photo)
        except Exception:
            pass

        # Force taskbar icon (Win32) for stubborn Windows builds.
        try:
            if sys.platform == 'win32':
                self.root.update_idletasks()
                _win32_force_window_icon(int(self.root.winfo_id()), resource_path("graph-mail-delete.ico"))
        except Exception:
            pass
        
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
        self.graph_save_auth_token_var = tk.BooleanVar(value=False)

        def on_log_level_change_request():
            # Shared handler for Tools menu and main UI
            val = self.log_level_var.get()
            if val == "Expert":
                confirm = messagebox.askyesno("警告", "日志排错专用，日志量会很大且包含敏感信息，慎选！\n\n确认开启专家模式吗？")
                if not confirm:
                    self.log_level_var.set("Normal")
                    return
            # Update runtime logger immediately
            try:
                self.logger.set_level(self.log_level_var.get())
            except Exception:
                pass

            # If leaving Expert, force auth token saving OFF to stay safe
            if self.log_level_var.get() != "Expert":
                try:
                    self.graph_save_auth_token_var.set(False)
                except Exception:
                    pass

        def on_graph_save_auth_toggle():
            if not self.graph_save_auth_token_var.get():
                return

            if self.log_level_var.get() != "Expert":
                messagebox.showwarning("提示", "该选项仅在 Expert 日志级别下生效。\n\n请先将日志级别切换为 Expert。")
                self.graph_save_auth_token_var.set(False)
                return

            confirm = messagebox.askyesno(
                "高风险警告",
                "开启后会在 Expert 日志中保存 Authorization Token，存在敏感信息泄漏风险。\n\n确认开启吗？"
            )
            if not confirm:
                self.graph_save_auth_token_var.set(False)
        
        log_menu.add_radiobutton(label="默认 (Default)", variable=self.log_level_var, value="Normal", command=on_log_level_change_request)
        log_menu.add_radiobutton(label="高级 (Advanced - 记录 Graph/EWS 请求)", variable=self.log_level_var, value="Advanced", command=on_log_level_change_request)
        log_menu.add_radiobutton(label="专家 (Expert - 记录 Graph/EWS 请求和响应)", variable=self.log_level_var, value="Expert", command=on_log_level_change_request)

        log_menu.add_separator()
        log_menu.add_checkbutton(
            label="Graph Expert 保存 Authorization Token (危险)",
            variable=self.graph_save_auth_token_var,
            command=on_graph_save_auth_toggle,
        )
        
        tools_menu.add_cascade(label="日志配置 (Log Level)", menu=log_menu)
        menubar.add_cascade(label="工具 (Tools)", menu=tools_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于 (About)", command=self.show_about)
        menubar.add_cascade(label="帮助 (Help)", menu=help_menu)
        root.config(menu=menubar)

        # --- Variables ---
        # Graph Config
        self.graph_auth_mode_var = tk.StringVar(value="Auto") # Auto (Cert) or Manual (Secret) or Token
        self.app_id_var = tk.StringVar()
        self.tenant_id_var = tk.StringVar()
        self.thumbprint_var = tk.StringVar()
        self.client_secret_var = tk.StringVar()
        self.graph_env_var = tk.StringVar(value="Global")
        self.graph_token_var = tk.StringVar()
        self.graph_cache_token_var = tk.BooleanVar(value=True)
        self._graph_token_protected_cache = ""

        # EWS Config
        self.ews_server_var = tk.StringVar()
        self.ews_user_var = tk.StringVar()
        self.ews_pass_var = tk.StringVar()
        self.ews_auth_type_var = tk.StringVar(value="Impersonation") # Impersonation or Delegate
        self.ews_use_autodiscover = tk.BooleanVar(value=True)
        # EWS Auth Method: NTLM / Basic (Legacy) / OAuth2 (Modern) / Token
        self.ews_auth_method_var = tk.StringVar(value="NTLM")
        self.ews_oauth_app_id_var = tk.StringVar()
        self.ews_oauth_tenant_id_var = tk.StringVar()
        self.ews_oauth_secret_var = tk.StringVar()
        self.ews_token_var = tk.StringVar()
        self.ews_cache_token_var = tk.BooleanVar(value=True)
        self._ews_token_protected_cache = ""

        # Cleanup Config
        self.source_type_var = tk.StringVar(value="Graph") # Graph or EWS
        self.csv_path_var = tk.StringVar()
        self.target_single_email_var = tk.StringVar()
        self.report_only_var = tk.BooleanVar(value=True)
        self.permanent_delete_var = tk.BooleanVar(value=False)
        # Soft delete: move items to Deleted Items (best-effort).
        # Default OFF to preserve existing behavior.
        self.soft_delete_var = tk.BooleanVar(value=False)
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

        # Email folder scan scope
        # Keep existing behavior by default:
        # - Graph: All mailbox (messages)
        # - EWS: Inbox only
        self.mail_folder_scope_var = tk.StringVar(value="自动 (Auto)")

        # Progress tracking
        self._progress_total = 0
        self._progress_done = 0

        # Scan results cache for interactive deletion
        self._scan_results_data: list[dict] = []   # rows from CSV
        self._scan_results_columns: list[str] = [] # column headers
        self._last_report_path: str = ""            # path of most recent report CSV
        self._scan_checked: dict[str, bool] = {}   # iid -> checked

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

        self.tab_results = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_results, text="3. 扫描结果")

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
        self.logger.set_level(self.log_level_var.get())

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
        self.build_results_tab()

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
        self._last_report_path = path
        def _update():
            self.report_link_lbl.config(text=f"最新报告: {path}")
            self.report_link_lbl.bind("<Button-1>", lambda e: os.startfile(path) if os.path.exists(path) else None)
        self.root.after(0, _update)

    def show_history(self):
        history_window = tk.Toplevel(self.root)
        history_window.title("版本历史")
        history_window.geometry("600x400")

        try:
            history_window.iconbitmap(resource_path("graph-mail-delete.ico"))
        except Exception:
            pass
        
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

        try:
            about.iconbitmap(resource_path("graph-mail-delete.ico"))
        except Exception:
            pass

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
                    try:
                        self.graph_cache_token_var.set(bool(config.get('graph_cache_token', True)))
                    except Exception:
                        pass
                    try:
                        self._graph_token_protected_cache = config.get('graph_token_protected', '') or ''
                    except Exception:
                        self._graph_token_protected_cache = ''
                    # EWS
                    self.ews_server_var.set(config.get('ews_server', ''))
                    self.ews_user_var.set(config.get('ews_user', ''))
                    self.ews_use_autodiscover.set(config.get('ews_autodiscover', True))
                    self.ews_auth_type_var.set(config.get('ews_auth_type', 'Impersonation'))
                    self.ews_auth_method_var.set(config.get('ews_auth_method', 'NTLM'))
                    self.ews_oauth_app_id_var.set(config.get('ews_oauth_app_id', ''))
                    self.ews_oauth_tenant_id_var.set(config.get('ews_oauth_tenant_id', ''))
                    self.ews_oauth_secret_var.set(config.get('ews_oauth_secret', ''))
                    try:
                        self.ews_cache_token_var.set(bool(config.get('ews_cache_token', True)))
                    except Exception:
                        pass
                    try:
                        self._ews_token_protected_cache = config.get('ews_token_protected', '') or ''
                    except Exception:
                        self._ews_token_protected_cache = ''
                    # Common
                    self.source_type_var.set(config.get('source_type', 'EWS')) # Default to EWS if not set
                    self.csv_path_var.set(config.get('csv_path', ''))
                    try:
                        self.target_single_email_var.set(config.get('target_single_email', ''))
                    except Exception:
                        pass
                    try:
                        self.mail_folder_scope_var.set(config.get('mail_folder_scope', '自动 (Auto)'))
                    except Exception:
                        pass
                    try:
                        self.permanent_delete_var.set(bool(config.get('permanent_delete', False)))
                    except Exception:
                        pass
                    try:
                        self.soft_delete_var.set(bool(config.get('soft_delete', False)))
                    except Exception:
                        pass
                    try:
                        if bool(self.permanent_delete_var.get()) and bool(self.soft_delete_var.get()):
                            self.soft_delete_var.set(False)
                    except Exception:
                        pass
                    self.log(">>> 配置已加载。")
            except Exception as e:
                self.log(f"X 加载配置失败: {e}", "ERROR")
        else:
            # No config file, set default to EWS
            self.source_type_var.set("EWS")
            self.toggle_connection_ui()

    def save_config(self):
        # Keep DPAPI-protected token blob separate from plain UI value.
        # Graph token caching
        try:
            cache_enabled = bool(self.graph_cache_token_var.get())
        except Exception:
            cache_enabled = True

        token_ui = ""
        try:
            token_ui = (self.graph_token_var.get() or '').strip()
        except Exception:
            token_ui = ""
        if token_ui.lower().startswith('bearer '):
            token_ui = token_ui[7:].strip()

        if not cache_enabled:
            self._graph_token_protected_cache = ""
        else:
            if token_ui:
                protected = _dpapi_protect_text(token_ui)
                if protected:
                    self._graph_token_protected_cache = protected

        # EWS token caching
        try:
            ews_cache_enabled = bool(self.ews_cache_token_var.get())
        except Exception:
            ews_cache_enabled = True

        ews_token_ui = ""
        try:
            ews_token_ui = (self.ews_token_var.get() or '').strip()
        except Exception:
            ews_token_ui = ""
        if ews_token_ui.lower().startswith('bearer '):
            ews_token_ui = ews_token_ui[7:].strip()

        if not ews_cache_enabled:
            self._ews_token_protected_cache = ""
        else:
            if ews_token_ui:
                protected = _dpapi_protect_text(ews_token_ui)
                if protected:
                    self._ews_token_protected_cache = protected

        config = {
            'graph_auth_mode': self.graph_auth_mode_var.get(),
            'app_id': self.app_id_var.get(),
            'tenant_id': self.tenant_id_var.get(),
            'thumbprint': self.thumbprint_var.get(),
            'client_secret': self.client_secret_var.get(),
            'graph_env': self.graph_env_var.get(),
            'graph_cache_token': bool(self.graph_cache_token_var.get()),
            'graph_token_protected': self._graph_token_protected_cache,
            'ews_server': self.ews_server_var.get(),
            'ews_user': self.ews_user_var.get(),
            'ews_autodiscover': self.ews_use_autodiscover.get(),
            'ews_auth_type': self.ews_auth_type_var.get(),
            'ews_auth_method': self.ews_auth_method_var.get(),
            'ews_oauth_app_id': self.ews_oauth_app_id_var.get(),
            'ews_oauth_tenant_id': self.ews_oauth_tenant_id_var.get(),
            'ews_oauth_secret': self.ews_oauth_secret_var.get(),
            'ews_cache_token': bool(self.ews_cache_token_var.get()),
            'ews_token_protected': self._ews_token_protected_cache,
            'source_type': self.source_type_var.get(),
            'csv_path': self.csv_path_var.get(),
            'target_single_email': self.target_single_email_var.get(),
            'mail_folder_scope': self.mail_folder_scope_var.get(),
            'permanent_delete': bool(self.permanent_delete_var.get()),
            'soft_delete': bool(self.soft_delete_var.get()),
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
        self.ews_frame = ttk.LabelFrame(main_frame, text="EWS 配置 (Exchange On-Premise / Online)")
        self.ews_frame.pack(fill="x", pady=5, ipady=5)

        # --- EWS Auth Method Selection ---
        ews_method_frame = ttk.Frame(self.ews_frame)
        ews_method_frame.pack(fill="x", padx=10, pady=(5, 2))
        ttk.Label(ews_method_frame, text="验证方式:").pack(side="left")
        ttk.Radiobutton(ews_method_frame, text="NTLM", variable=self.ews_auth_method_var, value="NTLM", command=self.toggle_ews_auth_ui).pack(side="left", padx=5)
        ttk.Radiobutton(ews_method_frame, text="Basic (Legacy)", variable=self.ews_auth_method_var, value="Basic", command=self.toggle_ews_auth_ui).pack(side="left", padx=5)
        ttk.Radiobutton(ews_method_frame, text="OAuth2 (Modern)", variable=self.ews_auth_method_var, value="OAuth2", command=self.toggle_ews_auth_ui).pack(side="left", padx=5)
        ttk.Radiobutton(ews_method_frame, text="直接输入 Token", variable=self.ews_auth_method_var, value="Token", command=self.toggle_ews_auth_ui).pack(side="left", padx=5)

        # --- EWS NTLM Frame (same fields as Basic — UPN + Password) ---
        self.ews_ntlm_frame = ttk.Frame(self.ews_frame)
        ews_ntlm_grid = ttk.Frame(self.ews_ntlm_frame)
        ews_ntlm_grid.pack(anchor="w", padx=10, pady=5)
        grid_opts = {"sticky": "w", "padx": 5, "pady": 5}

        ttk.Label(ews_ntlm_grid, text="管理员账号 (UPN 或 DOMAIN\\User):").grid(row=0, column=0, **grid_opts)
        ttk.Entry(ews_ntlm_grid, textvariable=self.ews_user_var, width=40).grid(row=0, column=1, **grid_opts)

        ttk.Label(ews_ntlm_grid, text="管理员密码:").grid(row=1, column=0, **grid_opts)
        ttk.Entry(ews_ntlm_grid, textvariable=self.ews_pass_var, show="*", width=40).grid(row=1, column=1, **grid_opts)

        ttk.Label(ews_ntlm_grid, text="NTLM 适用于本地 Exchange 或混合部署 (域账号)。").grid(row=2, column=1, **grid_opts)

        # --- EWS Basic (Legacy) Frame ---
        self.ews_basic_frame = ttk.Frame(self.ews_frame)
        ews_basic_grid = ttk.Frame(self.ews_basic_frame)
        ews_basic_grid.pack(anchor="w", padx=10, pady=5)

        ttk.Label(ews_basic_grid, text="管理员账号 (UPN):").grid(row=0, column=0, **grid_opts)
        ttk.Entry(ews_basic_grid, textvariable=self.ews_user_var, width=40).grid(row=0, column=1, **grid_opts)

        ttk.Label(ews_basic_grid, text="管理员密码:").grid(row=1, column=0, **grid_opts)
        ttk.Entry(ews_basic_grid, textvariable=self.ews_pass_var, show="*", width=40).grid(row=1, column=1, **grid_opts)

        ttk.Label(ews_basic_grid, text="Basic Auth 已被 Microsoft 365 弃用，仅适用于旧版本 Exchange。").grid(row=2, column=1, **grid_opts)

        # --- EWS OAuth2 (Modern) Frame ---
        self.ews_oauth2_frame = ttk.Frame(self.ews_frame)
        ews_oauth2_grid = ttk.Frame(self.ews_oauth2_frame)
        ews_oauth2_grid.pack(anchor="w", padx=10, pady=5)

        ttk.Label(ews_oauth2_grid, text="Application ID:").grid(row=0, column=0, **grid_opts)
        ttk.Entry(ews_oauth2_grid, textvariable=self.ews_oauth_app_id_var, width=50).grid(row=0, column=1, **grid_opts)

        ttk.Label(ews_oauth2_grid, text="Tenant ID:").grid(row=1, column=0, **grid_opts)
        ttk.Entry(ews_oauth2_grid, textvariable=self.ews_oauth_tenant_id_var, width=50).grid(row=1, column=1, **grid_opts)

        ttk.Label(ews_oauth2_grid, text="Client Secret:").grid(row=2, column=0, **grid_opts)
        ttk.Entry(ews_oauth2_grid, textvariable=self.ews_oauth_secret_var, width=50, show="*").grid(row=2, column=1, **grid_opts)

        ttk.Label(
            ews_oauth2_grid,
            text="使用 Client Credentials 流程获取 EWS 访问令牌 (需 full_access_as_app 权限)",
        ).grid(row=3, column=1, **grid_opts)

        # --- EWS Token Frame ---
        self.ews_token_frame = ttk.Frame(self.ews_frame)
        ews_token_grid = ttk.Frame(self.ews_token_frame)
        ews_token_grid.pack(anchor="w", fill="x", expand=True, padx=10, pady=5)

        ttk.Label(ews_token_grid, text="Access Token (Bearer):").grid(row=0, column=0, **grid_opts)
        ttk.Entry(ews_token_grid, textvariable=self.ews_token_var, width=60, show="*").grid(row=0, column=1, **grid_opts)

        ttk.Checkbutton(
            ews_token_grid,
            text="加密缓存 Token (Windows 当前用户 / DPAPI)",
            variable=self.ews_cache_token_var,
        ).grid(row=1, column=1, **grid_opts)

        ttk.Label(
            ews_token_grid,
            text="提示：Token 通常有过期时间。留空则尝试使用已缓存 Token。",
        ).grid(row=2, column=1, **grid_opts)

        # --- EWS Common Settings (Server / Autodiscover / Access Type) ---
        ews_common_frame = ttk.Frame(self.ews_frame)
        ews_common_frame.pack(fill="x", padx=10, pady=2)
        ews_grid = ttk.Frame(ews_common_frame)
        ews_grid.pack(anchor="w")

        ttk.Label(ews_grid, text="EWS 服务器:").grid(row=0, column=0, **grid_opts)
        self.entry_ews_server = ttk.Entry(ews_grid, textvariable=self.ews_server_var, width=40)
        self.entry_ews_server.grid(row=0, column=1, **grid_opts)
        
        self.chk_ews_auto = ttk.Checkbutton(ews_grid, text="使用自动发现 (Autodiscover)", variable=self.ews_use_autodiscover, 
                                   command=lambda: self.entry_ews_server.config(state='disabled' if self.ews_use_autodiscover.get() else 'normal'))
        self.chk_ews_auto.grid(row=0, column=2, padx=5)

        ttk.Label(ews_grid, text="访问类型:").grid(row=1, column=0, **grid_opts)
        auth_frame = ttk.Frame(ews_grid)
        auth_frame.grid(row=1, column=1, sticky="w")
        ttk.Radiobutton(auth_frame, text="模拟 (Impersonation)", variable=self.ews_auth_type_var, value="Impersonation").pack(side="left", padx=5)
        ttk.Radiobutton(auth_frame, text="代理 (Delegate)", variable=self.ews_auth_type_var, value="Delegate").pack(side="left", padx=5)

        ttk.Button(self.ews_frame, text="测试 EWS 连接", command=self.test_ews_connection).pack(anchor="w", padx=10, pady=10)

        # Show correct auth frame initially
        self.toggle_ews_auth_ui()

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
        ttk.Radiobutton(mode_frame, text="直接输入 Token", variable=self.graph_auth_mode_var, value="Token", command=self.toggle_graph_ui).pack(side="left", padx=5)

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

        # Graph Token Frame
        self.graph_token_frame = ttk.Frame(self.graph_frame)
        token_grid = ttk.Frame(self.graph_token_frame)
        token_grid.pack(anchor="w", fill="x", expand=True)

        ttk.Label(token_grid, text="Access Token (Bearer):").grid(row=0, column=0, **grid_opts)
        ttk.Entry(token_grid, textvariable=self.graph_token_var, width=60, show="*").grid(row=0, column=1, **grid_opts)

        ttk.Checkbutton(
            token_grid,
            text="加密缓存 Token (Windows 当前用户 / DPAPI)",
            variable=self.graph_cache_token_var,
        ).grid(row=1, column=1, **grid_opts)

        ttk.Label(
            token_grid,
            text="提示：Token 通常有过期时间。留空则尝试使用已缓存 Token。",
        ).grid(row=2, column=1, **grid_opts)

        # Initial Toggle
        self.toggle_connection_ui()

    def toggle_connection_ui(self):
        mode = self.source_type_var.get()
        if mode == "EWS":
            self._enable_frame(self.ews_frame)
            self._disable_frame(self.graph_frame)
            self.toggle_ews_auth_ui()
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
            try:
                self.graph_token_frame.pack_forget()
            except Exception:
                pass
            self.graph_auto_frame.pack(fill="x", padx=10, pady=5)
        elif mode == "Manual":
            self.graph_auto_frame.pack_forget()
            try:
                self.graph_token_frame.pack_forget()
            except Exception:
                pass
            self.graph_manual_frame.pack(fill="x", padx=10, pady=5)
        else:  # Token
            self.graph_auto_frame.pack_forget()
            self.graph_manual_frame.pack_forget()
            self.graph_token_frame.pack(fill="x", padx=10, pady=5)

    def toggle_ews_auth_ui(self):
        """Show/hide EWS auth sub-frames based on selected auth method."""
        method = self.ews_auth_method_var.get()
        for f in (self.ews_ntlm_frame, self.ews_basic_frame, self.ews_oauth2_frame, self.ews_token_frame):
            try:
                f.pack_forget()
            except Exception:
                pass
        if method == "NTLM":
            self.ews_ntlm_frame.pack(fill="x", padx=10, pady=2)
        elif method == "Basic":
            self.ews_basic_frame.pack(fill="x", padx=10, pady=2)
        elif method == "OAuth2":
            self.ews_oauth2_frame.pack(fill="x", padx=10, pady=2)
        else:  # Token
            self.ews_token_frame.pack(fill="x", padx=10, pady=2)

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

    def _get_ews_access_token_oauth2(self):
        """Obtain an EWS access token via OAuth2 Client Credentials flow."""
        app_id = self.ews_oauth_app_id_var.get().strip()
        tenant_id = self.ews_oauth_tenant_id_var.get().strip()
        secret = self.ews_oauth_secret_var.get().strip()
        if not app_id or not tenant_id or not secret:
            raise Exception("OAuth2 模式需要 Application ID, Tenant ID 和 Client Secret。")
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": app_id,
            "client_secret": secret,
            "scope": "https://outlook.office365.com/.default",
        }
        resp = requests.post(token_url, data=data)
        if resp.status_code != 200:
            raise Exception(f"OAuth2 token 获取失败: {resp.status_code} {resp.text}")
        token = resp.json().get("access_token")
        if not token:
            raise Exception("OAuth2 返回中缺少 access_token")
        return token

    def _get_ews_credentials(self):
        """Return (credentials_or_None, access_token_or_None) based on current EWS auth method."""
        method = self.ews_auth_method_var.get()
        if method in ("Basic", "NTLM"):
            user = self.ews_user_var.get()
            pwd = self.ews_pass_var.get()
            if not user or not pwd:
                raise Exception(f"{method} 模式需要用户名和密码。")
            return Credentials(user, pwd), None

        if method == "OAuth2":
            token = self._get_ews_access_token_oauth2()
            if OAuth2Credentials is not None:
                creds = OAuth2Credentials(
                    client_id=self.ews_oauth_app_id_var.get().strip(),
                    client_secret=self.ews_oauth_secret_var.get().strip(),
                    tenant_id=self.ews_oauth_tenant_id_var.get().strip(),
                )
                return creds, token
            return None, token

        # Token mode
        token_ui = (self.ews_token_var.get() or '').strip()
        if token_ui.lower().startswith('bearer '):
            token_ui = token_ui[7:].strip()
        token = token_ui
        from_cache = False
        if not token:
            protected = getattr(self, '_ews_token_protected_cache', '') or ''
            if protected:
                token = _dpapi_unprotect_text(protected) or ''
                from_cache = True
        if not token:
            raise Exception("未提供 Token，且无法从缓存读取 Token")
        # Refresh cache
        try:
            cache_enabled = bool(self.ews_cache_token_var.get())
        except Exception:
            cache_enabled = True
        if cache_enabled and (not from_cache) and token:
            protected_new = _dpapi_protect_text(token)
            if protected_new:
                self._ews_token_protected_cache = protected_new
                try:
                    self.save_config()
                except Exception:
                    pass
        return None, token

    def _test_ews(self):
        try:
            self.log(">>> 正在测试 EWS 连接...")
            method = self.ews_auth_method_var.get()
            server = self._clean_server_address(self.ews_server_var.get())
            use_auto = self.ews_use_autodiscover.get()

            creds, token = self._get_ews_credentials()

            # Determine exchangelib auth_type for NTLM vs Basic
            ews_proto_auth_type = None
            if method == "NTLM":
                ews_proto_auth_type = NTLM
            elif method == "Basic":
                ews_proto_auth_type = BASIC

            # Determine primary_smtp_address for testing
            test_email = self.ews_user_var.get().strip() or self.target_single_email_var.get().strip()
            if not test_email:
                raise Exception("请填写管理员账号 (UPN) 或单个目标邮箱，以便测试连接。")

            if token and creds is None:
                # Pure bearer token mode
                self.log("Token 模式：尝试直接 HTTP 验证...")
                ews_url = f"https://{server or 'outlook.office365.com'}/ews/exchange.asmx" if not use_auto else "https://outlook.office365.com/ews/exchange.asmx"
                resp = requests.get(ews_url, headers={"Authorization": f"Bearer {token}"}, timeout=15)
                if resp.status_code in (200, 302, 401):
                    self.log(f"√ EWS 端点可达 (HTTP {resp.status_code})。如果 401，请检查 Token 是否有效。")
                else:
                    self.log(f"EWS 端点返回 HTTP {resp.status_code}", level="ERROR")
            else:
                if use_auto:
                    self.log(f"Using Autodiscover ({method})...")
                    account = Account(primary_smtp_address=test_email, credentials=creds, autodiscover=True)
                    if account.protocol.service_endpoint:
                        self.ews_server_var.set(account.protocol.service_endpoint)
                        self.log(f"Autodiscover found server: {account.protocol.service_endpoint}")
                else:
                    if not server:
                        raise Exception("Server URL required if Autodiscover is off.")
                    self.log(f"Connecting to server: {server} ({method})")
                    config_kwargs = {"server": server, "credentials": creds}
                    if ews_proto_auth_type:
                        config_kwargs["auth_type"] = ews_proto_auth_type
                    config = Configuration(**config_kwargs)
                    account = Account(primary_smtp_address=test_email, config=config, autodiscover=False)

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

        src_row1 = ttk.Frame(src_frame)
        src_row1.pack(fill="x", padx=5, pady=(5, 2))

        src_row2 = ttk.Frame(src_frame)
        src_row2.pack(fill="x", padx=5, pady=(2, 5))

        ttk.Label(src_row1, text="源系统:").pack(side="left", padx=(0, 5))
        ttk.Radiobutton(src_row1, text="Graph API", variable=self.source_type_var, value="Graph").pack(side="left", padx=5)
        ttk.Radiobutton(src_row1, text="Exchange EWS", variable=self.source_type_var, value="EWS").pack(side="left", padx=5)

        ttk.Label(src_row1, text="| 目标用户 CSV:").pack(side="left", padx=5)
        self.entry_csv_path = ttk.Entry(src_row1, textvariable=self.csv_path_var, width=50)
        self.entry_csv_path.pack(side="left", padx=5)
        self.btn_csv_browse = ttk.Button(
            src_row1,
            text="浏览...",
            command=lambda: self.csv_path_var.set(filedialog.askopenfilename(filetypes=[("CSV", "*.csv")]))
        )
        self.btn_csv_browse.pack(side="left")

        # TXT mailbox list import
        self.btn_txt_import = ttk.Button(
            src_row1,
            text="导入TXT邮箱列表...",
            command=self._import_mailbox_txt,
        )
        self.btn_txt_import.pack(side="left", padx=(5, 0))

        ttk.Label(src_row2, text="单个目标邮箱:").pack(side="left", padx=(0, 5))
        self.entry_single_target = ttk.Entry(src_row2, textvariable=self.target_single_email_var, width=40)
        self.entry_single_target.pack(side="left", padx=5)
        ttk.Label(src_row2, text="(填写后将忽略并禁用 CSV)").pack(side="left", padx=5)

        def _sync_target_input_state(*_args):
            has_single = bool((self.target_single_email_var.get() or '').strip())
            try:
                self.entry_csv_path.configure(state='disabled' if has_single else 'normal')
                self.btn_csv_browse.configure(state='disabled' if has_single else 'normal')
            except Exception:
                pass

        try:
            self.target_single_email_var.trace_add('write', _sync_target_input_state)
        except Exception:
            pass

        _sync_target_input_state()

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
            text="提示：会议不填写日期范围则不展开循环实例；填写开始+结束日期后，Graph/EWS 都会在该范围内展开循环会议 occurrence/exception。"
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
                if self.permanent_delete_var.get() and self.cleanup_target_var.get() == "Email":
                    messagebox.showwarning(
                        "警告",
                        "您已取消 '仅报告' 模式，并启用了【彻底删除(不可恢复)】！\n\nGraph 将尝试 permanentDelete（不进入 Recoverable Items）。\nEWS 将尽力使用更强删除类型（具体可恢复性取决于租户策略）。\n\n请务必确认 CSV 和 筛选条件 正确！",
                    )
                elif self.soft_delete_var.get() and self.cleanup_target_var.get() == "Email":
                    messagebox.showwarning(
                        "提示",
                        "您已取消 '仅报告' 模式，并启用了【软删除(移动到 Deleted Items)】。\n\nGraph 将优先使用 move -> deleteditems；EWS 将尽力使用 MoveToDeletedItems。\n\n请务必确认 CSV 和 筛选条件 正确！",
                    )
                else:
                    messagebox.showwarning(
                        "警告",
                        "您已取消 '仅报告' 模式！\n\n此删除通常属于可恢复删除（可能进入 Recoverable Items）。\n如需不可恢复删除，请勾选【彻底删除】（仅 Email 生效）。",
                    )

            try:
                if self.report_only_var.get():
                    self.chk_permanent_delete.configure(state="disabled")
                    try:
                        self.chk_soft_delete.configure(state="disabled")
                    except Exception:
                        pass
                else:
                    self.chk_permanent_delete.configure(state="normal" if self.cleanup_target_var.get() == "Email" else "disabled")
                    try:
                        self.chk_soft_delete.configure(state="normal" if self.cleanup_target_var.get() == "Email" else "disabled")
                    except Exception:
                        pass
            except Exception:
                pass

        def on_permanent_delete_change():
            # Keep checkbox meaningful
            if self.permanent_delete_var.get() and self.report_only_var.get():
                self.report_only_var.set(False)
                on_report_only_change()
            if self.permanent_delete_var.get():
                try:
                    self.soft_delete_var.set(False)
                except Exception:
                    pass

        def on_soft_delete_change():
            if self.soft_delete_var.get() and self.report_only_var.get():
                self.report_only_var.set(False)
                on_report_only_change()
            if self.soft_delete_var.get():
                try:
                    self.permanent_delete_var.set(False)
                except Exception:
                    pass

        ttk.Checkbutton(opt_frame, text="仅报告 (不删除)", variable=self.report_only_var, command=on_report_only_change).pack(side="left", padx=10)

        self.chk_soft_delete = ttk.Checkbutton(
            opt_frame,
            text="软删除(移动到 Deleted Items)",
            variable=self.soft_delete_var,
            command=on_soft_delete_change,
        )
        self.chk_soft_delete.pack(side="left", padx=5)

        self.chk_permanent_delete = ttk.Checkbutton(
            opt_frame,
            text="彻底删除(不可恢复)",
            variable=self.permanent_delete_var,
            command=on_permanent_delete_change,
        )
        self.chk_permanent_delete.pack(side="left", padx=5)
        try:
            self.chk_permanent_delete.configure(state="disabled" if self.report_only_var.get() else "normal")
            self.chk_soft_delete.configure(state="disabled" if self.report_only_var.get() else "normal")
        except Exception:
            pass

        ttk.Label(opt_frame, text="| 文件夹范围(Email):").pack(side="left", padx=5)
        self.mail_folder_scope_cb = ttk.Combobox(
            opt_frame,
            textvariable=self.mail_folder_scope_var,
            values=[
                "自动 (Auto)",
                "仅收件箱 (Inbox only)",
                "收件箱及子文件夹 (Inbox + Subfolders)",
                "常用文件夹 (Common: Inbox/Outbox/Sent/Deleted/Junk/Drafts/Archive)",
                "仅已发送 (Sent Items only)",
                "仅发件箱 (Outbox only)",
                "仅已删除 (Deleted Items only)",
                "仅可恢复删除 (Recoverable Items - Deletions)",
                "仅可恢复清除 (Recoverable Items - Purges)",
                "仅垃圾邮件 (Junk Email only)",
                "仅草稿 (Drafts only)",
                "仅存档 (Archive only)",
                "全邮箱 (All folders)",
            ],
            state="readonly",
            width=28,
        )
        try:
            if not (self.mail_folder_scope_var.get() or "").strip():
                self.mail_folder_scope_var.set("自动 (Auto)")
        except Exception:
            pass
        self.mail_folder_scope_cb.pack(side="left", padx=5)
        
        ttk.Label(opt_frame, text="| 日志级别:").pack(side="left", padx=5)
        
        def on_log_level_click():
            # Delegate to shared handler from Tools menu
            val = self.log_level_var.get()
            if val == "Expert":
                confirm = messagebox.askyesno("警告", "日志排错专用，日志量会很大且包含敏感信息，慎选！\n\n确认开启专家模式吗？")
                if not confirm:
                    self.log_level_var.set("Normal")
                    return
            try:
                self.logger.set_level(self.log_level_var.get())
            except Exception:
                pass

        ttk.Radiobutton(opt_frame, text="默认 (Default)", variable=self.log_level_var, value="Normal", command=on_log_level_click).pack(side="left", padx=5)
        ttk.Radiobutton(opt_frame, text="高级 (Advanced)", variable=self.log_level_var, value="Advanced", command=on_log_level_click).pack(side="left", padx=5)
        ttk.Radiobutton(opt_frame, text="专家 (Expert)", variable=self.log_level_var, value="Expert", command=on_log_level_click).pack(side="left", padx=5)

        # Start
        ttk.Button(frame, textvariable=self.btn_start_text, command=self.start_cleanup_thread).pack(pady=10, ipadx=20, ipady=5)

        # Progress bar
        progress_frame = ttk.Frame(frame)
        progress_frame.pack(fill="x", pady=(0, 5))

        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", length=600)
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self._progress_label_var = tk.StringVar(value="")
        self._progress_label = ttk.Label(progress_frame, textvariable=self._progress_label_var, width=30)
        self._progress_label.pack(side="right")

    # --- Tab 3: Scan Results ---
    def build_results_tab(self):
        frame = ttk.Frame(self.tab_results, padding=10)
        frame.pack(fill="both", expand=True)

        # --- Row 1: Info label ---
        info_frame = ttk.Frame(frame)
        info_frame.pack(fill="x", pady=(0, 3))
        self._results_info_var = tk.StringVar(value="尚无扫描结果。请先在「任务配置」中以【仅报告】模式运行扫描。")
        ttk.Label(info_frame, textvariable=self._results_info_var, wraplength=900).pack(side="left", fill="x", expand=True)

        # --- Row 2: Toolbar buttons ---
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(0, 5))
        ttk.Button(toolbar, text="刷新 / 加载报告", command=self._load_last_report).pack(side="left", padx=(0, 10))
        ttk.Button(toolbar, text="全选", width=8, command=self._select_all_results).pack(side="left", padx=2)
        ttk.Button(toolbar, text="取消全选", width=8, command=self._deselect_all_results).pack(side="left", padx=2)
        ttk.Button(toolbar, text="反选", width=8, command=self._invert_selection_results).pack(side="left", padx=2)

        # Separator
        ttk.Separator(toolbar, orient="vertical").pack(side="left", fill="y", padx=8)

        self._btn_delete_selected = ttk.Button(toolbar, text="删除选中项", command=self._delete_selected_results)
        self._btn_delete_selected.pack(side="left", padx=2)

        self._results_count_var = tk.StringVar(value="")
        ttk.Label(toolbar, textvariable=self._results_count_var, foreground="gray").pack(side="right")

        # --- Treeview with Scrollbar ---
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        # Create Treeview with style for row height
        style = ttk.Style()
        style.configure("Results.Treeview", rowheight=22)

        self.results_tree = ttk.Treeview(tree_frame, show="headings", selectmode="extended", style="Results.Treeview")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.results_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.results_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        # Placeholder columns — will be replaced when data loads
        self.results_tree["columns"] = ("☑", "提示")
        self.results_tree.heading("☑", text="☑", command=self._toggle_all_results)
        self.results_tree.column("☑", width=30, minwidth=30, stretch=False, anchor="center")
        self.results_tree.heading("提示", text="等待扫描结果...")
        self.results_tree.column("提示", width=600, stretch=True)

        # Click to toggle check
        self.results_tree.bind("<Button-1>", self._on_results_tree_click)
        # Double-click to toggle too
        self.results_tree.bind("<Double-1>", self._on_results_tree_click)

    # Column width presets based on data type
    _COL_WIDTH_MAP = {
        "☑": (30, False),
        "UserPrincipalName": (220, False),
        "Subject": (300, True),
        "Sender": (180, False),
        "Sender/Organizer": (180, False),
        "Organizer": (180, False),
        "Attendees": (220, True),
        "Received": (150, False),
        "Time": (150, False),
        "Start": (150, False),
        "End": (150, False),
        "Type": (70, False),
        "Details": (220, True),
        "UserRole": (75, False),
        "IsCancelled": (75, False),
        "ResponseStatus": (100, False),
        "RecurrencePattern": (110, False),
        "PatternDetails": (130, True),
        "RecurrenceDuration": (120, False),
        "IsEndless": (70, False),
    }

    # Columns to hide from Treeview display (still kept in CSV)
    _HIDDEN_COLS = {
        "Action", "Status", "MessageId", "ItemId",
        "MeetingGOID", "CleanGOID", "iCalUId", "SeriesMasterId",
    }

    def _populate_results_tree(self, columns: list[str], rows: list[dict]):
        """Fill the results Treeview with data."""
        self._scan_results_columns = columns
        self._scan_results_data = rows
        self._scan_checked.clear()

        # Clear old data
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)

        # Filter out hidden columns for display
        visible_cols = [c for c in columns if c not in self._HIDDEN_COLS]

        # Setup columns: ☑ + visible data columns
        display_cols = ["☑"] + visible_cols
        self.results_tree["columns"] = display_cols

        # ☑ column
        self.results_tree.heading("☑", text="☑", command=self._toggle_all_results)
        self.results_tree.column("☑", width=30, minwidth=30, stretch=False, anchor="center")

        # Data columns with smart widths
        for col in visible_cols:
            self.results_tree.heading(col, text=col, command=lambda c=col: self._sort_results_by(c))
            width, stretch = self._COL_WIDTH_MAP.get(col, (120, False))
            self.results_tree.column(col, width=width, minwidth=50, stretch=stretch)

        # Insert rows (only visible columns)
        for i, row in enumerate(rows):
            vals = ["☐"] + [str(row.get(c, "") or "") for c in visible_cols]
            iid = self.results_tree.insert("", "end", iid=str(i), values=vals)
            self._scan_checked[iid] = False

        count = len(rows)
        self._results_info_var.set(f"共 {count} 条结果。可勾选后点「删除选中项」进行删除。")
        self._results_count_var.set(f"已选: 0 / {count}")

    def _sort_results_by(self, col: str):
        """Sort treeview rows by a column (toggle asc/desc)."""
        try:
            items = list(self.results_tree.get_children())
            cols = list(self.results_tree["columns"])
            ci = cols.index(col)

            # Determine current sort direction
            reverse = getattr(self, '_sort_reverse', False)
            self._sort_reverse = not reverse

            items.sort(key=lambda iid: str(self.results_tree.item(iid, "values")[ci]).lower(), reverse=self._sort_reverse)
            for idx, iid in enumerate(items):
                self.results_tree.move(iid, "", idx)
        except Exception:
            pass

    def _on_results_tree_click(self, event):
        """Toggle checkbox when user clicks on the ☑ column."""
        region = self.results_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.results_tree.identify_column(event.x)
        if col != "#1":  # First column is ☑
            return
        iid = self.results_tree.identify_row(event.y)
        if not iid:
            return
        checked = self._scan_checked.get(iid, False)
        self._scan_checked[iid] = not checked
        vals = list(self.results_tree.item(iid, "values"))
        vals[0] = "☑" if not checked else "☐"
        self.results_tree.item(iid, values=vals)
        self._update_selection_count()
        return "break"

    def _update_selection_count(self):
        """Update the selected count label."""
        total = len(self._scan_checked)
        selected = sum(1 for v in self._scan_checked.values() if v)
        self._results_count_var.set(f"已选: {selected} / {total}")

    def _select_all_results(self):
        for iid in self.results_tree.get_children():
            self._scan_checked[iid] = True
            vals = list(self.results_tree.item(iid, "values"))
            vals[0] = "☑"
            self.results_tree.item(iid, values=vals)
        self._update_selection_count()

    def _deselect_all_results(self):
        for iid in self.results_tree.get_children():
            self._scan_checked[iid] = False
            vals = list(self.results_tree.item(iid, "values"))
            vals[0] = "☐"
            self.results_tree.item(iid, values=vals)
        self._update_selection_count()

    def _invert_selection_results(self):
        for iid in self.results_tree.get_children():
            checked = self._scan_checked.get(iid, False)
            self._scan_checked[iid] = not checked
            vals = list(self.results_tree.item(iid, "values"))
            vals[0] = "☑" if not checked else "☐"
            self.results_tree.item(iid, values=vals)
        self._update_selection_count()

    def _toggle_all_results(self):
        """Toggle all: if any unchecked, select all; otherwise deselect all."""
        any_unchecked = any(not v for v in self._scan_checked.values())
        if any_unchecked:
            self._select_all_results()
        else:
            self._deselect_all_results()

    def _load_last_report(self):
        """Load the most recent scan report CSV into the results Treeview."""
        path = self._last_report_path
        if not path or not os.path.exists(path):
            # Try to find latest report in reports dir
            try:
                report_files = sorted(
                    [f for f in os.listdir(self.reports_dir) if f.endswith('.csv')],
                    key=lambda f: os.path.getmtime(os.path.join(self.reports_dir, f)),
                    reverse=True,
                )
                if report_files:
                    path = os.path.join(self.reports_dir, report_files[0])
                else:
                    messagebox.showinfo("提示", "未找到任何报告文件。请先运行扫描。")
                    return
            except Exception as e:
                messagebox.showerror("错误", f"查找报告文件失败: {e}")
                return

        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                columns = reader.fieldnames or []
                rows = list(reader)
            self._last_report_path = path
            self._populate_results_tree(columns, rows)
            self.log(f">>> 已加载报告到扫描结果: {os.path.basename(path)} ({len(rows)} 条)")
            self.notebook.select(self.tab_results)
        except Exception as e:
            messagebox.showerror("错误", f"加载报告失败: {e}")

    def _delete_selected_results(self):
        """Delete the checked items via Graph or EWS."""
        selected_iids = [iid for iid, checked in self._scan_checked.items() if checked]
        if not selected_iids:
            messagebox.showinfo("提示", "未选中任何项目。请先勾选要删除的项目。")
            return

        count = len(selected_iids)
        confirm = messagebox.askyesno("确认删除", f"即将删除 {count} 个选中项目。\n\n此操作不可撤销。是否继续？")
        if not confirm:
            return

        source = self.source_type_var.get()
        self._btn_delete_selected.configure(state="disabled")
        threading.Thread(target=self._do_delete_selected, args=(selected_iids, source), daemon=True).start()

    def _do_delete_selected(self, selected_iids: list[str], source: str):
        """Background thread: delete selected items."""
        total = len(selected_iids)
        self.log(f">>> 开始删除 {total} 个选中项目 ({source})...")
        success = 0
        fail = 0

        try:
            if source == "Graph":
                success, fail = self._do_delete_graph(selected_iids)
            elif source == "EWS":
                success, fail = self._do_delete_ews(selected_iids)
        except Exception as e:
            self.log(f"删除过程出错: {e}", "ERROR")
        finally:
            self.root.after(0, lambda: self._btn_delete_selected.configure(state="normal"))

        self.log(f">>> 删除完成。成功: {success}, 失败: {fail}")
        self._update_selection_count()
        self.root.after(0, lambda s=success, f=fail: messagebox.showinfo("完成", f"删除完成。\n成功: {s}\n失败: {f}"))

    def _do_delete_graph(self, selected_iids: list[str]) -> tuple[int, int]:
        """Delete selected items via Graph API. Returns (success, fail)."""
        auth_mode = self.graph_auth_mode_var.get()
        tenant_id = self.tenant_id_var.get()
        app_id = self.app_id_var.get()
        thumbprint = self.thumbprint_var.get()
        client_secret = self.client_secret_var.get()
        env = self.graph_env_var.get()
        token = self._get_graph_access_token(auth_mode, tenant_id, app_id, thumbprint, client_secret, env)
        if not token:
            self.log("无法获取 Graph 访问令牌。", "ERROR")
            return 0, len(selected_iids)

        graph_endpoint = "https://microsoftgraph.chinacloudapi.cn" if env == "China" else "https://graph.microsoft.com"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

        success = 0
        fail = 0
        for i, iid in enumerate(selected_iids, 1):
            try:
                idx = int(iid)
                row = self._scan_results_data[idx]
                user = row.get("UserPrincipalName", "")
                item_id = row.get("ItemId", "") or row.get("MessageId", "")
                if not user or not item_id:
                    self.log(f"  跳过第 {i} 项: 缺少 UserPrincipalName 或 ItemId", "ERROR")
                    fail += 1
                    continue

                item_type = row.get("Type", "Email")
                resource = "events" if item_type not in ("Email", "") else "messages"

                del_url = f"{graph_endpoint}/v1.0/users/{user}/{resource}/{item_id}"
                resp = requests.delete(del_url, headers=headers, timeout=30)
                if resp.status_code in (200, 202, 204):
                    success += 1
                    self._update_result_row_status(iid, "Deleted", "success")
                else:
                    fail += 1
                    self._update_result_row_status(iid, f"Failed ({resp.status_code})", "error")
                    self.log(f"  删除失败 [{i}/{len(selected_iids)}]: {row.get('Subject', '?')} - HTTP {resp.status_code}", "ERROR")
            except Exception as e:
                fail += 1
                self.log(f"  删除异常 [{i}/{len(selected_iids)}]: {e}", "ERROR")

            if i % 10 == 0:
                self.log(f"  进度: {i}/{len(selected_iids)}  (成功: {success}, 失败: {fail})")

        return success, fail

    def _do_delete_ews(self, selected_iids: list[str]) -> tuple[int, int]:
        """Delete selected items via EWS. Returns (success, fail)."""
        creds, token = self._get_ews_credentials()
        if creds is None and token:
            self.log("EWS Token 模式下暂不支持从扫描结果删除，请使用 NTLM/Basic/OAuth2 模式。", "ERROR")
            return 0, len(selected_iids)

        server = self._clean_server_address(self.ews_server_var.get())
        use_auto = self.ews_use_autodiscover.get()
        auth_type = self.ews_auth_type_var.get()
        ews_auth_method = self.ews_auth_method_var.get()

        ews_proto_auth_type = None
        if ews_auth_method == "NTLM":
            ews_proto_auth_type = NTLM
        elif ews_auth_method == "Basic":
            ews_proto_auth_type = BASIC

        config = None
        if not use_auto and creds is not None:
            config_kwargs = {"server": server, "credentials": creds}
            if ews_proto_auth_type:
                config_kwargs["auth_type"] = ews_proto_auth_type
            config = Configuration(**config_kwargs)

        # Group by user for efficiency
        user_items: dict[str, list[tuple[str, dict]]] = {}
        for iid in selected_iids:
            idx = int(iid)
            row = self._scan_results_data[idx]
            user = row.get("UserPrincipalName", "")
            if user:
                user_items.setdefault(user, []).append((iid, row))

        access_type_val = IMPERSONATION if auth_type == "Impersonation" else DELEGATE
        success = 0
        fail = 0

        for target_email, items_list in user_items.items():
            try:
                if use_auto:
                    account = Account(primary_smtp_address=target_email, credentials=creds, autodiscover=True, access_type=access_type_val)
                else:
                    account = Account(primary_smtp_address=target_email, config=config, autodiscover=False, access_type=access_type_val)

                self.log(f"  已连接邮箱: {target_email} ({len(items_list)} 项待删除)")

                # Collect EWS item IDs for bulk delete
                batch_ids = []
                batch_iids = []

                for iid, row in items_list:
                    msg_id = row.get("MessageId", "") or row.get("ItemId", "")
                    if not msg_id:
                        fail += 1
                        self._update_result_row_status(iid, "No ID", "error")
                        continue
                    batch_ids.append(EwsItemId(id=msg_id))
                    batch_iids.append(iid)

                # Bulk delete in batches of 50
                batch_size = 50
                for bi in range(0, len(batch_ids), batch_size):
                    chunk_ids = batch_ids[bi:bi + batch_size]
                    chunk_iids = batch_iids[bi:bi + batch_size]
                    try:
                        account.bulk_delete(ids=chunk_ids, delete_type='MoveToDeletedItems')
                        for ciid in chunk_iids:
                            success += 1
                            self._update_result_row_status(ciid, "Deleted", "success")
                    except Exception as e:
                        # Fallback: try one by one
                        self.log(f"  批量删除失败，逐个重试: {e}", is_advanced=True)
                        for eid, ciid in zip(chunk_ids, chunk_iids):
                            try:
                                account.bulk_delete(ids=[eid], delete_type='MoveToDeletedItems')
                                success += 1
                                self._update_result_row_status(ciid, "Deleted", "success")
                            except Exception as ex2:
                                fail += 1
                                self._update_result_row_status(ciid, "Failed", "error")
                                self.log(f"  EWS 删除失败: {ex2}", "ERROR")

            except Exception as e:
                fail += len(items_list)
                self.log(f"  EWS 连接 {target_email} 失败: {e}", "ERROR")

        return success, fail

    def _update_result_row_status(self, iid: str, status_text: str, status_type: str):
        """Update the row in the results Treeview after a delete action."""
        def _do():
            try:
                vals = list(self.results_tree.item(iid, "values"))
                cols = list(self.results_tree["columns"])

                # Update Details column with status if visible
                if "Details" in cols:
                    di = cols.index("Details")
                    vals[di] = status_text

                if status_type == "success":
                    self._scan_checked[iid] = False
                    vals[0] = "☐"
                    self.results_tree.item(iid, values=vals, tags=("deleted",))
                elif status_type == "error":
                    self.results_tree.item(iid, values=vals, tags=("failed",))
                else:
                    self.results_tree.item(iid, values=vals)

                # Apply tag colors
                self.results_tree.tag_configure("deleted", foreground="gray")
                self.results_tree.tag_configure("failed", foreground="red")
            except Exception:
                pass
        self.root.after(0, _do)

    def update_ui_for_target(self):
        target = self.cleanup_target_var.get()
        if target == "Meeting":
            self.meeting_opt_frame.pack(fill="x", pady=5, before=self.filter_frame)

            try:
                self.permanent_delete_var.set(False)
                self.chk_permanent_delete.configure(state="disabled")
                self.soft_delete_var.set(False)
                self.chk_soft_delete.configure(state="disabled")
            except Exception:
                pass
            
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

            try:
                self.chk_permanent_delete.configure(state="disabled" if self.report_only_var.get() else "normal")
                self.chk_soft_delete.configure(state="disabled" if self.report_only_var.get() else "normal")
            except Exception:
                pass

    def start_cleanup_thread(self):
        if (not self.csv_path_var.get()) and (not (self.target_single_email_var.get() or '').strip()):
            messagebox.showerror("错误", "请选择 CSV 文件，或填写单个目标邮箱地址。")
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
            elif mode == "Manual":
                if not self.app_id_var.get() or not self.tenant_id_var.get() or not self.client_secret_var.get():
                    messagebox.showwarning("配置缺失", "您选择了 Graph API (手动/Secret) 模式，但未配置 App ID, Tenant ID 或 Client Secret。\n请前往 '1. 连接配置' 标签页进行配置。")
                    self.notebook.select(self.tab_connection)
                    return
            else:  # Token
                token_ui = (self.graph_token_var.get() or '').strip()
                if token_ui.lower().startswith('bearer '):
                    token_ui = token_ui[7:].strip()
                cached = getattr(self, '_graph_token_protected_cache', '') or ''
                if (not token_ui) and (not cached):
                    messagebox.showwarning("配置缺失", "您选择了 Graph API (直接输入 Token) 模式，但未填写 Token，且未检测到已缓存 Token。\n请前往 '1. 连接配置' 标签页填写 Token。")
                    self.notebook.select(self.tab_connection)
                    return

        elif source == "EWS":
            ews_method = self.ews_auth_method_var.get()
            if ews_method in ("Basic", "NTLM"):
                if not self.ews_user_var.get() or not self.ews_pass_var.get():
                    messagebox.showwarning("配置缺失", f"您选择了 EWS {ews_method} 模式，但未配置用户名或密码。\n请前往 '1. 连接配置' 标签页进行配置。")
                    self.notebook.select(self.tab_connection)
                    return
            elif ews_method == "OAuth2":
                if not self.ews_oauth_app_id_var.get() or not self.ews_oauth_tenant_id_var.get() or not self.ews_oauth_secret_var.get():
                    messagebox.showwarning("配置缺失", "您选择了 EWS OAuth2 模式，但未配置 App ID, Tenant ID 或 Client Secret。\n请前往 '1. 连接配置' 标签页进行配置。")
                    self.notebook.select(self.tab_connection)
                    return
            else:  # Token
                token_ui = (self.ews_token_var.get() or '').strip()
                if token_ui.lower().startswith('bearer '):
                    token_ui = token_ui[7:].strip()
                cached = getattr(self, '_ews_token_protected_cache', '') or ''
                if (not token_ui) and (not cached):
                    messagebox.showwarning("配置缺失", "您选择了 EWS Token 模式，但未填写 Token，且未检测到已缓存 Token。\n请前往 '1. 连接配置' 标签页填写 Token。")
                    self.notebook.select(self.tab_connection)
                    return
            if not self.ews_use_autodiscover.get() and not self.ews_server_var.get():
                messagebox.showwarning("配置缺失", "您选择了 EWS 模式且未启用自动发现，但未配置服务器地址。\n请前往 '1. 连接配置' 标签页进行配置。")
                self.notebook.select(self.tab_connection)
                return

        # Date Range Validation for Meetings
        # - Graph/EWS: 不填日期范围 => 不展开循环实例
        # - 填了开始+结束 => 允许展开，但限制跨度 <= 2 年
        if self.cleanup_target_var.get() == "Meeting":
            start_str = self._normalize_date_input(self.criteria_start_date.get())
            end_str = self._normalize_date_input(self.criteria_end_date.get())

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
            if self.cleanup_target_var.get() == "Email" and self.permanent_delete_var.get():
                msg = (
                    "您当前处于【删除模式】并启用了【彻底删除(不可恢复)】！\n\n"
                    "Graph 将尝试 permanentDelete（不进入 Recoverable Items）。\n"
                    "EWS 将尽力使用更强删除类型（具体可恢复性取决于租户策略）。\n\n"
                    "是否确认继续？"
                )
            elif self.cleanup_target_var.get() == "Email" and self.soft_delete_var.get():
                msg = (
                    "您当前处于【删除模式】并启用了【软删除(移动到 Deleted Items)】。\n\n"
                    "Graph 将优先使用 move -> deleteditems；EWS 将尽力使用 MoveToDeletedItems。\n\n"
                    "是否确认继续？"
                )
            else:
                msg = (
                    "您当前处于【删除模式】！\n\n"
                    "此删除通常属于可恢复删除（可能进入 Recoverable Items）。\n"
                    "如需不可恢复删除，请勾选【彻底删除】（仅 Email 生效）。\n\n"
                    "是否确认继续？"
                )

            confirm1 = messagebox.askyesno("高风险操作确认", msg)
            if not confirm1:
                return
            
            confirm2 = messagebox.askyesno("最终确认", "请再次确认：\n\n1. 您已备份重要数据。\n2. 您已确认 CSV 用户列表无误。\n3. 您已确认筛选条件无误。\n\n点击 '是' 将立即开始删除操作！")
            if not confirm2:
                return

        self.logger.set_level(self.log_level_var.get().upper())
        self.save_config()
        
        threading.Thread(target=self.run_cleanup, daemon=True).start()

    def run_cleanup(self):
        # Reset progress bar
        self._progress_reset(0)
        self.log("-" * 60)
        self.log(f"任务开始: {datetime.now()}")
        mode_str = '仅报告 (Report Only)' if self.report_only_var.get() else '删除 (DELETE)'
        if not self.report_only_var.get() and self.cleanup_target_var.get() == 'Email':
            if self.permanent_delete_var.get():
                mode_str += ' | 彻底删除'
            elif self.soft_delete_var.get():
                mode_str += ' | 软删除(Deleted Items)'
            else:
                mode_str += ' | 可恢复删除'
        self.log(f"模式: {mode_str}")
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

    def _get_target_users(self):
        single = (self.target_single_email_var.get() or '').strip()
        if single:
            if self.csv_path_var.get():
                self.log("检测到单个目标邮箱已填写，将忽略 CSV 列表。", is_advanced=True)
            return [single]

        csv_path = self.csv_path_var.get()
        if not csv_path:
            return []

        users = []
        # Support both CSV (with UserPrincipalName header) and plain text (one email per line)
        try:
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                first_line = f.readline().strip()
            if 'UserPrincipalName' in first_line:
                # CSV with header
                with open(csv_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        key = next((k for k in row.keys() if 'UserPrincipalName' in k), None)
                        if key and row.get(key):
                            users.append(row[key].strip())
            else:
                # Plain text: one email per line
                with open(csv_path, 'r', encoding='utf-8-sig') as f:
                    for line in f:
                        email = line.strip()
                        if email and '@' in email:
                            users.append(email)
        except Exception as e:
            self.log(f"读取用户列表文件失败: {e}", "ERROR")
        return users

    def _import_mailbox_txt(self):
        """Import a plain text file with one email per line, convert to CSV for use."""
        txt_path = filedialog.askopenfilename(
            title="选择邮箱列表文件",
            filetypes=[
                ("文本文件", "*.txt"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ]
        )
        if not txt_path:
            return

        try:
            with open(txt_path, 'r', encoding='utf-8-sig') as f:
                first_line = f.readline().strip()

            # If it's already a CSV with UserPrincipalName header, just set path
            if 'UserPrincipalName' in first_line:
                self.csv_path_var.set(txt_path)
                self.log(f"已加载 CSV 邮箱列表: {txt_path}")
                return

            # Read emails from txt (one per line)
            emails = []
            with open(txt_path, 'r', encoding='utf-8-sig') as f:
                for line in f:
                    email = line.strip()
                    if email and '@' in email:
                        emails.append(email)

            if not emails:
                messagebox.showwarning("提示", "未在文件中找到有效的邮箱地址。\n\n格式要求：每行一个邮箱地址。")
                return

            # Convert to CSV in reports dir
            csv_out = os.path.join(self.reports_dir, f"imported_mailbox_list_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            with open(csv_out, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(['UserPrincipalName'])
                for email in emails:
                    writer.writerow([email])

            self.csv_path_var.set(csv_out)
            self.log(f"已导入 {len(emails)} 个邮箱地址 (从 {os.path.basename(txt_path)} 转换为 CSV)")
            messagebox.showinfo("导入成功", f"成功导入 {len(emails)} 个邮箱地址。")

        except Exception as e:
            messagebox.showerror("导入失败", f"读取文件失败: {e}")

    # --- Progress Bar Helpers ---
    def _progress_reset(self, total: int):
        """Reset the progress bar for a new task."""
        self._progress_total = total
        self._progress_done = 0
        def _do():
            self.progress_bar["maximum"] = max(total, 1)
            self.progress_bar["value"] = 0
            self._progress_label_var.set(f"0 / {total} (0%)")
        self.root.after(0, _do)

    def _progress_increment(self, label: str = ""):
        """Increment progress by one step."""
        self._progress_done += 1
        done = self._progress_done
        total = self._progress_total
        pct = int(done / max(total, 1) * 100)
        def _do():
            self.progress_bar["value"] = done
            text = f"{done} / {total} ({pct}%)"
            if label:
                text += f"  {label}"
            self._progress_label_var.set(text)
        self.root.after(0, _do)

    def _progress_finish(self, text: str = "完成"):
        """Mark progress as complete."""
        def _do():
            self.progress_bar["value"] = self.progress_bar["maximum"]
            self._progress_label_var.set(text)
        self.root.after(0, _do)

    def _get_graph_access_token(self, auth_mode, tenant_id, app_id, thumbprint, client_secret, env):
        if auth_mode == "Token":
            token_ui = (self.graph_token_var.get() or '').strip()
            if token_ui.lower().startswith('bearer '):
                token_ui = token_ui[7:].strip()

            token = token_ui
            from_cache = False
            if not token:
                protected = getattr(self, '_graph_token_protected_cache', '') or ''
                if protected:
                    token = _dpapi_unprotect_text(protected) or ''
                    from_cache = True

            if not token:
                raise Exception("未提供 Token，且无法从缓存读取 Token")

            # Refresh cache if enabled and token came from UI
            try:
                cache_enabled = bool(self.graph_cache_token_var.get())
            except Exception:
                cache_enabled = True

            if cache_enabled and (not from_cache) and token:
                protected_new = _dpapi_protect_text(token)
                if protected_new:
                    self._graph_token_protected_cache = protected_new
                    try:
                        # Persist best-effort
                        self.save_config()
                    except Exception:
                        pass

            return token

        if auth_mode == "Auto":
            return self.get_token_from_cert(tenant_id, app_id, thumbprint, env)
        return self.get_token_from_secret(tenant_id, app_id, client_secret, env)

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
                                  report_only, writer, csv_lock, calendar_view_start=None, calendar_view_end=None, mail_folder_scope: str | None = None, permanent_delete: bool = False, soft_delete: bool = False):
        self.log(f"--- 正在处理: {user} ---")
        try:
            req_headers = dict(headers)
            session = _get_pooled_session()

            def _graph_request(method: str, url: str, *, params: dict | None = None, json_body=None):
                # Fast path: handle throttling/transient errors with limited retries.
                # We keep this conservative to avoid making rate-limit worse.
                max_attempts = 6
                base_sleep = 0.6
                for attempt in range(1, max_attempts + 1):
                    resp = session.request(method, url, headers=req_headers, params=params, json=json_body)

                    if resp.status_code in (429, 503, 502, 504):
                        retry_after = resp.headers.get('Retry-After')
                        if retry_after:
                            try:
                                sleep_s = float(retry_after)
                            except Exception:
                                sleep_s = base_sleep * (2 ** (attempt - 1))
                        else:
                            sleep_s = base_sleep * (2 ** (attempt - 1))
                        # jitter to spread concurrent threads
                        sleep_s = min(12.0, sleep_s) + random.random() * 0.25
                        if attempt < max_attempts:
                            self.log(f"Graph 请求被限流/暂时失败({resp.status_code})，等待 {sleep_s:.2f}s 后重试...", is_advanced=True)
                            time.sleep(sleep_s)
                            continue
                    return resp
                return resp

            def _graph_batch_send(batch_requests: list[dict]) -> dict | None:
                # Graph $batch supports up to 20 requests
                if not batch_requests:
                    return None
                batch_url = f"{graph_endpoint}/v1.0/$batch"
                resp = _graph_request("POST", batch_url, json_body={"requests": batch_requests})
                if resp.status_code != 200:
                    return None
                try:
                    return resp.json()
                except Exception:
                    return None

            def _to_batch_rel(full_url: str) -> str:
                # Graph $batch expects relative urls starting with '/'
                if full_url.startswith(graph_endpoint):
                    rel = full_url[len(graph_endpoint):]
                    return rel if rel.startswith('/') else '/' + rel
                return full_url

            def _infer_email_folder_mode() -> str:
                s = (mail_folder_scope or "自动 (Auto)").strip().lower()
                if "common" in s or "常用" in s:
                    return "common"
                if "sub" in s or "子文件夹" in s:
                    return "inbox_subtree"
                if "sent" in s or "已发送" in s:
                    return "sent_only"
                if "outbox" in s or "发件箱" in s:
                    return "outbox_only"
                if "recoverable" in s or "可恢复" in s:
                    if "purge" in s or "清除" in s:
                        return "recoverable_items_purges"
                    return "recoverable_items_deletions"
                if "deleted" in s or "已删除" in s:
                    return "deleted_only"
                if "junk" in s or "垃圾" in s:
                    return "junk_only"
                if "draft" in s or "草稿" in s:
                    return "drafts_only"
                if "archive" in s or "存档" in s:
                    return "archive_only"
                if "inbox" in s or "收件箱" in s:
                    return "inbox_only"
                if "all" in s or "全邮箱" in s:
                    return "all"
                return "auto"

            def _graph_get_json(url: str, *, params: dict | None = None) -> dict:
                graph_log_level = self.log_level_var.get()
                if graph_log_level in ("Advanced", "Expert"):
                    save_auth = bool(graph_log_level == "Expert" and getattr(self, 'graph_save_auth_token_var', None) and self.graph_save_auth_token_var.get())
                    self.logger.log_to_file_only(f"GRAPH REQ: GET {url}")
                    self.logger.log_to_file_only(f"HEADERS: {json.dumps(redact_sensitive_headers(req_headers, save_authorization=save_auth), default=str)}")
                    if params:
                        self.logger.log_to_file_only(f"PARAMS: {json.dumps(params, default=str)}")
                resp = _graph_request("GET", url, params=params)
                if resp.status_code != 200:
                    raise Exception(f"Graph folder query failed: {resp.status_code} {resp.text}")
                return resp.json()

            def _paged_folder_values(url: str, *, params: dict | None = None) -> list[dict]:
                out: list[dict] = []
                next_url = url
                next_params = params
                while next_url:
                    data = _graph_get_json(next_url, params=next_params)
                    out.extend(data.get('value', []) or [])
                    next_url = data.get('@odata.nextLink')
                    next_params = None
                return out

            email_folder_mode = _infer_email_folder_mode()
            if target_type == "Email" and email_folder_mode == "auto":
                # Preserve existing Graph behavior: mailbox-wide /messages
                email_folder_mode = "all"

            # Meetings: prefer calendarView to expand recurrence into occurrence/exception within a date range
            if target_type == "Meeting" and resource == "calendarView":
                if not calendar_view_start or not calendar_view_end:
                    raise Exception("Graph calendarView requires startDateTime and endDateTime")
                url = f"{graph_endpoint}/v1.0/users/{user}/calendarView"
                params = {
                    "startDateTime": calendar_view_start,
                    "endDateTime": calendar_view_end,
                    "$top": 100,
                    "$select": "id,subject,organizer,attendees,start,end,type,isCancelled,iCalUId,seriesMasterId,responseStatus,recurrence,bodyPreview",
                    # Try to read GOID via MAPI extended property PidLidGlobalObjectId (PSETID_Meeting, Id 0x0003)
                    "$expand": "singleValueExtendedProperties($filter=id eq 'Binary {6ED8DA90-450B-101B-98DA-00AA003F1305} Id 0x0003')",
                }
            else:
                if target_type != "Email":
                    url = f"{graph_endpoint}/v1.0/users/{user}/{resource}"
                    params = {
                        "$top": 100,
                        "$select": "id,subject,organizer,attendees,start,end,type,isCancelled,iCalUId,seriesMasterId,responseStatus,recurrence,bodyPreview",
                        "$expand": "singleValueExtendedProperties($filter=id eq 'Binary {6ED8DA90-450B-101B-98DA-00AA003F1305} Id 0x0003')",
                    }
                else:
                    # Email: optionally restrict by folder scope
                    base_resources: list[str]
                    if email_folder_mode == "common":
                        base_resources = [
                            "mailFolders/inbox/messages",
                            "mailFolders/outbox/messages",
                            "mailFolders/sentitems/messages",
                            "mailFolders/deleteditems/messages",
                            "mailFolders/junkemail/messages",
                            "mailFolders/drafts/messages",
                            "mailFolders/archive/messages",
                        ]
                    elif email_folder_mode == "sent_only":
                        base_resources = ["mailFolders/sentitems/messages"]
                    elif email_folder_mode == "outbox_only":
                        base_resources = ["mailFolders/outbox/messages"]
                    elif email_folder_mode == "deleted_only":
                        base_resources = ["mailFolders/deleteditems/messages"]
                    elif email_folder_mode == "recoverable_items_deletions":
                        base_resources = ["mailFolders/recoverableitemsdeletions/messages"]
                    elif email_folder_mode == "recoverable_items_purges":
                        base_resources = ["mailFolders/recoverableitemspurges/messages"]
                    elif email_folder_mode == "junk_only":
                        base_resources = ["mailFolders/junkemail/messages"]
                    elif email_folder_mode == "drafts_only":
                        base_resources = ["mailFolders/drafts/messages"]
                    elif email_folder_mode == "archive_only":
                        base_resources = ["mailFolders/archive/messages"]
                    elif email_folder_mode in ("inbox_only", "inbox_subtree"):
                        if email_folder_mode == "inbox_only":
                            base_resources = ["mailFolders/inbox/messages"]
                        else:
                            try:
                                inbox_meta = _graph_get_json(
                                    f"{graph_endpoint}/v1.0/users/{user}/mailFolders/inbox",
                                    params={"$select": "id,childFolderCount"},
                                )
                                inbox_id = inbox_meta.get('id')
                                if not inbox_id:
                                    raise Exception("Inbox id missing")
                                folder_ids: list[str] = []
                                q: list[str] = [inbox_id]
                                while q:
                                    fid = q.pop(0)
                                    folder_ids.append(fid)
                                    kids = _paged_folder_values(
                                        f"{graph_endpoint}/v1.0/users/{user}/mailFolders/{fid}/childFolders",
                                        params={"$top": 200, "$select": "id,childFolderCount"},
                                    )
                                    for k in kids:
                                        kid = k.get('id')
                                        if kid:
                                            q.append(kid)
                                base_resources = [f"mailFolders/{fid}/messages" for fid in folder_ids]
                            except Exception as e:
                                self.log(f"警告: 无法枚举 Inbox 子文件夹，回退为仅 Inbox。原因: {e}", level="ERROR")
                                base_resources = ["mailFolders/inbox/messages"]
                    else:
                        base_resources = [resource]

                    # Email listing can be chatty; reduce payload when body filter is not used.
                    select_fields = "id,subject,from,receivedDateTime"
                    if body_keyword:
                        select_fields = "id,subject,from,receivedDateTime,body"
                    params = {"$top": 500, "$select": select_fields}

                    # Iterate each base resource separately (folder scope)
                    for _res in base_resources:
                        url = f"{graph_endpoint}/v1.0/users/{user}/{_res}"

                        params2 = dict(params or {})
                        if filter_str:
                            params2["$filter"] = filter_str
                        if body_keyword:
                            params2["$search"] = f'"body:{body_keyword}"'
                            req_headers["ConsistencyLevel"] = "eventual"

                        next_url = url
                        local_params = params2
                        while next_url:
                            graph_log_level = self.log_level_var.get()
                            if graph_log_level in ("Advanced", "Expert"):
                                save_auth = bool(graph_log_level == "Expert" and getattr(self, 'graph_save_auth_token_var', None) and self.graph_save_auth_token_var.get())
                                self.logger.log_to_file_only(f"GRAPH REQ: GET {next_url}")
                                self.logger.log_to_file_only(f"HEADERS: {json.dumps(redact_sensitive_headers(req_headers, save_authorization=save_auth), default=str)}")
                                if local_params:
                                    self.logger.log_to_file_only(f"PARAMS: {json.dumps(local_params, default=str)}")

                            self.log(f"请求: GET {next_url} | 参数: {local_params}", is_advanced=True)
                            resp = _graph_request("GET", next_url, params=local_params if "users" in next_url and "?" not in next_url else None)

                            if graph_log_level in ("Advanced", "Expert"):
                                self.logger.log_to_file_only(f"GRAPH RESP: {resp.status_code}")
                                self.logger.log_to_file_only(f"HEADERS: {json.dumps(dict(resp.headers), default=str)}")
                                body_text = resp.text or ""
                                if graph_log_level == "Advanced":
                                    body_text = body_text[:4096]
                                else:
                                    body_text = body_text[:50000]
                                self.logger.log_to_file_only(f"BODY: {body_text}")

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

                            perm_enabled = bool(permanent_delete and target_type == "Email" and str(delete_resource).lower() == "messages")
                            soft_enabled = bool((not perm_enabled) and (target_type == "Email") and bool(soft_delete) and str(delete_resource).lower() == "messages")

                            delete_candidates: list[tuple[dict, str, str]] = []  # (row_data, item_id, del_url)
                            for item in items:
                                should_delete = True
                                if body_keyword and "$search" not in (local_params or {}):
                                    content = item.get('body', {}).get('content', '')
                                    if body_keyword.lower() not in content.lower():
                                        should_delete = False

                                if should_delete:
                                    item_id = item['id']
                                    subject = item.get('subject', '无主题')
                                    sender = item.get('from', {}).get('emailAddress', {}).get('address', '未知')
                                    time_val = item.get('receivedDateTime')
                                    item_type = "Email"

                                    row_data = {
                                        'UserPrincipalName': user,
                                        'ItemId': item_id,
                                        'Subject': subject,
                                        'Sender/Organizer': sender,
                                        'Time': time_val,
                                        'Type': item_type,
                                        'Action': 'ReportOnly' if report_only else ('PermanentDelete' if perm_enabled else ('SoftDelete' if soft_enabled else 'Delete')),
                                        'Status': 'Pending',
                                        'Details': ''
                                    }

                                    if report_only:
                                        self.log(f"  [报告] 发现: {subject} ({item_type})")
                                        row_data['Status'] = 'Skipped'
                                        row_data['Details'] = '仅报告模式'
                                    if report_only:
                                        with csv_lock:
                                            writer.writerow(row_data)
                                    else:
                                        del_url = f"{graph_endpoint}/v1.0/users/{user}/{delete_resource}/{item_id}"
                                        delete_candidates.append((row_data, item_id, del_url))

                            # If we are deleting, use Graph $batch (20 req per call)
                            if (not report_only) and delete_candidates:
                                # Chunk into batches of 20
                                for i in range(0, len(delete_candidates), 20):
                                    chunk = delete_candidates[i:i+20]
                                    batch_requests = []
                                    id_to_row = {}
                                    id_to_delurl = {}
                                    id_to_mode = {}
                                    for j, (row_data, _item_id, del_url) in enumerate(chunk, start=1):
                                        req_id = str(j)
                                        id_to_row[req_id] = row_data
                                        id_to_delurl[req_id] = del_url

                                        if perm_enabled:
                                            method = "POST"
                                            url_rel = _to_batch_rel(f"{del_url}/permanentDelete")
                                            id_to_mode[req_id] = "perm"
                                        elif soft_enabled:
                                            method = "POST"
                                            url_rel = _to_batch_rel(f"{del_url}/move")
                                            id_to_mode[req_id] = "soft"
                                        else:
                                            method = "DELETE"
                                            url_rel = _to_batch_rel(del_url)
                                            id_to_mode[req_id] = "delete"

                                        req = {
                                            "id": req_id,
                                            "method": method,
                                            "url": url_rel,
                                            "headers": {"Content-Type": "application/json"},
                                        }
                                        if soft_enabled:
                                            req["body"] = {"destinationId": "deleteditems"}
                                        batch_requests.append(req)

                                    batch_json = _graph_batch_send(batch_requests)
                                    resp_map = {}
                                    if batch_json and isinstance(batch_json, dict):
                                        for r in (batch_json.get('responses') or []):
                                            if isinstance(r, dict) and 'id' in r:
                                                resp_map[str(r.get('id'))] = r

                                    for req_id, row_data in id_to_row.items():
                                        r = resp_map.get(req_id)
                                        status = None
                                        if r is not None:
                                            status = r.get('status')

                                        # If batch failed entirely, fall back to single-request
                                        if status is None:
                                            del_url = id_to_delurl.get(req_id)
                                            mode = id_to_mode.get(req_id) or "delete"
                                            self.log(f"  正在删除(回退): {row_data.get('Subject', '')}")
                                            if mode == "perm":
                                                del_resp = _graph_request("POST", f"{del_url}/permanentDelete")
                                                ok_codes = (200, 201, 202, 204)
                                            elif mode == "soft":
                                                del_resp = _graph_request("POST", f"{del_url}/move", json_body={"destinationId": "deleteditems"})
                                                ok_codes = (200, 201, 202, 204)
                                            else:
                                                del_resp = _graph_request("DELETE", del_url)
                                                ok_codes = (202, 204)

                                            if del_resp.status_code in ok_codes:
                                                row_data['Status'] = 'Success'
                                            else:
                                                row_data['Status'] = 'Failed'
                                                row_data['Details'] = f"状态码: {del_resp.status_code}"
                                            with csv_lock:
                                                writer.writerow(row_data)
                                            continue

                                        # permanentDelete may be unsupported; fall back
                                        if perm_enabled and status in (404, 405):
                                            del_url = id_to_delurl.get(req_id)
                                            self.log("    ! permanentDelete 不可用，回退普通删除（可能进入 Recoverable Items）。", level="ERROR")
                                            del_resp = _graph_request("DELETE", del_url)
                                            if del_resp.status_code in (204, 202):
                                                row_data['Status'] = 'Success'
                                            else:
                                                row_data['Status'] = 'Failed'
                                                row_data['Details'] = f"状态码: {del_resp.status_code}"
                                            with csv_lock:
                                                writer.writerow(row_data)
                                            continue

                                        # move may be unsupported; fall back
                                        if soft_enabled and status in (404, 405):
                                            del_url = id_to_delurl.get(req_id)
                                            self.log("    ! move 不可用，回退普通删除（可能进入 Recoverable Items）。", level="ERROR")
                                            del_resp = _graph_request("DELETE", del_url)
                                            if del_resp.status_code in (204, 202):
                                                row_data['Status'] = 'Success'
                                                row_data['Details'] = 'move 不可用，已回退 DELETE'
                                            else:
                                                row_data['Status'] = 'Failed'
                                                row_data['Details'] = f"状态码: {del_resp.status_code}"
                                            with csv_lock:
                                                writer.writerow(row_data)
                                            continue

                                        if soft_enabled and status == 400:
                                            try:
                                                body = r.get('body') if isinstance(r, dict) else None
                                                msg = ((body or {}).get('error') or {}).get('message') if isinstance(body, dict) else ''
                                                msg = (msg or '').lower()
                                                if 'destination' in msg and ('same' in msg or 'identical' in msg):
                                                    row_data['Status'] = 'Success'
                                                    row_data['Details'] = '已在 Deleted Items，无需移动'
                                                    with csv_lock:
                                                        writer.writerow(row_data)
                                                    continue
                                            except Exception:
                                                pass

                                        if status in (204, 202, 200, 201):
                                            row_data['Status'] = 'Success'
                                        else:
                                            row_data['Status'] = 'Failed'
                                            row_data['Details'] = f"状态码: {status}"
                                        with csv_lock:
                                            writer.writerow(row_data)

                            next_url = data.get('@odata.nextLink')
                            local_params = None

                    # Email handled above; return to avoid running legacy single-resource path
                    return

            if filter_str: params["$filter"] = filter_str
            
            if body_keyword:
                params["$search"] = f'"body:{body_keyword}"'
                req_headers["ConsistencyLevel"] = "eventual"
            
            while url:
                graph_log_level = self.log_level_var.get()
                if graph_log_level in ("Advanced", "Expert"):
                    save_auth = bool(graph_log_level == "Expert" and getattr(self, 'graph_save_auth_token_var', None) and self.graph_save_auth_token_var.get())
                    self.logger.log_to_file_only(f"GRAPH REQ: GET {url}")
                    self.logger.log_to_file_only(f"HEADERS: {json.dumps(redact_sensitive_headers(req_headers, save_authorization=save_auth), default=str)}")
                    if params:
                        self.logger.log_to_file_only(f"PARAMS: {json.dumps(params, default=str)}")

                self.log(f"请求: GET {url} | 参数: {params}", is_advanced=True)
                resp = _graph_request("GET", url, params=params if "users" in url and "?" not in url else None) # Simple check to avoid double params
                
                if graph_log_level in ("Advanced", "Expert"):
                    self.logger.log_to_file_only(f"GRAPH RESP: {resp.status_code}")
                    self.logger.log_to_file_only(f"HEADERS: {json.dumps(dict(resp.headers), default=str)}")
                    body_text = resp.text or ""
                    if graph_log_level == "Advanced":
                        body_text = body_text[:4096]
                    else:
                        body_text = body_text[:50000]
                    self.logger.log_to_file_only(f"BODY: {body_text}")
                
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

                            goid_hex = decode_graph_goid_base64_to_hex(goid_b64)

                            user_role = 'Attendee'
                            try:
                                if sender and user and sender.strip().lower() == user.strip().lower():
                                    user_role = 'Organizer'
                            except Exception:
                                pass

                            response_status = format_graph_meeting_response_status(
                                user_email=user,
                                user_role=user_role,
                                organizer_email=sender,
                                attendees=attendees,
                                item_response_status=(item.get('responseStatus') or {}),
                            )

                            # Align with EWS: MeetingGOID uses iCalUId (same semantic as item.uid in EWS)
                            meeting_goid = ical_uid or ''
                            clean_goid = (meeting_goid or item_id).strip().lower()

                            details_hint = ''
                            if goid_b64 or goid_hex:
                                details_hint = f"GOID(b64)={goid_b64}; GOID(hex)={goid_hex}".strip('; ')

                            row_data = {
                                'UserPrincipalName': user,
                                'Subject': subject,
                                'Type': item_type,
                                'MeetingGOID': meeting_goid,
                                'CleanGOID': clean_goid,
                                'iCalUId': ical_uid,
                                'SeriesMasterId': series_master_id,
                                'Organizer': sender,
                                'Attendees': ';'.join(attendee_emails),
                                'Start': start_val,
                                'End': end_val,
                                'UserRole': user_role,
                                'IsCancelled': is_cancelled,
                                'ResponseStatus': response_status,
                                'RecurrencePattern': '',
                                'PatternDetails': '',
                                'RecurrenceDuration': '',
                                'IsEndless': '',
                                'Action': 'ReportOnly' if report_only else 'Delete',
                                'Status': 'Pending',
                                'Details': details_hint
                            }

                            if is_cancelled:
                                row_data['Type'] = f"{row_data['Type']} (Cancelled)"

                            # If current item already has recurrence (usually seriesMaster), format it.
                            try:
                                if item.get('recurrence'):
                                    rec = item.get('recurrence') or {}
                                    p_name, p_details = format_graph_recurrence_pattern(rec.get('pattern') or {})
                                    dur, endless = format_graph_recurrence_range(rec.get('range') or {})
                                    row_data['RecurrencePattern'] = p_name
                                    row_data['PatternDetails'] = p_details
                                    row_data['RecurrenceDuration'] = dur
                                    row_data['IsEndless'] = endless
                            except Exception:
                                pass

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
                                        # Use pooled session + retry
                                        m_resp = _graph_request("GET", master_url, params=master_params)
                                        if m_resp.status_code == 200:
                                            user_cache[series_master_id] = m_resp.json()
                                        else:
                                            user_cache[series_master_id] = None
                                    master_obj = user_cache.get(series_master_id)
                                    if master_obj and master_obj.get('recurrence'):
                                        rec = master_obj.get('recurrence')
                                        pattern = (rec.get('pattern') or {})
                                        rng = (rec.get('range') or {})
                                        p_name, p_details = format_graph_recurrence_pattern(pattern)
                                        dur, endless = format_graph_recurrence_range(rng)
                                        row_data['RecurrencePattern'] = p_name
                                        row_data['PatternDetails'] = p_details
                                        row_data['RecurrenceDuration'] = dur
                                        row_data['IsEndless'] = endless
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
                            row_data['Details'] = ((row_data.get('Details') + '; ') if row_data.get('Details') else '') + '仅报告模式'
                        else:
                            self.log(f"  正在删除: {subject}")
                            del_url = f"{graph_endpoint}/v1.0/users/{user}/{delete_resource}/{item_id}"
                            
                            if graph_log_level in ("Advanced", "Expert"):
                                save_auth = bool(graph_log_level == "Expert" and getattr(self, 'graph_save_auth_token_var', None) and self.graph_save_auth_token_var.get())
                                self.logger.log_to_file_only(f"GRAPH REQ: DELETE {del_url}")
                                self.logger.log_to_file_only(f"HEADERS: {json.dumps(redact_sensitive_headers(req_headers, save_authorization=save_auth), default=str)}")

                            self.log(f"请求: DELETE {del_url}", is_advanced=True)
                            del_resp = _graph_request("DELETE", del_url)
                            
                            if graph_log_level in ("Advanced", "Expert"):
                                self.logger.log_to_file_only(f"GRAPH RESP: {del_resp.status_code}")
                                body_text = del_resp.text or ""
                                if graph_log_level == "Advanced":
                                    body_text = body_text[:4096]
                                else:
                                    body_text = body_text[:50000]
                                self.logger.log_to_file_only(f"BODY: {body_text}")
                            
                            if del_resp.status_code == 204:
                                self.log("    √ 已删除")
                                row_data['Status'] = 'Success'
                            else:
                                self.log(f"    X 删除失败: {del_resp.status_code}", "ERROR")
                                self.log(f"响应: {del_resp.text}", is_advanced=True)
                                row_data['Status'] = 'Failed'
                                err_detail = f"状态码: {del_resp.status_code}"
                                row_data['Details'] = ((row_data.get('Details') + '; ') if row_data.get('Details') else '') + err_detail
                        
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
            token = self._get_graph_access_token(auth_mode, tenant_id, app_id, thumbprint, client_secret, env)
                
            if not token: raise Exception("获取 Token 失败")
            
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            self.log("√ Token 获取成功")

            users = self._get_target_users()
            
            self.log(f">>> 找到 {len(users)} 个用户")
            self._progress_reset(len(users))

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
                    # If both start+end present -> use calendarView to expand recurrence instances
                    # Otherwise -> fallback to events (no recurrence expansion)
                    if start_date and end_date:
                        resource = "calendarView"
                        delete_resource = "events"
                    else:
                        resource = "events"
                        delete_resource = "events"

                    if self.criteria_sender.get():
                        filters.append(f"organizer/emailAddress/address eq '{self.criteria_sender.get()}'")

                    if resource == "events":
                        if start_date:
                            filters.append(f"start/dateTime ge '{start_date}T00:00:00'")
                        if end_date:
                            filters.append(f"end/dateTime le '{end_date}T23:59:59'")
                    
                    # Meeting Specifics
                    if self.meeting_only_cancelled_var.get():
                        filters.append("isCancelled eq true")
                    
                    scope = self.meeting_scope_var.get()
                    if "Single" in scope:
                        filters.append("type eq 'singleInstance'")
                    elif "Series" in scope:
                        if resource == "calendarView":
                            # Include expanded instances too (occurrence/exception) for EWS-like scan
                            filters.append("type eq 'seriesMaster' or type eq 'occurrence' or type eq 'exception'")
                        else:
                            filters.append("type eq 'seriesMaster'")
                    # If All, no type filter

                filter_str = " and ".join(filters)
                body_keyword = self.criteria_body.get()

                calendar_view_start = None
                calendar_view_end = None
                if target_type == "Meeting" and resource == "calendarView":
                    # Use UTC ISO; calendarView requires both
                    calendar_view_start = f"{start_date}T00:00:00Z"
                    calendar_view_end = f"{end_date}T23:59:59Z"

                csv_lock = threading.Lock()
                report_only = self.report_only_var.get()
                permanent_delete = bool(self.permanent_delete_var.get()) and (not report_only) and (target_type == "Email")
                
                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = []
                    mail_folder_scope = self.mail_folder_scope_var.get()
                    soft_delete = bool(self.soft_delete_var.get()) and (not report_only) and (target_type == "Email") and (not permanent_delete)
                    for user in users:
                        futures.append(executor.submit(
                            self.process_single_user_graph, 
                            user, graph_endpoint, headers, resource, delete_resource, target_type, filter_str, body_keyword,
                            report_only, writer, csv_lock, calendar_view_start, calendar_view_end, mail_folder_scope, permanent_delete, soft_delete
                        ))
                    
                    # Wait for all to complete
                    for future in futures:
                        try:
                            future.result()
                        except Exception as e:
                            self.log(f"Task Error: {e}", "ERROR")
                        self._progress_increment()

            self._progress_finish("Graph 任务完成")
            self.log(f">>> 任务完成! 报告: {report_path}")
            msg_title = "完成"
            if self.report_only_var.get():
                msg_body = f"扫描生成报告任务完成。\n报告: {report_path}"
                # Auto-load results into tab 3
                self.root.after(100, self._load_last_report)
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
                                report_only, writer, csv_lock, log_level, mail_folder_scope: str | None = None, permanent_delete: bool = False, soft_delete: bool = False,
                                access_token: str | None = None):
        try:
            self.log(f"--- 正在处理: {target_email} ---")
            
            # Build Account — support Basic credentials or OAuth2/Token
            access_type_val = IMPERSONATION if auth_type == "Impersonation" else DELEGATE

            if creds is not None:
                # Basic or OAuth2Credentials path
                if use_auto:
                    account = Account(primary_smtp_address=target_email, credentials=creds, autodiscover=True, access_type=access_type_val)
                else:
                    account = Account(primary_smtp_address=target_email, config=config, autodiscover=False, access_type=access_type_val)
            elif access_token:
                # Token-only mode: inject bearer token via OAuth2AuthorizationCodeCredentials if available,
                # otherwise use direct OAuth2Credentials wrapper
                try:
                    from exchangelib.credentials import OAuth2AuthorizationCodeCredentials
                    token_creds = OAuth2AuthorizationCodeCredentials(access_token={'access_token': access_token, 'token_type': 'Bearer'})
                except (ImportError, TypeError):
                    if OAuth2Credentials is not None:
                        token_creds = OAuth2Credentials(
                            client_id='token-mode', client_secret='token-mode', tenant_id='token-mode',
                            identity=None,
                        )
                    else:
                        raise Exception("当前 exchangelib 版本不支持 OAuth2 Token 模式，请升级 exchangelib >= 4.7")
                if use_auto:
                    account = Account(primary_smtp_address=target_email, credentials=token_creds, autodiscover=True, access_type=access_type_val)
                else:
                    token_config = Configuration(server=config.server if config else 'outlook.office365.com', credentials=token_creds)
                    account = Account(primary_smtp_address=target_email, config=token_config, autodiscover=False, access_type=access_type_val)
            else:
                raise Exception(f"No credentials or token provided for {target_email}")

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

            # Optimization: Larger page size for fewer round-trips
            page_size = 200

            if target_type == "Email":
                def _infer_scope() -> str:
                    s = (mail_folder_scope or "自动 (Auto)").strip().lower()
                    if "common" in s or "常用" in s:
                        return "common"
                    if "sub" in s or "子文件夹" in s:
                        return "inbox_subtree"
                    if "sent" in s or "已发送" in s:
                        return "sent_only"
                    if "outbox" in s or "发件箱" in s:
                        return "outbox_only"
                    if "recoverable" in s or "可恢复" in s:
                        if "purge" in s or "清除" in s:
                            return "recoverable_items_purges"
                        return "recoverable_items_deletions"
                    if "deleted" in s or "已删除" in s:
                        return "deleted_only"
                    if "junk" in s or "垃圾" in s:
                        return "junk_only"
                    if "draft" in s or "草稿" in s:
                        return "drafts_only"
                    if "archive" in s or "存档" in s:
                        return "archive_only"
                    if "inbox" in s or "收件箱" in s:
                        return "inbox_only"
                    if "all" in s or "全邮箱" in s:
                        return "all"
                    return "auto"

                scope = _infer_scope()
                if scope == "auto":
                    # Preserve existing EWS behavior
                    scope = "inbox_only"

                folders = []
                try:
                    if scope == "common":
                        # Common set across mailboxes
                        cand = []
                        try:
                            cand.append(account.inbox)
                        except Exception:
                            pass
                        try:
                            cand.append(getattr(account, "outbox"))
                        except Exception:
                            pass
                        for attr in ("sent", "deleted_items", "junk", "drafts"):
                            try:
                                cand.append(getattr(account, attr))
                            except Exception:
                                pass
                        # Archive is not guaranteed; try attribute then name-search
                        try:
                            cand.append(getattr(account, "archive"))
                        except Exception:
                            try:
                                for f in account.root.walk():
                                    if (getattr(f, 'name', '') or '').strip().lower() in ("archive", "存档"):
                                        cand.append(f)
                                        break
                            except Exception:
                                pass

                        for f in cand:
                            if f and hasattr(f, 'all'):
                                folders.append(f)
                    elif scope == "sent_only":
                        folders = [getattr(account, "sent")]
                    elif scope == "outbox_only":
                        folders = [getattr(account, "outbox")]
                    elif scope == "deleted_only":
                        folders = [getattr(account, "deleted_items")]
                    elif scope == "recoverable_items_deletions":
                        folders = [getattr(account, "recoverable_items_deletions")]
                    elif scope == "recoverable_items_purges":
                        folders = [getattr(account, "recoverable_items_purges")]
                    elif scope == "junk_only":
                        folders = [getattr(account, "junk")]
                    elif scope == "drafts_only":
                        folders = [getattr(account, "drafts")]
                    elif scope == "archive_only":
                        try:
                            folders = [getattr(account, "archive")]
                        except Exception:
                            found = None
                            try:
                                for f in account.root.walk():
                                    if (getattr(f, 'name', '') or '').strip().lower() in ("archive", "存档"):
                                        found = f
                                        break
                            except Exception:
                                found = None
                            folders = [found] if found else []
                    elif scope == "all":
                        for f in account.root.walk():
                            cc = (getattr(f, 'container_class', '') or '')
                            if cc and not str(cc).startswith('IPF.Note'):
                                continue
                            if hasattr(f, 'all'):
                                folders.append(f)
                    elif scope == "inbox_subtree":
                        for f in account.inbox.walk():
                            cc = (getattr(f, 'container_class', '') or '')
                            if cc and not str(cc).startswith('IPF.Note'):
                                continue
                            if hasattr(f, 'all'):
                                folders.append(f)
                    else:
                        folders = [account.inbox]
                except Exception:
                    folders = [account.inbox]

                folders = [f for f in folders if f]

                def _flush_delete_batch(folder, batch_items, batch_rows):
                    if not batch_items:
                        return

                    dt = None
                    if permanent_delete and (DeleteType is not None):
                        dt = getattr(DeleteType, 'PURGE', None) or getattr(DeleteType, 'HARD_DELETE', None)
                    elif (not permanent_delete) and soft_delete and (DeleteType is not None):
                        dt = getattr(DeleteType, 'MOVE_TO_DELETED_ITEMS', None) or getattr(DeleteType, 'MOVE_TO_DELETEDITEMS', None)

                    # Try bulk delete first (much faster)
                    try:
                        if hasattr(folder, 'bulk_delete'):
                            ids = []
                            for it in batch_items:
                                iid = getattr(it, 'id', None)
                                ck = getattr(it, 'changekey', None)
                                if iid and ck:
                                    ids.append((iid, ck))
                                elif iid:
                                    ids.append(iid)
                            if ids:
                                try:
                                    if dt is not None:
                                        folder.bulk_delete(ids, delete_type=dt)
                                    else:
                                        folder.bulk_delete(ids)
                                except TypeError:
                                    # Older exchangelib signatures
                                    folder.bulk_delete(ids)

                                for r in batch_rows:
                                    r['Status'] = 'Success'
                                    with csv_lock:
                                        writer.writerow(r)
                                return
                    except Exception as e:
                        self.log(f"  批量删除失败，回退逐个删除: {e}", "ERROR")

                    # Fallback: per-item delete
                    for it, r in zip(batch_items, batch_rows):
                        try:
                            if dt is not None:
                                it.delete(delete_type=dt)
                            else:
                                it.delete()
                            r['Status'] = 'Success'
                        except Exception as e:
                            r['Status'] = 'Failed'
                            r['Details'] = ((r.get('Details') + '; ') if r.get('Details') else '') + str(e)
                        with csv_lock:
                            writer.writerow(r)

                for folder in folders:
                    batch_items = []
                    batch_rows = []
                    try:
                        qs = folder.all().order_by('-datetime_received')
                        qs.page_size = page_size
                        if start_dt:
                            qs = qs.filter(datetime_received__gte=start_dt)
                        if end_dt:
                            qs = qs.filter(datetime_received__lt=end_dt)
                        if criteria_sender:
                            qs = qs.filter(sender__icontains=criteria_sender)

                        fields = ['id', 'changekey', 'subject', 'sender', 'datetime_received']
                        if criteria_body:
                            fields.append('body')
                        try:
                            qs = qs.only(*fields)
                        except Exception:
                            pass

                        for item in qs:
                            if criteria_body:
                                try:
                                    if criteria_body.lower() not in (item.body or "").lower():
                                        continue
                                except Exception:
                                    continue

                            item_id = getattr(item, 'id', None) or (item.item_id if hasattr(item, 'item_id') else getattr(item, 'message_id', 'Unknown ID'))
                            subject = item.subject
                            sender_val = item.sender.email_address if item.sender else 'Unknown'
                            received_val = getattr(item, 'datetime_received', 'Unknown')
                            row = {
                                'UserPrincipalName': target_email,
                                'MessageId': item_id,
                                'Subject': subject,
                                'Sender': sender_val,
                                'Received': received_val,
                                'Action': 'Report' if report_only else ('PermanentDelete' if permanent_delete else ('SoftDelete' if soft_delete else 'Delete')),
                                'Status': 'Pending',
                                'Details': ''
                            }

                            if report_only:
                                self.log(f"  [报告] 发现: {item.subject}")
                                row['Status'] = 'Skipped'
                                with csv_lock:
                                    writer.writerow(row)
                            else:
                                batch_items.append(item)
                                batch_rows.append(row)
                                if len(batch_items) >= 200:
                                    _flush_delete_batch(folder, batch_items, batch_rows)
                                    batch_items = []
                                    batch_rows = []

                        if batch_items:
                            _flush_delete_batch(folder, batch_items, batch_rows)
                    except Exception as e:
                        self.log(f"  文件夹扫描失败: {getattr(folder, 'name', '')} | {e}", "ERROR")

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
            
            # Execute query for Meeting (Email already streamed above)
            if target_type != "Email":
                items = list(qs)
                if not items:
                    self.log(f"用户 {target_email} 未找到项目。")
            else:
                # Email already handled
                return
            
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
                    if permanent_delete and (target_type == "Email") and (DeleteType is not None):
                        try:
                            dt = getattr(DeleteType, 'PURGE', None) or getattr(DeleteType, 'HARD_DELETE', None)
                            if dt is not None:
                                item.delete(delete_type=dt)
                            else:
                                item.delete()
                        except Exception:
                            item.delete()
                    else:
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

        # Configure Advanced/Expert Logging for EWS
        ews_log_handlers = []
        log_level = self.log_level_var.get()
        
        if log_level in ("Advanced", "Expert"):
            try:
                # Create Handler
                # Route EWS trace into level-specific debug log file
                debug_path = self.logger.get_current_debug_log_path() if self.logger else None
                if not debug_path:
                    debug_path = os.path.join(self.documents_dir, "app_advanced_fallback.log")
                file_handler = logging.FileHandler(debug_path, encoding='utf-8')
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
                if log_level == "Expert":
                    date_str = datetime.now().strftime("%Y-%m-%d")
                    EwsTraceAdapter.response_log_path = os.path.join(self.documents_dir, f"ews_getitem_responses_expert_{date_str}.log")
                else:
                    EwsTraceAdapter.response_log_path = None
                BaseProtocol.HTTP_ADAPTER_CLS = EwsTraceAdapter
                
                # Check permission for response log if Expert
                if log_level == "Expert":
                    try:
                        test_path = EwsTraceAdapter.response_log_path
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
            
            # 1. Connect — support NTLM / Basic / OAuth2 / Token
            server = self._clean_server_address(self.ews_server_var.get())
            use_auto = self.ews_use_autodiscover.get()
            auth_type = self.ews_auth_type_var.get()
            ews_auth_method = self.ews_auth_method_var.get()

            creds, token = self._get_ews_credentials()

            # Determine exchangelib auth_type constant for NTLM vs Basic
            ews_proto_auth_type = None
            if ews_auth_method == "NTLM":
                ews_proto_auth_type = NTLM
            elif ews_auth_method == "Basic":
                ews_proto_auth_type = BASIC

            config = None
            if not use_auto:
                self.log(f"Connecting to server: {server}")
                config_kwargs = {"server": server}
                if creds is not None:
                    config_kwargs["credentials"] = creds
                if ews_proto_auth_type:
                    config_kwargs["auth_type"] = ews_proto_auth_type
                config = Configuration(**config_kwargs)

            # 2. Read CSV
            users = self._get_target_users()
            
            self.log(f"目标列表中共有 {len(users)} 个邮箱。")
            self._progress_reset(len(users))

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
                mail_folder_scope = self.mail_folder_scope_var.get()
                permanent_delete = bool(self.permanent_delete_var.get()) and (not report_only) and (target_type == "Email")
                soft_delete = bool(self.soft_delete_var.get()) and (not report_only) and (target_type == "Email") and (not permanent_delete)
                
                csv_lock = threading.Lock()

                with ThreadPoolExecutor(max_workers=10) as executor:
                    futures = []
                    for target_email in users:
                        futures.append(executor.submit(
                            self.process_single_user_ews,
                            target_email, creds, config, auth_type, use_auto, target_type,
                            start_date_str, end_date_str, criteria_sender, criteria_msg_id,
                            criteria_subject, criteria_body, meeting_only_cancelled, meeting_scope,
                            report_only, writer, csv_lock, log_level, mail_folder_scope, permanent_delete, soft_delete,
                            token
                        ))
                    
                    for future in futures:
                        try:
                            future.result()
                        except Exception as e:
                            self.log(f"Task Error: {e}", "ERROR")
                        self._progress_increment()

            self._progress_finish("EWS 任务完成")
            self.log(f">>> 任务完成。报告: {report_path}")
            
            msg_title = "完成"
            if self.report_only_var.get():
                msg_body = "扫描生成报告任务完成。"
                # Auto-load results into tab 3
                self.root.after(100, self._load_last_report)
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
            EwsTraceAdapter.response_log_path = None

if __name__ == "__main__":
    root = tk.Tk()
    app = UniversalEmailCleanerApp(root)
    root.mainloop()