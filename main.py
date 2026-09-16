# -*- coding: utf-8 -*-
"""
Excel 图片粘贴助手
==================
从微信等外部程序复制图片后，在 Excel 中选中目标单元格，按快捷键(Ctrl+Alt+V)，
图片即按设定的尺寸规则自动插入，无需手动调整大小。

支持三种插入模式：
  1. 适应单元格  —— 图片缩小到刚好放进当前单元格内
  2. 自动加高行  —— 图片宽度撑满单元格列宽，行高自动加高到图片高度
  3. 固定宽度    —— 图片统一缩放到设定的宽度（磅），高度按比例
"""

import ctypes
import ctypes.wintypes as wt
import json
import os
import sys
import tempfile
import threading


# ---------------------------------------------------------------------------
# 自动切换到项目虚拟环境（.venv）
# 用任意 Python 启动本脚本时，若缺少依赖且 .venv 存在，则自动用 .venv 重启自身
# ---------------------------------------------------------------------------

def _bootstrap_venv():
    if os.environ.get("_EXCEL_PASTE_HELPER_IN_VENV"):
        return
    try:
        import PIL  # noqa: F401
        import win32com.client  # noqa: F401
        return  # 依赖齐全，无需切换
    except ImportError:
        pass
    try:
        this_file = os.path.abspath(__file__)
    except NameError:  # 极少数 exec 场景没有 __file__
        this_file = os.path.join(os.getcwd(), "main.py")
    venv_py = os.path.join(
        os.path.dirname(this_file), ".venv", "Scripts", "python.exe")
    if os.path.exists(venv_py):
        os.environ["_EXCEL_PASTE_HELPER_IN_VENV"] = "1"
        os.execv(venv_py, [venv_py, this_file] + sys.argv[1:])
    # 没有 .venv 就提示后退出
    try:
        import tkinter.messagebox as _mb
        _mb.showerror(
            "缺少依赖",
            "未找到依赖库（Pillow / pywin32），且程序文件夹内没有 .venv 虚拟环境。\n"
            "请在程序文件夹打开命令行执行：\n"
            "    python -m venv .venv\n"
            "    .venv\\Scripts\\pip install -r requirements.txt")
    except Exception:
        print("缺少依赖 Pillow/pywin32，且未找到 .venv。请先运行：")
        print("    python -m venv .venv")
        print("    .venv\\Scripts\\pip install -r requirements.txt")
    sys.exit(1)


_bootstrap_venv()

import tkinter as tk
from tkinter import ttk, messagebox

from PIL import Image, ImageGrab

# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------

# 打包成 exe（PyInstaller）后，__file__ 指向临时目录；
# 此时应以 exe 所在目录为准，配置/日志/临时图片都放 exe 旁边
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
LOG_PATH = os.path.join(APP_DIR, "log.txt")
CACHE_DIR = os.path.join(APP_DIR, "cache")          # 原图缓存，供双击预览用
CACHE_INDEX = os.path.join(CACHE_DIR, "index.json")
CACHE_MAX = 400                                     # 缓存条目上限

DEFAULT_CONFIG = {
    "hotkey_mod": "ctrl+alt",      # ctrl+alt / ctrl+shift / alt
    "hotkey_key": "v",
    "intercept_ctrl_v": True,      # 拦截 Excel 中的 Ctrl+V：剪贴板有图片时自动按规则插入
    "dblclick_preview": True,      # 双击 Excel 中的图片弹出大图预览
    "mode": "rowheight",           # fit | rowheight | fixed
    "padding": 3,                  # 图片与单元格边缘留白（磅），四边各留这么多
    "fixed_width": 120,            # 固定宽度模式的宽度（磅）
    "no_upscale": False,           # 小图不放大
}


def file_log(msg):
    """写入文件日志，便于排查问题（保留最近 200 行）"""
    try:
        import datetime
        lines = []
        if os.path.exists(LOG_PATH):
            with open(LOG_PATH, "r", encoding="utf-8", errors="ignore") as f:
                lines = f.read().splitlines()
        lines.append(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}")
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write("\n".join(lines[-200:]) + "\n")
    except Exception:
        pass

VK_MAP = {c: 0x41 + i for i, c in enumerate("ABCDEFGHIJKLMNOPQRSTUVWXYZ")}
VK_MAP.update({str(d): 0x30 + d for d in range(10)})
MOD_BITS = {
    "ctrl": 0x0002,
    "alt": 0x0001,
    "shift": 0x0004,
}
WM_HOTKEY = 0x0312
HOTKEY_ID = 0xB001


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 核心：取剪贴板图片 -> 插入 Excel
# ---------------------------------------------------------------------------

def get_clipboard_image():
    """返回 PIL.Image 或 None"""
    try:
        data = ImageGrab.grabclipboard()
    except Exception:
        return None
    if isinstance(data, Image.Image):
        return data
    if isinstance(data, list) and data:
        for p in data:
            try:
                return Image.open(p).convert("RGB")
            except Exception:
                continue
    return None


def connect_excel():
    """连接正在运行的表格程序：优先微软 Excel，其次 WPS 表格"""
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    last_err = None
    for progid in ("Excel.Application", "ket.Application", "et.Application"):
        try:
            return win32com.client.GetActiveObject(progid)
        except Exception as e:
            last_err = e
    pythoncom.CoUninitialize()
    raise last_err


def do_paste(cfg, log):
    """执行一次粘贴，log 为回调函数 log(msg)"""
    img = get_clipboard_image()
    if img is None:
        log("剪贴板中没有图片。请先从微信等处复制图片，再按快捷键。")
        return False

    # 存为临时 PNG
    tmp = os.path.join(tempfile.gettempdir(), "excel_paste_helper.png")
    img.save(tmp, "PNG")

    try:
        excel = connect_excel()
    except Exception:
        log("未检测到正在运行的 Excel。请先打开 Excel 并选中一个单元格。")
        return False

    try:
        cell = excel.ActiveCell
        ws = excel.ActiveSheet
        cw = float(cell.Width)    # 列宽（磅）
        ch = float(cell.Height)   # 行高（磅）
        iw, ih = img.size
        pad = float(cfg.get("padding", 2))

        if cfg.get("mode") == "fit":
            # 缩到刚好放进单元格
            scale = min((cw - 2 * pad) / iw, (ch - 2 * pad) / ih)
            if cfg.get("no_upscale"):
                scale = min(scale, 72.0 / 96.0)  # 96dpi 原尺寸（磅）
            w, h = iw * scale, ih * scale
        elif cfg.get("mode") == "fixed":
            w = float(cfg.get("fixed_width", 120))
            h = w * ih / iw
        else:  # rowheight：宽度撑满列宽，行高自动加高
            w = max(cw - 2 * pad, 10)
            h = w * ih / iw
            need = h + 2 * pad
            if cell.EntireRow.RowHeight < need:
                cell.EntireRow.RowHeight = need

        # 关键：图片不能贴死在单元格左上角，否则会压住左、上两条网格线。
        # 锚点向内偏移 pad，配合已按 pad 收窄的宽高，四边留白才对称。
        left = float(cell.Left) + pad
        top = float(cell.Top) + pad
        shape = ws.Shapes.AddPicture(tmp, False, True, left, top, w, h)
        # 收尾步骤单独隔离：即使失败，图片也已经插入成功
        try:
            shape.Placement = 1  # xlMoveAndSize，随单元格移动缩放
        except Exception as e:
            log(f"提示：设置图片随单元格移动失败（不影响插入）：{e}")
        try:
            cache_original_image(img, ws.Parent.Name, ws.Name, shape.Name)
        except Exception as e:
            file_log(f"缓存原图失败（仅影响双击预览清晰度）：{e}")
        try:
            cell.Offset(1, 0).Select()  # 自动跳到下一行，方便连续粘贴
        except Exception as e:
            log(f"提示：跳转下一行失败（不影响插入）：{e}")
        log(f"已插入图片 {iw}x{ih}px -> {w:.0f}x{h:.0f}磅，四边留白 {pad:.0f}磅 @ {cell.Address}")
        return True
    except Exception as e:
        log(f"插入失败：{e}")
        return False


def get_selected_picture_image(log):
    """若 Excel 当前选中的是图片，返回 (PIL.Image, 说明)，否则 (None, 原因)"""
    import time as _time
    try:
        excel = connect_excel()
    except Exception:
        return None, "未检测到正在运行的 Excel"
    try:
        sel = excel.Selection
    except Exception as e:
        return None, f"读取选中对象失败：{e}"
    # 单元格选区有 Address 属性；图片（Shape/Picture）没有
    if hasattr(sel, "Address"):
        return None, "当前选中的是单元格，不是图片"
    if not (hasattr(sel, "Width") and hasattr(sel, "Height")):
        return None, "当前选中的不是图片"

    # 优先取插入时缓存的原图（分辨率最高、最清晰）
    try:
        ws_obj = sel.Parent
        wb_obj = ws_obj.Parent
        cached = lookup_cached_image(wb_obj.Name, ws_obj.Name, sel.Name)
        if cached is not None:
            return cached, "ok-原图缓存"
    except Exception:
        pass

    # 回退：把图片复制到剪贴板再读（只有屏幕分辨率，会偏糊）
    try:
        sel.CopyPicture(1, 2)          # xlScreen, xlBitmap
        _time.sleep(0.3)
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image):
            return img, "ok-剪贴板"
        return None, "已选中图片，但读取图像数据失败"
    except Exception as e:
        return None, f"读取图片失败：{e}"


def _load_cache_index():
    try:
        with open(CACHE_INDEX, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_cache_index(idx):
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(CACHE_INDEX, "w", encoding="utf-8") as f:
            json.dump(idx, f, ensure_ascii=False)
    except Exception:
        pass


def cache_original_image(img, wb_name, ws_name, shape_name):
    """插入图片时把原图存进缓存，双击预览时就能用原始分辨率"""
    import time as _time
    os.makedirs(CACHE_DIR, exist_ok=True)
    fname = f"img_{int(_time.time() * 1000)}.png"
    img.save(os.path.join(CACHE_DIR, fname), "PNG")
    idx = _load_cache_index()
    idx[f"{wb_name}|{ws_name}|{shape_name}"] = fname
    # 超出上限时，删掉最旧的
    if len(idx) > CACHE_MAX:
        for k in list(idx.keys())[:len(idx) - CACHE_MAX]:
            old = idx.pop(k, None)
            if old:
                try:
                    os.remove(os.path.join(CACHE_DIR, old))
                except Exception:
                    pass
    _save_cache_index(idx)


def lookup_cached_image(wb_name, ws_name, shape_name):
    idx = _load_cache_index()
    fname = idx.get(f"{wb_name}|{ws_name}|{shape_name}")
    if not fname:
        return None
    path = os.path.join(CACHE_DIR, fname)
    if not os.path.exists(path):
        return None
    try:
        im = Image.open(path)
        im.load()
        return im
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 全局热键（RegisterHotKey） + Ctrl+V 拦截（低级键盘钩子）
#                            + 双击图片预览（低级鼠标钩子）
# ---------------------------------------------------------------------------

# Excel / WPS 表格主窗口类名
EXCEL_WND_CLASSES = ("XLMAIN", "ket_main", "ET_Main")
VK_V = 0x56
VK_CONTROL = 0x11
VK_MENU = 0x12  # Alt
VK_SHIFT = 0x10
WH_KEYBOARD_LL = 13
WH_MOUSE_LL = 14
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
CF_BITMAP, CF_DIB, CF_DIBV5 = 2, 8, 17
DEBOUNCE_SEC = 0.8  # 按住不放 / 连按时防重复插入


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [("vkCode", wt.DWORD),
                ("scanCode", wt.DWORD),
                ("flags", wt.DWORD),
                ("time", wt.DWORD),
                ("dwExtraInfo", ctypes.POINTER(wt.ULONG))]


class POINT(ctypes.Structure):
    _fields_ = [("x", wt.LONG), ("y", wt.LONG)]


class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [("pt", POINT),
                ("mouseData", wt.DWORD),
                ("flags", wt.DWORD),
                ("time", wt.DWORD),
                ("dwExtraInfo", ctypes.POINTER(wt.ULONG))]


LRESULT = ctypes.c_ssize_t
HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wt.WPARAM, wt.LPARAM)


def _setup_user32():
    u = ctypes.windll.user32
    u.SetWindowsHookExW.restype = ctypes.c_void_p
    u.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC,
                                    ctypes.c_void_p, wt.DWORD]
    u.CallNextHookEx.restype = LRESULT
    u.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int,
                                 wt.WPARAM, wt.LPARAM]
    u.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
    return u


class KeyInterceptor:
    """三引擎：
    1. RegisterHotKey 注册 Ctrl+Alt+V 手动触发
    2. WH_KEYBOARD_LL 低级钩子：Excel 前台按 Ctrl+V 且剪贴板是图片时，
       吞掉按键改为按规则插入；剪贴板是文字等则完全放行，零干扰
    3. WH_MOUSE_LL 低级鼠标钩子：Excel 前台双击图片时，吞掉第二次点击
       （避免弹出"设置图片格式"面板）并触发大图预览
    """

    def __init__(self, on_trigger):
        self.on_trigger = on_trigger
        self.thread = None
        self.running = False
        self.hook = None
        self.mouse_hook = None
        self._proc = None          # 防止回调被 GC
        self._mproc = None     # 鼠标钩子回调引用（防GC）
        self.user32 = _setup_user32()
        self.intercept_ctrl_v = True
        self.dblclick_preview = True
        self._last_trigger = 0.0
        self._last_click = 0.0     # 上一次左键按下时间
        self._last_click_pos = (0, 0)
        self._suppress_up = False  # 吞掉配对的上抬事件

    # -- 快捷键解析 ----------------------------------------------------------
    def _parse(self, cfg):
        mods = cfg.get("hotkey_mod", "ctrl+alt").lower().split("+")
        mod = 0
        for m in mods:
            mod |= MOD_BITS.get(m.strip(), 0)
        key = cfg.get("hotkey_key", "v").upper()
        return mod, VK_MAP.get(key, ord("V"))

    def _is_excel_foreground(self):
        hwnd = self.user32.GetForegroundWindow()
        buf = ctypes.create_unicode_buffer(64)
        self.user32.GetClassNameW(hwnd, buf, 64)
        return buf.value in EXCEL_WND_CLASSES, buf.value

    # -- 低级鼠标钩子回调（必须快速返回，不做 COM） --------------------------
    def _mouse_proc(self, nCode, wParam, lParam):
        if nCode == 0:
            try:
                if wParam == WM_LBUTTONUP and self._suppress_up:
                    self._suppress_up = False
                    return 1  # 吞掉双击的第二下上抬
                if wParam == WM_LBUTTONDOWN:
                    if self._suppress_up:
                        self._suppress_up = False
                        return 1
                    if self.dblclick_preview:
                        ms = ctypes.cast(ctypes.c_void_p(lParam),
                                         ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                        x, y = ms.pt.x, ms.pt.y
                        import time
                        now = time.monotonic()
                        dbl_time = self.user32.GetDoubleClickTime() / 1000.0
                        dx = abs(x - self._last_click_pos[0])
                        dy = abs(y - self._last_click_pos[1])
                        if (now - self._last_click) <= dbl_time and dx <= 6 and dy <= 6:
                            self._last_click = 0.0
                            is_excel, _cls = self._is_excel_foreground()
                            if is_excel:
                                self._suppress_up = True
                                self.on_trigger("dblclick")
                                return 1  # 吞掉第二下按下
                        self._last_click = now
                        self._last_click_pos = (x, y)
            except Exception:
                pass
        return self.user32.CallNextHookEx(None, nCode, wParam, lParam)

    # -- 低级键盘钩子回调（必须快速返回） ------------------------------------
    def _hook_proc(self, nCode, wParam, lParam):
        if nCode == 0 and wParam == WM_KEYDOWN and self.intercept_ctrl_v:
            try:
                kb = ctypes.cast(ctypes.c_void_p(lParam),
                                 ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                if kb.vkCode == VK_V:
                    u = self.user32
                    ctrl = u.GetAsyncKeyState(VK_CONTROL) & 0x8000
                    alt = u.GetAsyncKeyState(VK_MENU) & 0x8000
                    shift = u.GetAsyncKeyState(VK_SHIFT) & 0x8000
                    if ctrl and not alt and not shift:
                        # 前台窗口是否为 Excel / WPS 表格
                        hwnd = u.GetForegroundWindow()
                        buf = ctypes.create_unicode_buffer(64)
                        u.GetClassNameW(hwnd, buf, 64)
                        # 剪贴板是否有图片（快速检查，不打开剪贴板）
                        has_img = any(
                            u.IsClipboardFormatAvailable(f)
                            for f in (CF_BITMAP, CF_DIB, CF_DIBV5))
                        file_log(f"DEBUG Ctrl+V: 前台窗口类={buf.value!r} "
                                 f"剪贴板有图={has_img}")
                        if buf.value in EXCEL_WND_CLASSES:
                            if has_img:
                                import time
                                now = time.monotonic()
                                if now - self._last_trigger > DEBOUNCE_SEC:
                                    self._last_trigger = now
                                    self.on_trigger("ctrlv")
                                return 1  # 吞掉按键，阻止 Excel 原生粘贴
            except Exception as e:
                file_log(f"钩子异常: {e}")
        return self.user32.CallNextHookEx(None, nCode, wParam, lParam)

    # -- 监听线程 ------------------------------------------------------------
    def _loop(self, mod, vk):
        u = self.user32

        registered = u.RegisterHotKey(None, HOTKEY_ID, mod, vk)
        if not registered:
            file_log("快捷键注册失败（可能被其他程序占用），Ctrl+V 拦截不受影响")
            self.on_trigger(None)

        self._proc = HOOKPROC(self._hook_proc)
        self.hook = u.SetWindowsHookExW(WH_KEYBOARD_LL, self._proc, None, 0)
        if not self.hook:
            file_log("Ctrl+V 拦截钩子安装失败，仅快捷键模式可用")
            self.on_trigger("nohook")

        # 鼠标钩子：双击图片预览
        self._mproc = HOOKPROC(self._mouse_proc)
        self.mouse_hook = u.SetWindowsHookExW(WH_MOUSE_LL, self._mproc, None, 0)
        if not self.mouse_hook:
            file_log("双击预览钩子安装失败，双击功能不可用")

        msg = wt.MSG()
        while self.running and u.GetMessageW(
                ctypes.byref(msg), None, 0, 0) not in (-1, 0):
            if msg.message == WM_HOTKEY and msg.wParam == HOTKEY_ID:
                self.on_trigger("hotkey")

        if self.hook:
            u.UnhookWindowsHookEx(self.hook)
            self.hook = None
        if self.mouse_hook:
            u.UnhookWindowsHookEx(self.mouse_hook)
            self.mouse_hook = None
        if registered:
            u.UnregisterHotKey(None, HOTKEY_ID)

    def start(self, cfg):
        self.stop()
        mod, vk = self._parse(cfg)
        self.intercept_ctrl_v = bool(cfg.get("intercept_ctrl_v", True))
        self.dblclick_preview = bool(cfg.get("dblclick_preview", True))
        self.running = True
        self.thread = threading.Thread(target=self._loop, args=(mod, vk), daemon=True)
        self.thread.start()
        file_log(f"监听已启动（Ctrl+V 拦截: {'开' if self.intercept_ctrl_v else '关'}"
                 f"，双击预览: {'开' if self.dblclick_preview else '关'}）")

    def stop(self):
        self.running = False
        if self.thread and self.thread.is_alive():
            # 向监听线程投递 WM_QUIT 让消息循环退出
            tid = self.thread.ident
            if tid:
                ctypes.windll.user32.PostThreadMessageW(tid, 0x0012, 0, 0)
            self.thread.join(timeout=1.0)
        self.thread = None


# ---------------------------------------------------------------------------
# 界面
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.queue = []

        root.title("Excel 图片粘贴助手")
        root.geometry("440x580")
        root.resizable(False, False)

        frm = ttk.Frame(root, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="插入模式", font=("Microsoft YaHei UI", 10, "bold")).grid(
            column=0, row=0, columnspan=2, sticky="w", pady=(0, 4))

        self.mode_var = tk.StringVar(value=self.cfg["mode"])
        modes = [
            ("rowheight", "自动加高行（推荐）：图片撑满列宽，行高自动适配"),
            ("fit", "适应单元格：图片缩小放进当前单元格内"),
            ("fixed", "固定宽度：所有图片统一宽度，高度按比例"),
        ]
        for i, (val, text) in enumerate(modes):
            ttk.Radiobutton(frm, text=text, value=val, variable=self.mode_var).grid(
                column=0, row=1 + i, columnspan=2, sticky="w", pady=1)

        ttk.Separator(frm).grid(column=0, row=4, columnspan=2, sticky="ew", pady=8)

        ttk.Label(frm, text="参数设置", font=("Microsoft YaHei UI", 10, "bold")).grid(
            column=0, row=5, columnspan=2, sticky="w", pady=(0, 4))

        self.pad_var = tk.DoubleVar(value=self.cfg["padding"])
        ttk.Label(frm, text="单元格留白（磅）").grid(column=0, row=6, sticky="w", pady=2)
        ttk.Spinbox(frm, from_=0, to=20, textvariable=self.pad_var, width=8).grid(
            column=1, row=6, sticky="e", pady=2)

        self.fw_var = tk.DoubleVar(value=self.cfg["fixed_width"])
        ttk.Label(frm, text="固定宽度（磅，仅固定宽度模式）").grid(column=0, row=7, sticky="w", pady=2)
        ttk.Spinbox(frm, from_=20, to=600, increment=10, textvariable=self.fw_var, width=8).grid(
            column=1, row=7, sticky="e", pady=2)

        self.no_up_var = tk.BooleanVar(value=self.cfg["no_upscale"])
        ttk.Checkbutton(frm, text="小图不放大（仅适应单元格模式）",
                        variable=self.no_up_var).grid(
            column=0, row=8, columnspan=2, sticky="w", pady=2)

        self.hk_var = tk.StringVar(value=self.cfg["hotkey_mod"])
        ttk.Label(frm, text="备用快捷键组合").grid(column=0, row=9, sticky="w", pady=2)
        hk = ttk.Combobox(frm, textvariable=self.hk_var, width=14, state="readonly",
                          values=["ctrl+alt", "ctrl+shift", "alt", "ctrl"])
        hk.grid(column=1, row=9, sticky="e", pady=2)

        self.icv_var = tk.BooleanVar(value=self.cfg.get("intercept_ctrl_v", True))
        tk.Checkbutton(frm, text="拦截 Excel 中的 Ctrl+V（推荐：粘贴图片自动缩放，粘贴文字不受影响）",
                       variable=self.icv_var, wraplength=380, anchor="w",
                       justify="left").grid(
            column=0, row=10, columnspan=2, sticky="w", pady=2)

        self.dbl_var = tk.BooleanVar(value=self.cfg.get("dblclick_preview", True))
        tk.Checkbutton(frm, text="双击图片弹出大图预览（双击或 Esc 关闭）",
                       variable=self.dbl_var, wraplength=380, anchor="w",
                       justify="left").grid(
            column=0, row=11, columnspan=2, sticky="w", pady=2)

        ttk.Separator(frm).grid(column=0, row=12, columnspan=2, sticky="ew", pady=8)

        btns = ttk.Frame(frm)
        btns.grid(column=0, row=13, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Button(btns, text="保存设置", command=self.save).pack(side="left", padx=(0, 6))
        ttk.Button(btns, text="立即插入一次", command=lambda: self.request_paste()).pack(side="left")

        ttk.Label(frm, text="运行日志", font=("Microsoft YaHei UI", 10, "bold")).grid(
            column=0, row=14, columnspan=2, sticky="w")
        self.log_text = tk.Text(frm, height=8, width=52, state="disabled",
                                font=("Consolas", 9))
        self.log_text.grid(column=0, row=15, columnspan=2, sticky="nsew")

        self.status = tk.StringVar(value="就绪 | Excel 中按 Ctrl+V 自动缩放图片")
        ttk.Label(frm, textvariable=self.status, foreground="#2a7").grid(
            column=0, row=16, columnspan=2, sticky="w", pady=(4, 0))

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(15, weight=1)

        self.hotkey = KeyInterceptor(lambda v: self.queue.append(v))
        self.hotkey.start(self.cfg)
        self.root.after(60, self.poll)
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -- 日志 / 队列 --------------------------------------------------------
    def log(self, msg):
        import datetime
        stamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stamp}] {msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        file_log(msg)

    def poll(self):
        while self.queue:
            v = self.queue.pop(0)
            if v is None:
                self.status.set("备用快捷键注册失败（Ctrl+V 拦截不受影响）")
                messagebox.showwarning(
                    "提示", "备用快捷键注册失败，可能被其他程序占用。\n"
                            "不影响 Ctrl+V 拦截功能，也可在设置中换一个组合键。")
            elif v == "nohook":
                self.status.set("Ctrl+V 拦截不可用，仅备用快捷键生效")
                self.log("Ctrl+V 拦截钩子安装失败，请重启程序再试")
            elif v == "dblclick":
                self.request_preview()
            else:
                source = "Ctrl+V 拦截" if v == "ctrlv" else "快捷键"
                self.log(f"触发方式：{source}")
                self.request_paste()
        self.root.after(60, self.poll)

    # -- 动作 ---------------------------------------------------------------
    def request_paste(self):
        self.status.set("正在插入…")
        cfg = self._collect()
        threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()

    def request_preview(self):
        threading.Thread(target=self._preview_worker, daemon=True).start()

    def _preview_worker(self):
        img, reason = get_selected_picture_image(
            lambda m: self.root.after(0, lambda m=m: self.log(m)))
        if img is None:
            self.root.after(0, lambda: self.log(f"双击未预览：{reason}"))
            return
        self.root.after(0, lambda: self.show_preview(img))

    def show_preview(self, pil_img):
        """置顶大图预览窗口：双击 / 单击 / Esc 关闭"""
        from PIL import ImageTk
        try:
            resample = Image.Resampling.LANCZOS
        except AttributeError:
            resample = Image.LANCZOS

        win = tk.Toplevel(self.root)
        win.title("图片预览")
        win.configure(bg="#1a1a1a")
        win.attributes("-topmost", True)

        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        maxw, maxh = int(sw * 0.78), int(sh * 0.78)
        iw, ih = pil_img.size
        scale = min(maxw / iw, maxh / ih)      # 小图放大、大图缩小，都适配屏幕
        nw, nh = max(1, int(iw * scale)), max(1, int(ih * scale))
        photo = ImageTk.PhotoImage(pil_img.resize((nw, nh), resample))
        win._photo = photo                      # 防止被回收

        tk.Label(win, text=f"{iw} × {ih} 像素　·　双击或按 Esc 关闭",
                 bg="#1a1a1a", fg="#bbbbbb",
                 font=("Microsoft YaHei UI", 9)).pack(pady=(6, 4))
        lbl = tk.Label(win, image=photo, bg="#1a1a1a", bd=0)
        lbl.pack(padx=8, pady=(0, 8))
        win.geometry(f"{nw + 24}x{nh + 48}+{(sw - nw - 24) // 2}+{(sh - nh - 48) // 2}")

        def close(_e=None):
            try:
                win.destroy()
            except Exception:
                pass

        for w in (win, lbl):
            w.bind("<Double-Button-1>", close)
            w.bind("<Button-1>", close)
            w.bind("<Escape>", close)
        win.focus_force()
        self.log(f"已弹出预览：{iw}×{ih} 像素（显示为 {nw}×{nh}）")

    def _worker(self, cfg):
        try:
            ok = do_paste(cfg, lambda m: self.root.after(0, lambda m=m: self.log(m)))
            self.root.after(0, lambda: self.status.set(
                "完成 ✓ | Ctrl+V / Ctrl+Alt+V 均可" if ok else "未插入 | 详见日志"))
        except Exception as e:
            msg = str(e)
            self.root.after(0, lambda: self.log(f"错误：{msg}"))
            self.root.after(0, lambda: self.status.set("出错 | 详见日志"))

    def _collect(self):
        return {
            "hotkey_mod": self.hk_var.get(),
            "hotkey_key": self.cfg.get("hotkey_key", "v"),
            "intercept_ctrl_v": self.icv_var.get(),
            "dblclick_preview": self.dbl_var.get(),
            "mode": self.mode_var.get(),
            "padding": self.pad_var.get(),
            "fixed_width": self.fw_var.get(),
            "no_upscale": self.no_up_var.get(),
        }

    def save(self):
        self.cfg = self._collect()
        save_config(self.cfg)
        self.hotkey.start(self.cfg)
        self.log(f"设置已保存并生效（模式：{self.cfg['mode']}）")

    def on_close(self):
        self.hotkey.stop()
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        from tkinter import font as tkfont
        tkfont.nametofont("TkDefaultFont").configure(family="Microsoft YaHei UI", size=10)
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
