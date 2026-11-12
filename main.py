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
import io
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
AMBIGUOUS_MARK = "!AMBIGUOUS"                       # 索引里表示"此键有多个候选"

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


def ensure_excel_no_compress():
    """关掉 Excel 的"保存时自动压缩图片"，否则插入的图会被降采样。

    Excel 默认会按图片的显示尺寸重采样到约 220 ppi：
    一张 2000x1500 的照片插进 45 磅宽的单元格后，只剩 138x103 像素，
    在 Excel 里放大或打印就是糊的。
    关掉后需重启 Excel 才生效。
    返回 (ok, 说明文字)
    """
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Office") as k:
            vers, i = [], 0
            while True:
                try:
                    name = winreg.EnumKey(k, i)
                except OSError:
                    break
                i += 1
                if name.replace(".", "").isdigit():
                    vers.append(name)
    except OSError:
        return False, "未找到 Office 注册表（可能未安装微软 Office）"
    if not vers:
        return False, "未找到 Office 注册表（可能未安装微软 Office）"

    changed = []
    for v in vers:
        path = r"Software\Microsoft\Office\%s\Excel\Options" % v
        try:
            with winreg.CreateKeyEx(
                    winreg.HKEY_CURRENT_USER, path, 0,
                    winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE) as k:
                try:
                    cur, _ = winreg.QueryValueEx(
                        k, "AutomaticPictureCompressionDefault")
                except FileNotFoundError:
                    cur = None
                if cur != 0:
                    winreg.SetValueEx(k, "AutomaticPictureCompressionDefault",
                                      0, winreg.REG_DWORD, 0)
                    changed.append(v)
        except OSError:
            continue

    if changed:
        return True, ("已关闭 Excel 图片自动压缩（Office %s），"
                      "重启 Excel 后生效" % "/".join(changed))
    return True, "Excel 图片自动压缩已处于关闭状态"


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

        # 关键 1：图片不能贴死在单元格左上角，否则会压住左、上两条网格线。
        #         锚点向内偏移 pad，配合已按 pad 收窄的宽高，四边留白才对称。
        left = float(cell.Left) + pad
        top = float(cell.Top) + pad

        # 关键 2：必须先按"原始像素尺寸"插入，再把显示尺寸改小。
        #   Excel 在 AddPicture 的那一刻就会按传入的 Width/Height 重采样图片：
        #   直接传 45 磅宽时，5712x4284 的手机照片会被压成 138x103 像素，
        #   在 Excel 里放大或打印就会糊。先按 1 像素 = 72/96 磅插入
        #   （不触发重采样），再用 Width/Height 缩放显示，原始分辨率完整保留。
        raw_w = max(iw * 72.0 / 96.0, 1.0)
        raw_h = max(ih * 72.0 / 96.0, 1.0)
        cap = 12000.0  # Excel 图形尺寸上限约 169 英寸，留余量
        if raw_w > cap or raw_h > cap:
            k = min(cap / raw_w, cap / raw_h)
            raw_w, raw_h = raw_w * k, raw_h * k
        shape = ws.Shapes.AddPicture(tmp, False, True, left, top, raw_w, raw_h)
        try:
            shape.LockAspectRatio = 0   # msoFalse，允许宽高分别设置
        except Exception:
            pass
        shape.Width = w
        shape.Height = h
        # 收尾步骤单独隔离：即使失败，图片也已经插入成功
        try:
            shape.Placement = 1  # xlMoveAndSize，随单元格移动缩放
        except Exception as e:
            log(f"提示：设置图片随单元格移动失败（不影响插入）：{e}")
        try:
            # 把原图存进缓存，并把缓存文件名写进图形的"替代文字"。
            # 这样标记随图片一起存进 xlsx —— 另存为改名、关掉重开、
            # 甚至换电脑把 cache 目录一起拷过去，双击预览都能找回原图。
            fname = cache_original_image(img, ws.Parent.Name, ws.Name, shape.Name)
            if fname:
                try:
                    shape.AlternativeText = "EPH|%s|%dx%d" % (fname, iw, ih)
                except Exception:
                    pass
        except Exception as e:
            file_log(f"缓存原图失败（仅影响双击预览清晰度）：{e}")
        try:
            # 自动跳到下一行，方便连续粘贴。
            # Offset(...).Select() 在部分 Excel / WPS 版本会抛 COM 异常，逐级降级
            moved = False
            try:
                cell.Offset(1, 0).Select()
                moved = True
            except Exception:
                pass
            if not moved:
                ws.Cells(int(cell.Row) + 1, int(cell.Column)).Select()
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

    # ① 首选：插入时写进"替代文字"的缓存标记。
    #    这是唯一随图片存进 xlsx 的线索，另存为改名 / 关掉重开都能命中。
    try:
        tag = str(sel.AlternativeText or "").strip()
        if tag.startswith("EPH|"):
            parts = tag.split("|")
            if len(parts) >= 2:
                im = _open_cached(parts[1].strip())
                if im is not None:
                    return im, "ok-原图缓存(标记)"
                log("原图缓存文件已丢失，尝试其它方式")
    except Exception:
        pass

    # ② 其次：查索引（含另存为改名后的宽松匹配）
    try:
        ws_obj = sel.Parent
        wb_obj = ws_obj.Parent
        fname = lookup_cached_fname(wb_obj.Name, ws_obj.Name, sel.Name)
        if fname:
            im = _open_cached(fname)
            if im is not None:
                # 老版本插入的图没有标记，这里顺手补写一个，
                # 下次打开（哪怕再改名）就不用再靠猜了
                try:
                    iw, ih = im.size
                    sel.AlternativeText = "EPH|%s|%dx%d" % (fname, iw, ih)
                except Exception:
                    pass
                return im, "ok-原图缓存(索引)"
    except Exception:
        pass

    # ③ 再次：直接从已保存的 xlsx 包里解出这张图的内嵌原图。
    #    完全不依赖 cache —— 换电脑、清过缓存、只拷走一个文件都能用。
    try:
        im = _image_from_xlsx_media(sel)
        if im is not None:
            return im, "ok-xlsx内嵌原图"
    except Exception:
        pass

    # ④ 兜底 1：Excel 复制图片时，剪贴板里可能带一份原图 PNG
    try:
        sel.Copy()
        _time.sleep(0.25)
        data = _clipboard_png_bytes()
        if data:
            im = Image.open(io.BytesIO(data))
            im.load()
            if min(im.size) >= 64:
                return im, "ok-剪贴板PNG"
    except Exception:
        pass

    # ⑤ 兜底 2：屏幕分辨率截图。拿到的只是图片"显示尺寸"的位图，
    #    一张 45 磅宽的图只有 20x11 像素，放大就是一坨马赛克。
    #    所以低于 64 像素直接判定失败，宁可明确报错也不弹糊图。
    try:
        sel.CopyPicture(1, 2)          # xlScreen, xlBitmap
        _time.sleep(0.3)
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image):
            if min(img.size) < 64:
                return None, ("没找到这张图的原图缓存（只能取到 %dx%d 的屏幕缩略图）。"
                              "原图缓存在程序目录的 cache 文件夹里，"
                              "请确认没有删除它，且是在插入图片的那台电脑上预览"
                              % img.size)
            return img, "ok-剪贴板"
        return None, "已选中图片，但读取图像数据失败"
    except Exception as e:
        return None, f"读取图片失败：{e}"


def _update_rels(xml):
    """解析 .rels 里的 <Relationship>，返回 [{'Id':..,'Target':..,'Type':..}]"""
    import re as _re
    out = []
    for m in _re.finditer(r"<Relationship\b([^>]*)/?>", xml):
        d = {}
        for am in _re.finditer(r'(\w+)="([^"]*)"', m.group(1)):
            d[am.group(1)] = am.group(2)
        out.append(d)
    return out


def _rel_path(p):
    """xl/drawings/drawing1.xml -> xl/drawings/_rels/drawing1.xml.rels"""
    d, f = p.rsplit("/", 1)
    return "%s/_rels/%s.rels" % (d, f)


def _norm_zip_path(target):
    """把 rels 里的相对 Target 规范成 zip 内的绝对路径"""
    target = target.replace("\\", "/")
    while target.startswith("../"):
        target = target[3:]
    if target.startswith("/"):
        return target.lstrip("/")
    return "xl/" + target


def _image_from_xlsx_media(sel):
    """直接从已保存的 xlsx/xlsm 包里解出这张图对应的原始媒体文件。

    完全不依赖 cache —— 换电脑、清过缓存、拷走单个文件，都能取回原图。
    解析链路：workbook.xml -> sheetN.xml -> drawingN.xml -> media/imageN.*

    定位这一张图时不能用 Shape.Name：中文版 Excel 的 COM 返回英文
    "Picture 2"，而文件里存的是本地化名"图片 2"，两者对不上。
    所以优先用我们写进"替代文字"的 EPH 标记（唯一且不随语言变），
    其次用 Shape.ID 对应 XML 里的 cNvPr id。
    失败返回 None（任何一步对不上都安静放弃，交给后面的兜底）。
    """
    import re as _re
    import zipfile
    try:
        ws = sel.Parent
        wb = ws.Parent
        path = str(wb.FullName or "")
        if not os.path.exists(path):
            return None
        if not path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
            return None
        ws_name = str(ws.Name)
        marker = ""
        shape_id = None
        try:
            marker = str(getattr(sel, "AlternativeText", "") or "").strip()
        except Exception:
            marker = ""
        try:
            shape_id = int(sel.ID)
        except Exception:
            shape_id = None

        with zipfile.ZipFile(path) as z:
            wbxml = z.read("xl/workbook.xml").decode("utf-8", "ignore")
            # 按工作表名找 r:id（比用 Index 猜 sheetN.xml 可靠）
            m = _re.search(
                r'<sheet\b[^>]*\bname="%s"[^>]*\br:id="(rId\d+)"' % _re.escape(ws_name),
                wbxml)
            if not m:
                return None
            wb_rels = _update_rels(
                z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore"))
            sheet_target = None
            for r in wb_rels:
                if r.get("Id") == m.group(1):
                    sheet_target = r.get("Target")
                    break
            if not sheet_target:
                return None
            sheet_path = _norm_zip_path(sheet_target)

            # 工作表 -> drawing
            s_rels = _update_rels(
                z.read(_rel_path(sheet_path)).decode("utf-8", "ignore"))
            draw_target = None
            for r in s_rels:
                if "/drawing" in (r.get("Type") or ""):
                    draw_target = r.get("Target")
                    break
            if not draw_target:
                return None
            draw_path = _norm_zip_path(draw_target)
            draw_xml = z.read(draw_path).decode("utf-8", "ignore")

            # drawing 里定位到这一张图
            blip = None
            for pm in _re.finditer(r"<xdr:pic\b.*?</xdr:pic>", draw_xml, _re.S):
                blk = pm.group(0)
                hit = False
                if marker and marker in blk:
                    hit = True            # 自己写的标记，唯一且不受语言影响
                elif shape_id is not None:
                    nm = _re.search(r'<xdr:cNvPr\b[^>]*\bid="(\d+)"', blk)
                    if nm and int(nm.group(1)) == shape_id:
                        hit = True
                if hit:
                    bm = _re.search(r'<a:blip\b[^>]*\br:embed="([^"]+)"', blk)
                    if bm:
                        blip = bm.group(1)
                        break
            if not blip:
                return None

            d_rels = _update_rels(
                z.read(_rel_path(draw_path)).decode("utf-8", "ignore"))
            media_target = None
            for r in d_rels:
                if r.get("Id") == blip:
                    media_target = r.get("Target")
                    break
            if not media_target:
                return None
            data = z.read(_norm_zip_path(media_target))

        im = Image.open(io.BytesIO(data))
        im.load()
        return im
    except Exception:
        return None


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


def _open_cached(fname):
    """按缓存文件名打开原图，失败返回 None"""
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


def _clipboard_png_bytes():
    """读剪贴板里的 PNG 原始字节（Excel 复制图形时有时会提供该格式）"""
    try:
        import win32clipboard
    except Exception:
        return None
    try:
        win32clipboard.OpenClipboard()
    except Exception:
        return None
    try:
        try:
            fmt = win32clipboard.RegisterClipboardFormat("PNG")
        except Exception:
            return None
        if not fmt or not win32clipboard.IsClipboardFormatAvailable(fmt):
            return None
        data = win32clipboard.GetClipboardData(fmt)
        return bytes(data) if data else None
    except Exception:
        return None
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


def cache_original_image(img, wb_name, ws_name, shape_name):
    """插入图片时把原图存进缓存，双击预览时就能用原始分辨率。

    索引同时写三套 key：
      wb|ws|shape   —— 传统 key，但工作簿"另存为"后名字会变，最不可靠
      ws|shape      —— 去掉工作簿名，改名后仍可用
      <文件名>      —— 供"替代文字"标记直接定位
    返回缓存文件名（失败返回 None）。
    """
    import time as _time
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        fname = "img_%d.png" % int(_time.time() * 1000)
        img.save(os.path.join(CACHE_DIR, fname), "PNG")
    except Exception:
        return None
    try:
        idx = _load_cache_index()
        idx["%s|%s|%s" % (wb_name, ws_name, shape_name)] = fname
        # 不含工作簿名的 key：多个工作簿有同名图时会互相覆盖，
        # 一旦发现冲突就打上歧义标记，宁可查不到也不张冠李戴
        ws_key = "%s|%s" % (ws_name, shape_name)
        prev = idx.get(ws_key)
        if prev and prev != fname:
            idx[ws_key] = AMBIGUOUS_MARK
        else:
            idx[ws_key] = fname
        idx[fname] = fname
        _save_cache_index(idx)
    except Exception:
        pass
    _prune_cache()
    return fname


def _prune_cache():
    """缓存文件超过上限时删最旧的（按文件数算，不再按索引条目数）"""
    try:
        files = [f for f in os.listdir(CACHE_DIR)
                 if f.startswith("img_") and f.endswith(".png")]
    except OSError:
        return
    if len(files) <= CACHE_MAX:
        return
    try:
        files.sort(key=lambda f: os.path.getmtime(os.path.join(CACHE_DIR, f)))
    except OSError:
        return
    dead = set(files[:len(files) - CACHE_MAX])
    for f in dead:
        try:
            os.remove(os.path.join(CACHE_DIR, f))
        except OSError:
            pass
    try:
        idx = {k: v for k, v in _load_cache_index().items() if v not in dead}
        _save_cache_index(idx)
    except Exception:
        pass


def lookup_cached_fname(wb_name, ws_name, shape_name):
    """按 工作簿|工作表|图片名 找缓存文件名；工作簿改过名时用宽松匹配兜底"""
    idx = _load_cache_index()
    for key in ("%s|%s|%s" % (wb_name, ws_name, shape_name),
                "%s|%s" % (ws_name, shape_name)):
        fn = idx.get(key)
        if fn and os.path.exists(os.path.join(CACHE_DIR, fn)):
            return fn
    # 兜底：另存为改名后工作簿名对不上，用"|工作表|图片名"后缀匹配。
    # 只有唯一命中才采用，多个候选宁可放弃也不张冠李戴。
    suffix = "|%s|%s" % (ws_name, shape_name)
    hits = {v for k, v in idx.items()
            if k.endswith(suffix) and v != AMBIGUOUS_MARK}
    if len(hits) == 1:
        fn = hits.pop()
        if os.path.exists(os.path.join(CACHE_DIR, fn)):
            return fn
    return None


def lookup_cached_image(wb_name, ws_name, shape_name):
    return _open_cached(lookup_cached_fname(wb_name, ws_name, shape_name))


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
GA_ROOT = 2
SM_CXDOUBLECLK, SM_CYDOUBLECLK = 36, 37
CF_BITMAP, CF_DIB, CF_DIBV5 = 2, 8, 17
CF_SYLK, CF_DIF = 4, 5
CF_METAFILEPICT, CF_ENHMETAFILE, CF_UNICODETEXT = 3, 14, 13

# Excel / WPS 复制"单元格内容"时特有的剪贴板格式。
# 关键：复制单元格时 Excel 会在剪贴板里额外放一份位图（供不支持表格的程序降级使用），
# 只看"有没有位图"会把单元格误判成图片。而这些表格格式只在复制表格数据时出现，
# 复制图片（微信 / 截图 / 网页 / Excel 图片对象）绝不会带，据此可可靠区分。
TABLE_FORMAT_NAMES = ("Biff12", "Biff8", "Biff5", "Biff4", "Biff3", "Biff",
                      "XML Spreadsheet", "Csv", "Rich Text Format")

# "图片文件"格式：只有复制真实图片（微信 / 截图 / 网页 / 图片文件）时才带，
# 用于兜底判断时排除"其实是图片"的情况
IMAGE_FILE_FORMAT_NAMES = ("PNG", "JFIF", "GIF", "image/png", "image/jpeg")
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
    u.RegisterClipboardFormatW.restype = wt.UINT
    u.RegisterClipboardFormatW.argtypes = [ctypes.c_wchar_p]
    u.IsClipboardFormatAvailable.restype = wt.BOOL
    u.IsClipboardFormatAvailable.argtypes = [wt.UINT]
    # 鼠标钩子给的是物理像素坐标，WindowFromPoint 在非 DPI 感知进程里按逻辑坐标
    # 解释，会错位。WindowFromPhysicalPoint（Win8+）专治这个问题。
    u.WindowFromPhysicalPoint.restype = ctypes.c_void_p
    u.WindowFromPhysicalPoint.argtypes = [POINT]
    u.GetAncestor.restype = ctypes.c_void_p
    u.GetAncestor.argtypes = [ctypes.c_void_p, ctypes.c_uint]
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
        self._suppress_time = 0.0  # 上抬吞掉的时间戳（防止标志永久卡住）
        self._table_fmt_ids = (CF_SYLK, CF_DIF)  # 表格/文本类剪贴板格式 ID
        self._imgfile_fmt_ids = ()               # 图片文件类剪贴板格式 ID

    # -- 表格格式注册（用于区分"粘单元格"和"粘图片"） ------------------------
    @staticmethod
    def _reg_ids(u, names):
        ids = []
        for name in names:
            try:
                fid = int(u.RegisterClipboardFormatW(name))
            except Exception:
                fid = 0
            if fid:
                ids.append(fid)
        return tuple(sorted(set(ids)))

    def _register_table_formats(self):
        u = self.user32
        self._table_fmt_ids = tuple(sorted(set(
            [CF_SYLK, CF_DIF] + list(self._reg_ids(u, TABLE_FORMAT_NAMES)))))
        self._imgfile_fmt_ids = self._reg_ids(u, IMAGE_FILE_FORMAT_NAMES)
        file_log(f"剪贴板格式识别已启用（表格 {len(self._table_fmt_ids)} 种 / "
                 f"图片 {len(self._imgfile_fmt_ids)} 种）")

    def _is_table_data(self, u):
        """判断剪贴板内容是否为"单元格/表格数据"（而非图片）——必须快速返回"""
        # 主判据：Excel 表格专有格式（Biff*、Csv、SYLK、DIF、RTF 等）
        if any(u.IsClipboardFormatAvailable(f) for f in self._table_fmt_ids):
            return True
        # 兜底：带文本 + 带增强图元文件 + 不含任何图片文件格式
        #      （覆盖不产生 Biff 格式的表格软件，如部分 WPS 版本）
        if (u.IsClipboardFormatAvailable(CF_UNICODETEXT)
                and u.IsClipboardFormatAvailable(CF_ENHMETAFILE)
                and not any(u.IsClipboardFormatAvailable(f)
                            for f in self._imgfile_fmt_ids)):
            return True
        return False

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

    def _is_excel_at(self, x, y):
        """点击位置是否落在 Excel 窗口上（即便 Excel 此刻不是前台窗口）。
        x, y 为鼠标钩子给出的物理像素坐标。"""
        if self._is_excel_foreground()[0]:
            return True
        try:
            hwnd = self.user32.WindowFromPhysicalPoint(POINT(int(x), int(y)))
            if not hwnd:
                return False
            root = self.user32.GetAncestor(hwnd, GA_ROOT) or hwnd
            buf = ctypes.create_unicode_buffer(64)
            self.user32.GetClassNameW(root, buf, 64)
            return buf.value in EXCEL_WND_CLASSES
        except Exception:
            return False

    # -- 低级鼠标钩子回调（必须快速返回，不做 COM） --------------------------
    def _mouse_proc(self, nCode, wParam, lParam):
        u = self.user32
        if nCode != 0 or wParam not in (WM_LBUTTONDOWN, WM_LBUTTONUP):
            return u.CallNextHookEx(None, nCode, wParam, lParam)
        try:
            import time
            now = time.monotonic()
            ms = ctypes.cast(ctypes.c_void_p(lParam),
                             ctypes.POINTER(MSLLHOOKSTRUCT)).contents
            x, y = ms.pt.x, ms.pt.y

            if wParam == WM_LBUTTONUP:
                if self._suppress_up and (now - self._suppress_time) < 1.0:
                    self._suppress_up = False
                    return 1  # 吞掉双击的第二下上抬
                self._suppress_up = False
                return u.CallNextHookEx(None, nCode, wParam, lParam)

            # ---- 左键按下 ----
            if self._suppress_up and (now - self._suppress_time) < 1.0:
                # 这是上一轮双击被吞掉的第二下。即使吞掉也要更新点击记录，
                # 否则下一次双击时 _last_click 仍是旧值，判据永远不成立
                self._last_click = now
                self._last_click_pos = (x, y)
                return 1
            self._suppress_up = False

            if self.dblclick_preview:
                dbl_time = u.GetDoubleClickTime() / 1000.0
                # 阈值取系统双击范围，并按 DPI 换算为物理像素。
                # 硬编码 6px 在 200% 缩放下只有 3 逻辑像素，手一抖就失败。
                try:
                    dpi = u.GetDpiForSystem() or 96
                except Exception:
                    dpi = 96
                k = dpi / 96.0
                cx = max(6, int(u.GetSystemMetrics(SM_CXDOUBLECLK) * k))
                cy = max(6, int(u.GetSystemMetrics(SM_CYDOUBLECLK) * k))
                dx = abs(x - self._last_click_pos[0])
                dy = abs(y - self._last_click_pos[1])
                hit = ((now - self._last_click) <= dbl_time
                       and dx <= cx and dy <= cy)
                self._last_click = now
                self._last_click_pos = (x, y)
                if hit:
                    if self._is_excel_at(x, y):
                        self._suppress_up = True
                        self._suppress_time = now
                        self.on_trigger("dblclick")
                        return 1  # 吞掉第二下按下
                    file_log("双击未触发：点击位置不在 Excel 窗口上")
        except Exception:
            pass
        return u.CallNextHookEx(None, nCode, wParam, lParam)

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
                        # 1) 前台窗口必须是 Excel / WPS 表格
                        hwnd = u.GetForegroundWindow()
                        buf = ctypes.create_unicode_buffer(64)
                        u.GetClassNameW(hwnd, buf, 64)
                        if buf.value in EXCEL_WND_CLASSES:
                            # 2) 剪贴板若是"表格/文本数据"（复制单元格、复制文字），
                            #    一律放行，交给 Excel 原生粘贴
                            if self._is_table_data(u):
                                file_log("Ctrl+V 放行：剪贴板为单元格/表格数据")
                            else:
                                # 3) 确认是真正的图片才接管
                                has_img = any(
                                    u.IsClipboardFormatAvailable(f)
                                    for f in (CF_BITMAP, CF_DIB, CF_DIBV5))
                                if has_img:
                                    import time
                                    now = time.monotonic()
                                    if now - self._last_trigger > DEBOUNCE_SEC:
                                        self._last_trigger = now
                                        file_log("Ctrl+V 拦截：剪贴板为图片，按规则插入")
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
        self._register_table_formats()
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
        try:
            _ok, _msg = ensure_excel_no_compress()
            self.log(_msg)
        except Exception as _e:
            self.log(f"检查 Excel 图片压缩设置失败：{_e}")
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
