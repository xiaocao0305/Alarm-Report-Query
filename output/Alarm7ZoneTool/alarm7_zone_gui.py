# -*- coding: utf-8 -*-
"""
Alarm Report Query —— 桌面 GUI 版
================================================
读取南盾报警主机 Alarm7 的历史报告库(Rpt)，按 ActID + ZoneID 关联
Master.mdb 的防区表，取出 AtPos（防区位置），支持导出 Excel。

功能特性：
  1. 明确区分「无记录」与「读取失败/缺月份」，失败或缺月份时禁用导出，
     超 3000 条时界面与 Excel 注明「仅显示部分记录」。
  2. 数据源可选择，显示 Master.mdb / Rpt 完整路径；
     记住上次选择，路径失效时提示重新选择。
  3. 查询条件变化后旧结果立即失效并禁用导出；只有查询成功才启用导出；
     结果区与 Excel 都显示实际日期、账号、数据源、查询时间。
  4. 每次查询开始时重新读取防区位置（一批查询读一次）。
  5. 列名「防区位置（当前配置）」，注明位置来自当前配置；
     未匹配到显示「未查到位置」，记录位置配置读取时间。
  6. 账号查询默认精确匹配，主动勾选「模糊查询」才按包含匹配；留空查全部。
  7. 支持「重新加载配置」；密码特殊字符按原文读取；报错/日志不泄露密码。
  8. Excel 按文本写入账号、防区号、名称、描述、位置，防公式注入、保留前导 0。
  9. 在 SQL 层按日期/类别/账号筛选，查询时窗口可操作，显示进度并支持取消。
  10. 统一版本号，记录代码版本、生成时间、文件名与校验值。
"""

import os
import sys
import glob
import io
import base64
import configparser
import traceback
import threading
from datetime import datetime, date, timedelta

import wx
import wx.adv

APP_NAME = "Alarm Report Query"
APP_VER = "3.1.0"
COPYRIGHT = "版权所有 2023-2030 IntelliVision Cloud。保留所有权利。"
BUILD_TIME = "2026-10-19 09:40"   # 由打包脚本在生成 exe 前写入

ALARM_TYPES = ["报警", "窃盗", "劫盗", "紧急", "求助"]
ZONE_TABLE = "CPZone"
DISPLAY_LIMIT = 3000

# 窗口标题栏图标（16/32/48 的 .ico，base64 内嵌，避免依赖外部文件）
ICON_B64 = "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"


def app_icon():
    """从内嵌数据生成窗口图标"""
    try:
        import tempfile
        path = os.path.join(tempfile.gettempdir(), "arq_app.ico")
        with io.open(path, "wb") as f:
            f.write(base64.b64decode(ICON_B64))
        icon = wx.Icon(path, wx.BITMAP_TYPE_ICO)
        return icon if icon.IsOk() else None
    except Exception:
        return None

DRIVER_CANDIDATES = [
    "Microsoft Access Driver (*.mdb, *.accdb)",
    "Microsoft Access Driver (*.mdb)",
    "Driver do Microsoft Access (*.mdb)",
    "Microsoft Access-Treiber (*.mdb)",
]


# ==================== 路径与配置 ====================
def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), "config.ini")


def load_config(path, strict=False):
    # interpolation=None：密码里的 % 等特殊字符按原文读取，不做 %(x)s 插值
    cfg = configparser.ConfigParser(interpolation=None)
    cfg["paths"] = {
        "alarm_dir": "", "master_db": "", "master_pwd": "",
        "rpt_dir": "", "rpt_pwd": "", "last_alarm_dir": "",
    }
    cfg["options"] = {"display_limit": str(DISPLAY_LIMIT)}
    if os.path.exists(path):
        if strict:
            # 语法错误/编码错误会抛出，供「重新加载配置」使用
            cfg.read(path, encoding="utf-8")
        else:
            try:
                cfg.read(path, encoding="utf-8")
            except (configparser.Error, UnicodeDecodeError):
                pass  # 启动容错：文件损坏时用默认值，不阻断启动
    return cfg


CONFIG = load_config(CONFIG_PATH)


def write_default_config():
    if os.path.exists(CONFIG_PATH):
        return
    try:
        with io.open(CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write(
                "# Alarm Report Query 配置文件\n"
                "# 路径与密码均在此定义，程序本身不内置任何密码。\n\n"
                "[paths]\n"
                "alarm_dir =\n"
                "master_db =\n"
                "# 用户数据库(Master.mdb)密码，请填写现场实际密码\n"
                "master_pwd =\n"
                "rpt_dir =\n"
                "rpt_pwd =\n"
                "# 上次选择的数据源（程序自动记录，一般不用改）\n"
                "last_alarm_dir =\n\n"
                "[options]\n"
                "display_limit = %d\n" % DISPLAY_LIMIT
            )
    except Exception:
        pass


def reload_config():
    global CONFIG
    CONFIG = load_config(CONFIG_PATH, strict=True)  # 严格模式，语法错误会抛出
    return CONFIG


# ==================== 数据源 ====================
def is_alarm_dir(d):
    try:
        return bool(d) and os.path.isfile(os.path.join(d, "Master.mdb"))
    except Exception:
        return False


def discover_installations():
    cands = []
    for key in ("alarm_dir", "last_alarm_dir"):
        v = CONFIG.get("paths", key, fallback="").strip()
        if v:
            cands.append(v)
    cfg_master = CONFIG.get("paths", "master_db", fallback="").strip()
    if cfg_master:
        cands.append(os.path.dirname(cfg_master))

    base = app_dir()
    cands += [base, os.path.join(base, "Alarm7"),
              os.path.dirname(base), os.path.join(os.path.dirname(base), "Alarm7")]
    for d in "CDEFGHIJ":
        cands.append("%s:\\Alarm7" % d)
        cands.append("%s:\\XB运维\\Alarm7" % d)

    found, seen = [], set()
    for c in cands:
        try:
            if not is_alarm_dir(c):
                continue
            ap = os.path.abspath(c)
            key = os.path.normcase(ap)
            if key in seen:
                continue
            seen.add(key)
            found.append(ap)
        except Exception:
            continue
    return found


SOURCE = {"alarm_dir": "", "master_db": "", "rpt_dir": "",
          "master_pwd": "", "rpt_pwd": ""}
INSTALLATIONS = []
_zone_cache = {}
_zone_load_time = ""


def apply_source(alarm_dir):
    """以所选 alarm_dir 为准推导 master_db 与 rpt_dir。

    关键：不再让 config.ini 里的手工 master_db/rpt_dir 覆盖所选数据源，
    否则会出现「下拉选了乙库，却读了甲库的 Master」的混库问题。
    只有当 alarm_dir 为空（纯手工模式）时才回退到手工路径。
    """
    global SOURCE
    cfg = CONFIG["paths"]
    if alarm_dir:
        master_db = os.path.join(alarm_dir, "Master.mdb")
        rpt_dir = os.path.join(alarm_dir, "Rpt")
    else:
        master_db = cfg.get("master_db", "").strip()
        rpt_dir = cfg.get("rpt_dir", "").strip()
    SOURCE = {
        "alarm_dir": alarm_dir,
        "master_db": master_db,
        "rpt_dir": rpt_dir,
        # 密码只来自外部 config.ini，源码不含任何默认密码
        "master_pwd": cfg.get("master_pwd", "").strip(),
        "rpt_pwd": cfg.get("rpt_pwd", "").strip(),
    }
    _zone_cache.clear()
    return SOURCE


def config_conflicts_with(alarm_dir):
    """检查 config.ini 手工路径是否与所选 alarm_dir 冲突（问题一）"""
    if not alarm_dir:
        return []
    cfg = CONFIG["paths"]
    conf = []
    md = cfg.get("master_db", "").strip()
    rd = cfg.get("rpt_dir", "").strip()
    if md and os.path.normcase(os.path.abspath(md)) != \
            os.path.normcase(os.path.abspath(os.path.join(alarm_dir, "Master.mdb"))):
        conf.append("master_db=%s" % md)
    if rd and os.path.normcase(os.path.abspath(rd)) != \
            os.path.normcase(os.path.abspath(os.path.join(alarm_dir, "Rpt"))):
        conf.append("rpt_dir=%s" % rd)
    return conf


def save_last_alarm_dir(d):
    """把用户选的数据源记录到 config.ini 的 last_alarm_dir，重启后保留"""
    try:
        reload_config()
        CONFIG.set("paths", "last_alarm_dir", d)
        with io.open(CONFIG_PATH, "w", encoding="utf-8") as f:
            CONFIG.write(f)
    except Exception:
        pass


def init_sources():
    """决定当前数据源，返回状态：
    'ok'           —— 已确认一个有效数据源，可查询
    'need_choice'  —— 多套数据且尚未选择，禁止查询，先要求选择
    'no_source'    —— 未发现任何有效数据源
    原则：优先用户上次选择，其次唯一安装；绝不按月份数量自动决定；
    多套且未选过时禁止查询（不偷偷选一个）。"""
    global INSTALLATIONS
    INSTALLATIONS = discover_installations()
    last = CONFIG.get("paths", "last_alarm_dir", fallback="").strip()
    cfg_dir = CONFIG.get("paths", "alarm_dir", fallback="").strip()

    if last and is_alarm_dir(last):
        apply_source(last)
        return "ok"
    if cfg_dir and is_alarm_dir(cfg_dir):
        apply_source(cfg_dir)
        return "ok"
    if len(INSTALLATIONS) == 1:
        apply_source(INSTALLATIONS[0])
        return "ok"
    if INSTALLATIONS:
        # 多套且未选过：进入「未确认」状态，禁止查询
        apply_source("")
        return "need_choice"
    apply_source("")
    return "no_source"


# ==================== 数据层 ====================
_DRIVER = None


def detect_driver():
    global _DRIVER
    if _DRIVER:
        return _DRIVER
    try:
        import pyodbc
    except ImportError:
        return None
    available = list(pyodbc.drivers())
    for cand in DRIVER_CANDIDATES:
        if cand in available:
            _DRIVER = cand
            return _DRIVER
    for d in available:
        if "Access" in d and ("mdb" in d or "accdb" in d):
            _DRIVER = d
            return _DRIVER
    return None


def _sanitize(msg, pwd):
    """把错误信息中的密码替换为 ***，避免日志/弹窗泄露口令"""
    if pwd and isinstance(msg, str):
        msg = msg.replace(pwd, "***")
    return msg


def connect(path, pwd=""):
    import pyodbc
    d = detect_driver()
    if not d:
        raise RuntimeError(
            "未找到 Access ODBC 驱动。\n"
            "本程序为 32 位，请确认系统存在 Access 驱动，"
            "或安装「Microsoft Access Database Engine 2016（32位）」。"
        )
    cs = "DRIVER={%s};DBQ=%s;PWD=%s;" % (d, path, pwd)
    try:
        return pyodbc.connect(cs, timeout=30)
    except Exception as e:
        raise RuntimeError("连接数据库失败：%s" % _sanitize(str(e), pwd))


def normalize_zone(zone):
    if zone is None:
        return ""
    s = str(zone).strip()
    return s.zfill(3) if s.isdigit() else s


def list_rpt_files():
    rpt = SOURCE["rpt_dir"]
    if not os.path.isdir(rpt):
        return []
    files = sorted(glob.glob(os.path.join(rpt, "AlmEvt*.mdb")))
    return [f for f in files if not f.lower().endswith(".ldb")]


def get_rpt_tables(conn):
    cur = conn.cursor()
    return sorted([
        r.table_name for r in cur.tables(tableType="TABLE")
        if r.table_name and not str(r.table_name).startswith("MSys")
    ])


def load_zone_map(zone_table):
    """读取防区位置映射，并记录读取时间（问题四、五）"""
    global _zone_load_time
    if not SOURCE["master_pwd"]:
        raise RuntimeError(
            "未配置用户数据库密码。\n\n"
            "请点击「重新加载配置」，或打开程序同目录的 config.ini，\n"
            "在 [paths] 段填写 master_pwd 后重新查询。")
    key = (SOURCE["master_db"], zone_table)
    if key in _zone_cache:
        return _zone_cache[key]
    conn = connect(SOURCE["master_db"], SOURCE["master_pwd"])
    try:
        cur = conn.cursor()
        cur.execute("SELECT ActID, ZoneID, AtPos FROM [%s]" % zone_table)
        mapping = {}
        for act_id, zone_id, at_pos in cur.fetchall():
            if at_pos is None or str(at_pos).strip() == "":
                continue
            mapping[(str(act_id).strip(), normalize_zone(zone_id))] = \
                str(at_pos).strip()
        _zone_load_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    finally:
        conn.close()
    _zone_cache[key] = mapping
    return mapping


def months_in_range(date_from, date_to):
    """返回日期范围内的月份列表，如 ['202610','202611']"""
    def ym(s):
        try:
            return int(s[:4]), int(s[5:7])
        except Exception:
            return 0, 0
    fy, fm = ym(date_from)
    ty, tm = ym(date_to)
    out = []
    if not (fy and ty):
        return out
    y, m = fy, fm
    while (y, m) <= (ty, tm):
        out.append("%04d%02d" % (y, m))
        m += 1
        if m > 12:
            m = 1
            y += 1
    return out


def month_files_for_range(date_from, date_to):
    files = list_rpt_files()
    if not date_from or not date_to:
        return files
    wanted = set(months_in_range(date_from, date_to))
    out = []
    for f in files:
        base = os.path.basename(f)
        if base.startswith("AlmEvt") and base.endswith(".mdb"):
            key = base[6:12]
            if key.isdigit() and key in wanted:
                out.append(f)
    return out


def missing_months_for_range(date_from, date_to):
    """返回日期范围内缺失的月份文件（问题一）"""
    wanted = months_in_range(date_from, date_to)
    if not wanted:
        return []
    present = set()
    for f in list_rpt_files():
        base = os.path.basename(f)
        if base.startswith("AlmEvt") and base.endswith(".mdb"):
            k = base[6:12]
            if k.isdigit():
                present.add(k)
    return [w for w in wanted if w not in present]


def build_sql(table, evt_types, user, fuzzy, date_from, date_to):
    """在 SQL 层按日期、类别、账号筛选（问题九）"""
    wheres, params = [], []
    if date_from:
        wheres.append("Evtdatetime >= ?")
        params.append(date_from + " 00:00:00")
    if date_to:
        wheres.append("Evtdatetime <= ?")
        params.append(date_to + " 23:59:59")
    if evt_types:
        wheres.append("EvtType IN (%s)" % ",".join("?" * len(evt_types)))
        params.extend(evt_types)
    if user:
        if fuzzy:
            wheres.append("ActID LIKE ?")
            params.append("%" + user + "%")
        else:
            wheres.append("ActID = ?")   # 完整账号精确匹配
            params.append(user)
    sql = "SELECT Evtdatetime, ActID, UsrName, EvtType, EvtMemo, ZoneID " \
          "FROM [%s]" % table
    if wheres:
        sql += " WHERE " + " AND ".join(wheres)
    return sql, params


def query_file(rpt_file, evt_types, user, fuzzy, limit, date_from, date_to,
               cancel_event, zone_map):
    """查询单个月份库，返回 (rows, truncated, errors, cancelled)"""
    results, truncated, errors = [], False, []
    conn = connect(rpt_file, SOURCE["rpt_pwd"])
    try:
        for table in get_rpt_tables(conn):
            if not table.upper().startswith("D"):
                continue
            try:
                sql, params = build_sql(table, evt_types, user, fuzzy,
                                        date_from, date_to)
                cur = conn.cursor()
                cur.execute(sql, params)
                while True:
                    if cancel_event and cancel_event.is_set():
                        return results, truncated, errors, True
                    batch = cur.fetchmany(1000)
                    if not batch:
                        break
                    for dt, act_id, usr_name, evt_type, memo, zone_id in batch:
                        dt_str = "" if dt is None else str(dt)
                        # 日期做二次校验（SQL 已筛，此处兜底防格式差异）
                        if date_from and dt_str[:10] < date_from:
                            continue
                        if date_to and dt_str[:10] > date_to:
                            continue
                        act_id = "" if act_id is None else str(act_id).strip()
                        evt_type = "" if evt_type is None else str(evt_type).strip()
                        z = normalize_zone(zone_id)
                        results.append({
                            "时间": dt_str,
                            "用户账号": act_id,
                            "用户名称": "" if usr_name is None else str(usr_name),
                            "事件类型": evt_type,
                            "事件描述": "" if memo is None else str(memo),
                            "防区号": z,
                            "防区位置": zone_map.get((act_id, z), ""),
                        })
                        if len(results) >= limit:
                            truncated = True
                            break
                    if truncated:
                        break
            except Exception as e:
                errors.append("%s 表 %s 读取失败" % (
                    os.path.basename(rpt_file), table))
            if truncated:
                break
    finally:
        conn.close()
    return results, truncated, errors, False


def run_query(date_from, date_to, user, fuzzy, cancel_event, progress):
    """执行一次完整查询，返回结果字典（问题一、三、四、九）"""
    _zone_cache.clear()                       # 每次查询重新读防区位置（问题四）
    zone_map = load_zone_map(ZONE_TABLE)
    zone_time = _zone_load_time

    files = month_files_for_range(date_from, date_to)
    missing = missing_months_for_range(date_from, date_to)

    rows, truncated, errors = [], False, []
    total = len(files) or 1
    for i, f in enumerate(files):
        if cancel_event and cancel_event.is_set():
            return {"cancelled": True, "rows": [], "truncated": False,
                    "errors": errors, "missing_months": missing,
                    "zone_time": zone_time}
        progress(int(i * 100 / total), "正在读取 %s（%d/%d）……"
                 % (os.path.basename(f), i + 1, total))
        part, tr, errs, cancelled = query_file(
            f, ALARM_TYPES, user, fuzzy, DISPLAY_LIMIT - len(rows),
            date_from, date_to, cancel_event, zone_map)
        rows.extend(part)
        errors.extend(errs)
        if tr:
            truncated = True
        if cancelled:
            return {"cancelled": True, "rows": rows, "truncated": truncated,
                    "errors": errors, "missing_months": missing,
                    "zone_time": zone_time}
        if truncated:
            break

    rows.sort(key=lambda r: r["时间"])
    return {"cancelled": False, "rows": rows, "truncated": truncated,
            "errors": errors, "missing_months": missing,
            "zone_time": zone_time}


# ==================== Excel 导出 ====================
XL_HEADER_FILL = "1565C0"
XL_HEADER_FONT = "FFFFFF"
XL_GREEN_FILL = "EAF7EE"
XL_ORANGE_FILL = "FDF3E7"
XL_COL_WIDTHS = [20, 12, 14, 10, 32, 8, 20]


def _write_text(ws, row, col, value, font, fill=None, align=None):
    """按文本写入单元格，防 = 开头被当公式、保留前导 0（问题八）"""
    cell = ws.cell(row=row, column=col, value="" if value is None else str(value))
    cell.data_type = "s"          # 强制字符串类型
    cell.font = font
    if fill is not None:
        cell.fill = fill
    if align is not None:
        cell.alignment = align
    return cell


def export_excel(rows, meta, path):
    """导出 Excel：结果表 + 查询说明表（问题三、五、八）"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "查询结果"

    h_fill = PatternFill("solid", fgColor=XL_HEADER_FILL)
    h_font = Font(name="微软雅黑", size=10, bold=True, color=XL_HEADER_FONT)
    g_fill = PatternFill("solid", fgColor=XL_GREEN_FILL)
    o_fill = PatternFill("solid", fgColor=XL_ORANGE_FILL)
    body_font = Font(name="微软雅黑", size=10)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    # 表头
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = h_fill
        cell.font = h_font
        cell.alignment = center

    # 数据行（全部按文本写入；用 FIELDS 取数据、HEADERS 显示）
    for i, r in enumerate(rows, 2):
        fill = g_fill if r["防区位置"] else o_fill
        for c, f in enumerate(FIELDS, 1):
            v = r[f]
            if f == "防区位置" and not v:
                v = "未查到位置"
            _write_text(ws, i, c, v, body_font, fill,
                        center if f == "防区号" else left)

    for c, w in enumerate(XL_COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # 查询说明 sheet
    ws2 = wb.create_sheet("查询说明")
    info = [
        ("查询日期范围", "%s ～ %s" % (meta["date_from"], meta["date_to"])),
        ("用户账号条件", meta.get("account_text", "")),
        ("查询方式", meta.get("match_text", "")),
        ("数据来源(Master)", meta.get("master_db", "")),
        ("报告目录(Rpt)", meta.get("rpt_dir", "")),
        ("查询时间", meta.get("query_time", "")),
        ("防区位置配置读取时间", meta.get("zone_time", "")),
        ("结果说明", meta.get("result_note", "")),
        ("重要提示", "「防区位置」来自当前配置，不保证与报警发生时一致。"),
    ]
    for i, (k, v) in enumerate(info, 1):
        _write_text(ws2, i, 1, k, Font(name="微软雅黑", size=10, bold=True))
        _write_text(ws2, i, 2, v, Font(name="微软雅黑", size=10))
    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 60

    wb.save(path)


# ==================== GUI ====================
HEADERS = ["时间", "用户账号", "用户名称", "事件类型", "事件描述",
           "防区号", "防区位置（当前配置）"]
FIELDS = ["时间", "用户账号", "用户名称", "事件类型", "事件描述",
          "防区号", "防区位置"]
COL_WIDTHS = [150, 95, 100, 80, 210, 65, 170]

GREEN = wx.Colour(234, 247, 238)
ORANGE = wx.Colour(253, 243, 231)


class MainFrame(wx.Frame):
    def __init__(self, source_status="ok"):
        style = wx.DEFAULT_FRAME_STYLE | wx.MINIMIZE_BOX | wx.MAXIMIZE_BOX
        super(MainFrame, self).__init__(
            None, title="%s v%s" % (APP_NAME, APP_VER),
            size=(1100, 700), style=style)
        self.rows = []
        self.state = None          # 查询结果状态
        self._task_id = 0          # 任务编号（问题二、三）
        self._running = False      # 是否有查询任务在运行
        self._cancel = None        # 当前任务的取消标志
        self._worker = None
        self._source_status = source_status
        self._build_ui()
        self._refresh_source_combo()
        self._invalidate_results()
        self._apply_source_status()
        icon = app_icon()
        if icon:
            self.SetIcon(icon)
        self.Centre()

    # ---------- UI 构建 ----------
    def _build_ui(self):
        panel = wx.Panel(self)
        top = wx.BoxSizer(wx.VERTICAL)

        # 第一行：数据源 + 重载配置
        r0 = wx.BoxSizer(wx.HORIZONTAL)
        r0.Add(wx.StaticText(panel, label="数据源："), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.cmb_src = wx.ComboBox(panel, style=wx.CB_READONLY, size=(420, -1))
        self.cmb_src.Bind(wx.EVT_COMBOBOX, self.on_switch_src)
        r0.Add(self.cmb_src, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)
        btn_reload = wx.Button(panel, label="重新加载配置")
        btn_reload.Bind(wx.EVT_BUTTON, self.on_reload_config)
        r0.Add(btn_reload, 0, wx.ALIGN_CENTER_VERTICAL)
        top.Add(r0, 0, wx.EXPAND | wx.ALL, 8)

        # 实际读取路径（问题一：界面显示 Master/Rpt 完整路径）
        self.lbl_paths = wx.StaticText(panel, label="", style=wx.ST_ELLIPSIZE_END)
        self.lbl_paths.SetForegroundColour(wx.Colour(60, 60, 60))
        top.Add(self.lbl_paths, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

        # 第二行：日期范围 + 快捷按钮
        r1 = wx.BoxSizer(wx.HORIZONTAL)
        r1.Add(wx.StaticText(panel, label="开始日期："), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.dp_start = wx.adv.DatePickerCtrl(
            panel, style=wx.adv.DP_DROPDOWN | wx.adv.DP_SHOWCENTURY)
        self.dp_start.Bind(wx.adv.EVT_DATE_CHANGED, self._on_cond_change)
        r1.Add(self.dp_start, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        r1.Add(wx.StaticText(panel, label="至"), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        self.dp_end = wx.adv.DatePickerCtrl(
            panel, style=wx.adv.DP_DROPDOWN | wx.adv.DP_SHOWCENTURY)
        self.dp_end.Bind(wx.adv.EVT_DATE_CHANGED, self._on_cond_change)
        r1.Add(self.dp_end, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 20)
        r1.Add(wx.StaticText(panel, label="快捷："), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        for label, handler in [("当日报告", self.on_today),
                               ("本周报告", self.on_week),
                               ("本月报告", self.on_month)]:
            b = wx.Button(panel, label=label)
            b.Bind(wx.EVT_BUTTON, handler)
            r1.Add(b, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        top.Add(r1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)

        # 默认日期：本月初到今天
        now = wx.DateTime.Now()
        self.dp_start.SetValue(wx.DateTime(1, now.GetMonth(), now.GetYear()))
        self.dp_end.SetValue(now)

        # 第三行：账号 + 查询方式 + 操作
        r2 = wx.BoxSizer(wx.HORIZONTAL)
        r2.Add(wx.StaticText(panel, label="用户账号："), 0,
               wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.txt_user = wx.TextCtrl(panel, size=(130, -1))
        self.txt_user.Bind(wx.EVT_TEXT, self._on_cond_change)
        r2.Add(self.txt_user, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 10)
        self.chk_fuzzy = wx.CheckBox(panel, label="模糊查询")
        self.chk_fuzzy.Bind(wx.EVT_CHECKBOX, self._on_cond_change)
        r2.Add(self.chk_fuzzy, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 16)

        self.btn_query = wx.Button(panel, label="开始查询")
        self.btn_query.Bind(wx.EVT_BUTTON, self.on_query)
        r2.Add(self.btn_query, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        self.btn_cancel = wx.Button(panel, label="取消")
        self.btn_cancel.Bind(wx.EVT_BUTTON, self.on_cancel)
        self.btn_cancel.Disable()
        r2.Add(self.btn_cancel, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        self.btn_export = wx.Button(panel, label="导出 Excel")
        self.btn_export.Bind(wx.EVT_BUTTON, self.on_export)
        self.btn_export.Disable()
        r2.Add(self.btn_export, 0, wx.ALIGN_CENTER_VERTICAL)
        top.Add(r2, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)

        # 进度条（查询时显示）
        self.gauge = wx.Gauge(panel, range=100, size=(-1, 8))
        self.gauge.Hide()
        top.Add(self.gauge, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        # 结果表格
        self.list = wx.ListCtrl(panel, style=wx.LC_REPORT)
        for i, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS)):
            self.list.InsertColumn(i, h, width=w)
        top.Add(self.list, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)

        # 结果元信息（问题三）
        self.lbl_meta = wx.StaticText(panel, label="", style=wx.ST_ELLIPSIZE_END)
        self.lbl_meta.SetForegroundColour(wx.Colour(90, 90, 90))
        top.Add(self.lbl_meta, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)

        # 固定可见的位置说明（问题五）
        lbl_note = wx.StaticText(
            panel, label="说明：「防区位置」来自当前配置，不保证与报警发生时一致；"
                         "未查到位置不代表该防区未配置。")
        lbl_note.SetForegroundColour(wx.Colour(180, 80, 0))
        top.Add(lbl_note, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 2)

        # 版权信息
        lbl_copy = wx.StaticText(panel, label=COPYRIGHT)
        lbl_copy.SetForegroundColour(wx.Colour(150, 150, 150))
        top.Add(lbl_copy, 0, wx.ALIGN_CENTER | wx.BOTTOM, 4)

        self.sb = self.CreateStatusBar(2)
        self.sb.SetStatusWidths([-3, -1])
        self.sb.SetStatusText("就绪。请确认数据源后查询。", 0)
        self.sb.SetStatusText("数据源：%s" % SOURCE["alarm_dir"], 1)

        panel.SetSizer(top)
        panel.Layout()

    # ---------- 数据源 ----------
    def _update_paths_label(self):
        if SOURCE["alarm_dir"]:
            self.lbl_paths.SetLabelText(
                "Master 库：%s　｜　报告目录：%s"
                % (SOURCE["master_db"], SOURCE["rpt_dir"]))
        else:
            self.lbl_paths.SetLabelText("尚未选择数据源。")

    def _refresh_source_combo(self):
        self.cmb_src.Clear()
        for d in INSTALLATIONS:
            self.cmb_src.Append(d, d)
        cur = SOURCE["alarm_dir"]
        if cur and is_alarm_dir(cur):
            if cur not in INSTALLATIONS:
                self.cmb_src.Append(cur, cur)
            self.cmb_src.SetValue(cur)
        self._update_paths_label()

    def _apply_source_status(self):
        """根据数据源确认状态决定是否允许查询（问题一）"""
        self._update_paths_label()
        if self._source_status == "ok":
            self.sb.SetStatusText("数据源：%s" % SOURCE["alarm_dir"], 1)
            return
        self.btn_query.Disable()
        self.btn_export.Disable()
        if self._source_status == "need_choice":
            self.sb.SetStatusText("请先选择数据源", 1)
            self.sb.SetStatusText(
                "本机有多套 Alarm7 数据，请在上方选择要查询的那一套后再查询。", 0)
        else:
            self.sb.SetStatusText("未找到数据源", 1)
            self.sb.SetStatusText(
                "未找到有效的 Alarm7 数据，请点击「重新加载配置」检查路径。", 0)

    def on_switch_src(self, event):
        idx = self.cmb_src.GetSelection()
        if idx == wx.NOT_FOUND:
            return
        d = self.cmb_src.GetClientData(idx)
        if not (d and is_alarm_dir(d)):
            return
        d = os.path.abspath(d)
        # 检查 config 手工路径是否与所选数据源冲突（问题一）
        conflicts = config_conflicts_with(d)
        apply_source(d)
        save_last_alarm_dir(d)
        self._source_status = "ok"
        self._invalidate_results()
        self.btn_query.Enable()
        self._update_paths_label()
        self.sb.SetStatusText("数据源：%s" % d, 1)
        if conflicts:
            msg = ("已选择数据源：\n%s\n\n"
                   "注意：config.ini 中的手工路径与所选数据源不一致，"
                   "已忽略手工路径，使用所选数据源：\n%s"
                   % (d, "\n".join(conflicts)))
            self.sb.SetStatusText("已切换数据源（有配置冲突，见提示）", 0)
            wx.MessageBox(msg, APP_NAME, wx.OK | wx.ICON_WARNING)
        else:
            self.sb.SetStatusText("已切换数据源：%s" % d, 0)

    def on_reload_config(self, event):
        try:
            reload_config()
        except (configparser.Error, UnicodeDecodeError) as e:
            self._invalidate_results()
            self.btn_query.Disable()
            self.btn_export.Disable()
            self.sb.SetStatusText("配置读取失败", 0)
            wx.MessageBox(
                "config.ini 读取失败，未使用新配置。\n\n原因：%s\n\n"
                "请修正 config.ini 格式后再点「重新加载配置」。"
                % _sanitize(str(e), CONFIG.get("paths", "master_pwd",
                                               fallback="")),
                APP_NAME, wx.OK | wx.ICON_ERROR)
            return
        self._source_status = init_sources()
        self._refresh_source_combo()
        self._invalidate_results()
        self._apply_source_status()
        if self._source_status == "ok":
            self.sb.SetStatusText("配置已重新加载，数据源：%s"
                                  % SOURCE["alarm_dir"], 0)
            wx.MessageBox("配置已重新加载。\n数据源：%s" % SOURCE["alarm_dir"],
                          APP_NAME, wx.OK | wx.ICON_INFORMATION)

    # ---------- 快捷日期 ----------
    def _wxdt(self, d):
        return wx.DateTime(d.day, d.month - 1, d.year)

    def _set_dates_and_query(self, d_start, d_end):
        self.dp_start.SetValue(d_start)
        self.dp_end.SetValue(d_end)
        self.on_query(None)

    def on_today(self, event):
        t = date.today()
        self._set_dates_and_query(self._wxdt(t), self._wxdt(t))

    def on_week(self, event):
        t = date.today()
        monday = t - timedelta(days=t.weekday())
        self._set_dates_and_query(self._wxdt(monday), self._wxdt(t))

    def on_month(self, event):
        t = date.today()
        self._set_dates_and_query(self._wxdt(t.replace(day=1)), self._wxdt(t))

    # ---------- 结果状态管理（问题三）----------
    def _on_cond_change(self, event):
        """日期/账号/查询方式/数据源变化后，旧结果立即失效；
        若查询进行中，则作废该任务（取消它，结果不再显示）。"""
        self._invalidate_results()
        if self._running and self._cancel:
            self._cancel.set()

    def _invalidate_results(self):
        self.state = None
        self.rows = []
        self.list.DeleteAllItems()
        self.lbl_meta.SetLabelText("")
        self.btn_export.Disable()

    # ---------- 查询 ----------
    def on_query(self, event):
        if self._running:
            self.sb.SetStatusText("已有查询进行中，请先取消或等待完成。", 0)
            return
        if self._source_status != "ok":
            wx.MessageBox(
                "尚未确认数据源，无法查询。\n"
                "请在上方「数据源」下拉框选择要查询的 Alarm7 数据。",
                APP_NAME, wx.OK | wx.ICON_WARNING)
            return
        sdt = self.dp_start.GetValue()
        edt = self.dp_end.GetValue()
        if not sdt.IsValid() or not edt.IsValid():
            wx.MessageBox("请选择有效的开始和结束日期。", APP_NAME,
                          wx.OK | wx.ICON_WARNING)
            return
        date_from = "%04d-%02d-%02d" % (sdt.GetYear(), sdt.GetMonth() + 1,
                                        sdt.GetDay())
        date_to = "%04d-%02d-%02d" % (edt.GetYear(), edt.GetMonth() + 1,
                                      edt.GetDay())
        if date_from > date_to:
            wx.MessageBox("开始日期不能晚于结束日期。", APP_NAME,
                          wx.OK | wx.ICON_WARNING)
            return
        user = self.txt_user.GetValue().strip() or None
        fuzzy = self.chk_fuzzy.GetValue()

        # 任务快照：整个任务用这份快照，不从输入框/全局 SOURCE 再取值（问题二）
        snapshot = {
            "date_from": date_from,
            "date_to": date_to,
            "user": user,
            "fuzzy": fuzzy,
            "alarm_dir": SOURCE["alarm_dir"],
            "master_db": SOURCE["master_db"],
            "rpt_dir": SOURCE["rpt_dir"],
            "account_text": user or "（全部账号）",
            "match_text": "模糊查询" if fuzzy else "完整账号精确匹配",
            "query_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        self._invalidate_results()
        self._task_id += 1
        task_id = self._task_id
        cancel = threading.Event()
        self._cancel = cancel
        self._running = True
        self.btn_query.Disable()
        self.btn_cancel.Enable()
        self.gauge.SetValue(0)
        self.gauge.Show()
        self.sb.SetStatusText("正在查询……", 0)

        self._worker = threading.Thread(
            target=self._query_thread,
            args=(task_id, snapshot, cancel),
            daemon=True)
        self._worker.start()

    def _query_thread(self, task_id, snapshot, cancel):
        try:
            result = run_query(
                snapshot["date_from"], snapshot["date_to"],
                snapshot["user"], snapshot["fuzzy"], cancel,
                progress=lambda p, m: wx.CallAfter(
                    self._on_progress, task_id, cancel, p, m))
            wx.CallAfter(self._on_query_done, task_id, snapshot, cancel,
                         result, None)
        except Exception as e:
            wx.CallAfter(self._on_query_done, task_id, snapshot, cancel,
                         None, _sanitize(str(e), SOURCE["master_pwd"]))

    def _on_progress(self, task_id, cancel, pct, msg):
        if cancel is not self._cancel:
            return  # 旧任务，丢弃进度
        if pct is not None:
            self.gauge.SetValue(pct)
        self.sb.SetStatusText(msg, 0)

    def _on_query_done(self, task_id, snapshot, cancel, result, err):
        if cancel is not self._cancel:
            return  # 旧任务晚返回，不能改动新任务（问题三）
        self._running = False
        self.gauge.Hide()
        self.btn_cancel.Disable()
        self.btn_query.Enable()
        if err:
            self.sb.SetStatusText("查询失败", 0)
            wx.MessageBox("查询失败：\n%s" % err, APP_NAME,
                          wx.OK | wx.ICON_ERROR)
            return
        if result.get("cancelled"):
            self.sb.SetStatusText("查询已取消", 0)
            return
        self._apply_result(result, snapshot)

    def on_cancel(self, event):
        if self._cancel:
            self._cancel.set()
        self.sb.SetStatusText("正在取消……", 0)

    def _apply_result(self, result, snapshot):
        rows = result["rows"]
        errors = result["errors"]
        missing = result["missing_months"]
        truncated = result["truncated"]
        zone_time = result["zone_time"]
        self.rows = rows
        self._fill_table(rows)

        hit = sum(1 for r in rows if r["防区位置"])
        notes = []
        if errors:
            notes.append("有 %d 个文件/表读取失败" % len(errors))
        if missing:
            notes.append("缺 %d 个月份文件（%s）" % (
                len(missing), "、".join(missing)))
        if truncated:
            notes.append("仅显示前 %d 条（部分结果）" % DISPLAY_LIMIT)

        if rows:
            msg = "共 %d 条｜命中位置 %d 条｜未查到位置 %d 条" % (
                len(rows), hit, len(rows) - hit)
        else:
            if errors or missing:
                msg = "无完整结果（读取失败或缺月份，不能确定没有记录）"
            else:
                msg = "没有符合条件的记录"
        if notes:
            msg += "｜" + "；".join(notes)
        self.sb.SetStatusText(msg, 0)
        self.sb.SetStatusText("数据源：%s" % snapshot["alarm_dir"], 1)

        # 元信息全部来自本批快照，不用当前输入框（问题二）
        result_note = "；".join(notes) if notes else "完整结果"
        self.lbl_meta.SetLabelText(
            "查询时间 %s ｜ 日期 %s ～ %s ｜ 账号 %s（%s）｜ 数据源 %s ｜ "
            "位置配置读取于 %s ｜ %s" % (
                snapshot["query_time"], snapshot["date_from"],
                snapshot["date_to"], snapshot["account_text"],
                snapshot["match_text"], snapshot["alarm_dir"],
                zone_time or "未知", result_note))

        # 导出控制：仅查询成功且无读取失败/缺月份才启用
        exportable = bool(rows) and not errors and not missing
        self.btn_export.Enable(exportable)
        self.state = {
            "date_from": snapshot["date_from"], "date_to": snapshot["date_to"],
            "account_text": snapshot["account_text"],
            "match_text": snapshot["match_text"],
            "master_db": snapshot["master_db"], "rpt_dir": snapshot["rpt_dir"],
            "query_time": snapshot["query_time"],
            "zone_time": zone_time or "", "result_note": result_note,
        }

        if errors or missing:
            detail = []
            if errors:
                detail.append("读取失败：\n" + "\n".join(
                    "  · " + e for e in errors))
            if missing:
                detail.append("缺少月份文件：%s（这些月份的数据无法确认）"
                              % "、".join(missing))
            detail.append("\n因此本批结果可能不完整，已禁用导出。")
            wx.MessageBox("\n".join(detail), APP_NAME,
                          wx.OK | wx.ICON_WARNING)

    def _fill_table(self, rows):
        self.list.Freeze()
        self.list.DeleteAllItems()
        for idx, r in enumerate(rows):
            self.list.InsertItem(idx, r["时间"])
            self.list.SetItem(idx, 1, r["用户账号"])
            self.list.SetItem(idx, 2, r["用户名称"])
            self.list.SetItem(idx, 3, r["事件类型"])
            self.list.SetItem(idx, 4, r["事件描述"])
            self.list.SetItem(idx, 5, r["防区号"])
            self.list.SetItem(idx, 6, r["防区位置"] or "未查到位置")
            if r["防区位置"]:
                self.list.SetItemBackgroundColour(idx, GREEN)
            else:
                self.list.SetItemBackgroundColour(idx, ORANGE)
        self.list.Thaw()

    # ---------- 导出 ----------
    def on_export(self, event):
        if not self.state or not self.rows:
            wx.MessageBox("当前没有可导出的结果。\n"
                          "请先成功查询（读取失败或缺月份时不能导出）。",
                          APP_NAME, wx.OK | wx.ICON_INFORMATION)
            return
        dlg = wx.FileDialog(
            self, "保存 Excel 文件", wildcard="Excel 文件 (*.xlsx)|*.xlsx",
            defaultFile="AlarmReportQuery_%s.xlsx"
                        % datetime.now().strftime("%Y%m%d_%H%M%S"),
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() != wx.ID_OK:
            dlg.Destroy()
            return
        path = dlg.GetPath()
        dlg.Destroy()
        try:
            export_excel(self.rows, self.state, path)
            wx.MessageBox("已导出 %d 条记录到：\n%s" % (len(self.rows), path),
                          APP_NAME, wx.OK | wx.ICON_INFORMATION)
        except ImportError:
            wx.MessageBox("缺少 openpyxl 模块，无法导出 Excel。\n"
                          "请执行：pip install openpyxl",
                          APP_NAME, wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox("导出失败：\n%s" % e, APP_NAME,
                          wx.OK | wx.ICON_ERROR)


def main():
    write_default_config()
    status = init_sources()

    app = wx.App(False)
    frame = MainFrame(status)
    frame.Show(True)

    if not detect_driver():
        wx.MessageBox(
            "未找到 Access ODBC 驱动。\n\n"
            "本程序为 32 位版本，正常情况下使用 Windows 自带的 "
            "Access 驱动即可。\n若仍无法读取数据库，请安装 "
            "「Microsoft Access Database Engine 2016（32位）」。",
            APP_NAME, wx.OK | wx.ICON_WARNING)

    app.MainLoop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        msg = traceback.format_exc()
        try:
            with io.open(os.path.join(app_dir(), "error.log"),
                         "w", encoding="utf-8") as f:
                f.write(msg)
        except Exception:
            pass
        try:
            print(msg)
        except Exception:
            pass
