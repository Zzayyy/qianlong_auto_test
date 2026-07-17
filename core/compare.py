# -*- coding: utf-8 -*-
"""结果比对核心逻辑
========================
职责：
  - 扫描两个文件夹，按文件名（/相对路径）匹配同名文件
  - 对同名文件做内容比对：
      * .xls / .xlsx / .xlsm 等表格文件 -> 逐工作表、逐行内容比对
      * 其它文件 -> 文件大小 + MD5 哈希比对
  - 先实现 xls/xlsx 的内容比对（行级差异），其余类型做二进制一致性比对
  - 生成可阅读的纯文本总结报告，也可导出为文件

说明：
  - 表格读取采用「多引擎 + HTML 兜底」策略（见 _read_excel_sheets）：
      * 优先使用 calamine 引擎（pandas engine="calamine"），它能同时读取
        老版 .xls（BIFF8）与新版 .xlsx；
      * 若未安装 python-calamine，则对 .xlsx 回退 openpyxl、对 .xls 回退 xlrd；
      * 交易终端「另存为 .xls」实为 HTML 表格（这是「读取失败：cannot detect
        file format」最常见的根因）→ 用 pd.read_html 兜底解析。
    任何引擎失败都会给出清晰可读的报错，便于定位缺失的依赖或异常的文件格式。
  - 行级比对支持两种模式：
      * 忽略行顺序（默认）：把每个工作表当作「行的多重集合」比较，
        直观呈现「新增 N 行 / 减少 N 行」，对重新导出/重排序的表格更友好。
      * 按位置对齐：逐行逐格比较，定位到具体的 (行,列) 差异。
"""

import os
import re
import hashlib
from collections import Counter
from datetime import datetime

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:  # pragma: no cover
    HAS_PANDAS = False


# 支持的表格扩展名（大小写不敏感）
EXCEL_EXTS = (".xls", ".xlsx", ".xlsm", ".xltx", ".xltm")

# 报告 / 明细中每行最多展示的差异行数（避免报告过长）
MAX_DETAIL_ROWS = 200
# 单元格差异（按位置对齐模式）最多展示数量
MAX_CELL_DIFFS = 500


# ====================== 文件扫描 ======================
def scan_folder(folder, recursive=False):
    """扫描文件夹，返回 {匹配键: 完整路径}。

    - recursive=False：以文件名为键（顶层文件）。
    - recursive=True：以相对根目录的相对路径为键（子目录结构需一致才能匹配）。
    """
    result = {}
    if not folder or not os.path.isdir(folder):
        return result
    if recursive:
        for root, _dirs, files in os.walk(folder):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, folder)
                result[rel.replace("/", os.sep)] = full
    else:
        for f in sorted(os.listdir(folder)):
            full = os.path.join(folder, f)
            if os.path.isfile(full):
                result[f] = full
    return result


# ====================== 通用工具 ======================
def _md5(path, chunk=1 << 20):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def _norm_cell(v):
    """把单元格值规整为可显示的字符串。"""
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:  # NaN
            return ""
        if v == int(v):
            return str(int(v))
        return repr(v)
    if isinstance(v, int):
        return str(v)
    return str(v).strip()


_NUM_RE = re.compile(r"^[+-]?[\d,]+(\.\d+)?$")


def _key_cell(s):
    """用于相等判断的「归一键」：数值按数值比，文本按原串比。

    这样 '1,000.00' / '1000' / '1000.00' 会视为相等，
    对金额、数量等列更鲁棒；合约代码等非数值文本则按原串比较。
    """
    t = s.replace(",", "").replace("，", "")
    try:
        return ("N", float(t))
    except ValueError:
        return ("S", s)


# ====================== Excel 读取 ======================
def _engine_available(engine):
    """探测某个 pandas 读取引擎是否真正可用（对应依赖已安装）。

    优先依据 pandas 的引擎注册表判断（跨版本兼容：旧版 import 'calamine'，
    新版 import 'python_calamine'），再回退到直接 import 校验，避免误判。
    例如本机安装的是 python-calamine 0.8.x（模块名 python_calamine），
    若只检查 `import calamine` 会误以为缺失而弃用最优引擎。
    """
    if not HAS_PANDAS:
        return False
    # 1) 通过 pandas 引擎注册表判断（最可靠）
    try:
        reg = None
        try:
            reg = pd.io.excel.ExcelFile._engines
        except Exception:
            reg = pd.io.excel._engines  # 部分 pandas 版本位置不同
        if reg is not None and engine in reg:
            return True
    except Exception:
        pass
    # 2) 回退：直接 import 对应依赖
    try:
        if engine == "calamine":
            try:
                import python_calamine  # noqa: F401  (新版包名)
            except ImportError:
                import calamine  # noqa: F401  (旧版包名)
        elif engine == "openpyxl":
            import openpyxl  # noqa: F401
        elif engine == "xlrd":
            import xlrd  # noqa: F401
        else:
            return False
        return True
    except Exception:
        return False


def _xlrd_cell_to_value(cell):
    """把 xlrd 单元格转为与原生读取一致的值（交给 _norm_cell 统一规整）。"""
    import xlrd
    ctype = cell.ctype
    if ctype == xlrd.XL_CELL_EMPTY:
        return ""
    if ctype == xlrd.XL_CELL_NUMBER:
        return cell.value
    if ctype == xlrd.XL_CELL_DATE:
        # 日期转为稳定字符串，便于与对照端（可能以文本/数值存储）一致比对
        try:
            return xlrd.xldate.xldate_as_datetime(cell.value, 0).strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        except Exception:
            return cell.value
    if ctype == xlrd.XL_CELL_BOOLEAN:
        return "TRUE" if cell.value else "FALSE"
    if ctype == xlrd.XL_CELL_ERROR:
        try:
            return xlrd.error_text_from_code.get(cell.value, str(cell.value))
        except Exception:
            return str(cell.value)
    return cell.value  # 文本


def _read_with_xlrd(path):
    """用 xlrd 读取 .xls，并显式指定中文编码。

    很多交易终端导出的老版 .xls 没有 CODEPAGE 记录，xlrd 默认会回退到
    iso-8859-1 导致中文乱码（并刷出大量 'No CODEPAGE record...' 警告）。
    这里依次尝试 gb18030/gbk 中文编码覆盖，最后再回退到 xlrd 默认行为。
    """
    import xlrd
    last_err = None
    for enc in ("gb18030", "gbk", None):
        try:
            if enc is None:
                book = xlrd.open_workbook(path)
            else:
                book = xlrd.open_workbook(path, encoding_override=enc)
            sheets = {}
            for sname in book.sheet_names():
                ws = book.sheet_by_name(sname)
                data = [
                    [_xlrd_cell_to_value(ws.cell(r, c))
                     for c in range(ws.ncols)]
                    for r in range(ws.nrows)
                ]
                sheets[sname] = pd.DataFrame(data, dtype=object)
            return sheets
        except Exception as e:
            last_err = e
    raise last_err


def _read_excel_sheets(path):
    """读取 Excel 全部工作表，返回 {sheet名: DataFrame(header=None, dtype=object)}。

    采用「多引擎 + HTML 兜底」策略，尽量兼容各种怪异的 .xls/.xlsx 文件：

      * 真正的 BIFF8 老版 .xls / 现代 .xlsx(.xlsm) -> calamine（首选，速度快、兼容好）；
      * 缺失 calamine 时 -> 对 .xlsx 回退 openpyxl，对 .xls 回退 xlrd；
      * 交易终端「另存为 .xls」实为 HTML 表格 -> 用 pd.read_html 兜底解析，
        这是「.xls 读取失败：cannot detect file format」最常见的根因。

    任何引擎失败时，会在异常信息中清晰指出已尝试的引擎及最终原因，便于排查。
    """
    ext = os.path.splitext(path)[1].lower()

    if ext == ".xls":
        engines = ["calamine", "xlrd"]
        html_fallback = True
    else:  # .xlsx / .xlsm / .xltx / .xltm
        engines = ["calamine", "openpyxl"]
        html_fallback = False

    # 只尝试真正可用的引擎，避免「指定未安装引擎」直接抛 ImportError
    usable = [e for e in engines if _engine_available(e)]
    tried = []
    last_err = None
    for eng in usable:
        try:
            if eng == "xlrd":
                # xlrd 需显式指定中文编码，否则无 CODEPAGE 的 .xls 会乱码
                return _read_with_xlrd(path)
            return pd.read_excel(
                path, sheet_name=None, engine=eng, header=None, dtype=object
            )
        except Exception as e:  # 尝试下一个引擎
            tried.append(eng)
            last_err = e

    # HTML 兜底：很多「.xls」其实是 HTML 表格（如交易终端导出 / 网页另存）。
    # 交易终端导出的 HTML 常为 GBK/GB18030 编码，依次尝试多种编码以正确读取中文。
    if html_fallback:
        for enc in ("utf-8", "gbk", "gb18030", None):
            try:
                tables = pd.read_html(path, header=None, encoding=enc)
                if tables:
                    return {
                        f"Sheet{i + 1}": df.astype(object)
                        for i, df in enumerate(tables)
                    }
            except Exception as e:
                last_err = e
        tried.append("read_html")

    # 构造清晰报错
    if not tried:
        # 连一个可用引擎都没有（HTML 兜底也不可用）→ 提示安装依赖
        needed = "python-calamine" + (" / xlrd" if ext == ".xls" else " / openpyxl")
        raise RuntimeError(
            f"缺少可用的 Excel 读取引擎（需安装 {needed}），且 HTML 兜底不可用，"
            f"无法读取 {os.path.basename(path)}"
        )
    hint = ("（很可能是交易终端另存的 .xls 实为 HTML，已尝试 HTML 兜底仍失败）"
            if html_fallback else "")
    raise RuntimeError(
        f"读取失败，已尝试引擎 {tried} 均失败{hint}：{last_err}"
    )


def _sheet_rows(df):
    """把 DataFrame 转为「列表的列表」（字符串规整后），便于比对。"""
    rows = []
    for _, row in df.iterrows():
        rows.append([_norm_cell(v) for v in row.tolist()])
    return rows


def _drop_empty_rows(rows):
    """丢弃「所有单元格均为空」的行，避免尾部/空白填充行产生幽灵差异。"""
    return [r for r in rows if any(c for c in r)]


def _build_key_map(rows):
    """返回 (keys列表, key->[原行...] 字典)。"""
    keys = [_tuple_key(r) for r in rows]
    grouped = {}
    for r, k in zip(rows, keys):
        grouped.setdefault(k, []).append(r)
    return keys, grouped


def _tuple_key(row):
    return tuple(_key_cell(c) for c in row)


# ====================== 单工作表比对 ======================
def _compare_sheet(rows_a, rows_b, ignore_row_order=True):
    """比对两个工作表（均为「列表的列表」）。"""
    if ignore_row_order:
        keys_a, group_a = _build_key_map(rows_a)
        keys_b, group_b = _build_key_map(rows_b)
        ca, cb = Counter(keys_a), Counter(keys_b)
        only_a = ca - cb  # 仅在 A（基准）中的行 -> 减少
        only_b = cb - ca  # 仅在对面（B）中的行 -> 新增

        removed, added = [], []
        for k, cnt in only_a.items():
            removed.extend(group_a[k][:cnt])
        for k, cnt in only_b.items():
            added.extend(group_b[k][:cnt])
        # 排序，保证报告稳定可读
        removed.sort(key=lambda r: _tuple_key(r))
        added.sort(key=lambda r: _tuple_key(r))

        equal = (ca == cb)
        return {
            "equal": equal,
            "mode": "ignore_order",
            "rows_a": len(rows_a),
            "rows_b": len(rows_b),
            "removed": removed,    # 减少的行（仅基准有）
            "added": added,        # 新增的行（仅对照有）
        }

    # 按位置对齐比较
    equal = True
    cell_diffs = []
    max_r = max(len(rows_a), len(rows_b))
    for r in range(max_r):
        ra = rows_a[r] if r < len(rows_a) else None
        rb = rows_b[r] if r < len(rows_b) else None
        if ra is None or rb is None:
            equal = False
            cell_diffs.append((r, None, ra, rb))
            if len(cell_diffs) >= MAX_CELL_DIFFS:
                break
            continue
        max_c = max(len(ra), len(rb))
        for c in range(max_c):
            va = ra[c] if c < len(ra) else ""
            vb = rb[c] if c < len(rb) else ""
            if _key_cell(va) != _key_cell(vb):
                equal = False
                cell_diffs.append((r, c, va, vb))
                if len(cell_diffs) >= MAX_CELL_DIFFS:
                    break
        if len(cell_diffs) >= MAX_CELL_DIFFS:
            break
    return {
        "equal": equal,
        "mode": "aligned",
        "rows_a": len(rows_a),
        "rows_b": len(rows_b),
        "cell_diffs": cell_diffs,
    }


# ====================== 两个 Excel 文件比对 ======================
def compare_excel(path_a, path_b, ignore_row_order=True):
    """比对两个 Excel 文件内容，返回结构化结果。"""
    try:
        sheets_a = _read_excel_sheets(path_a)
    except Exception as e:
        return {"ok": False, "type": "excel", "error": f"读取失败(基准): {e}"}
    try:
        sheets_b = _read_excel_sheets(path_b)
    except Exception as e:
        return {"ok": False, "type": "excel", "error": f"读取失败(对照): {e}"}

    names_a = list(sheets_a.keys())
    names_b = list(sheets_b.keys())

    # 工作表匹配：先按名称精确匹配；名称未匹配到的，再按位置兜底匹配。
    # 这样可避免「Sheet1 / 策略持仓」这类仅命名不同的表被误报为「缺/多工作表」，
    # 进而让真正的内容差异（新增/减少行）能被正确比较出来。
    matched = set()
    pairs = []  # (基准表名, 对照表名)
    for na in names_a:
        if na in sheets_b:
            pairs.append((na, na))
            matched.add(na)
    rest_a = [n for n in names_a if n not in matched]
    rest_b = [n for n in names_b if n not in matched]
    k = min(len(rest_a), len(rest_b))
    for i in range(k):
        pairs.append((rest_a[i], rest_b[i]))
    only_a = rest_a[k:]   # 真正仅基准存在的工作表
    only_b = rest_b[k:]   # 真正仅对照存在的工作表

    sheet_results = []
    all_equal = True
    for na, nb in pairs:
        rows_a = _drop_empty_rows(_sheet_rows(sheets_a[na]))
        rows_b = _drop_empty_rows(_sheet_rows(sheets_b[nb]))
        sr = _compare_sheet(rows_a, rows_b, ignore_row_order)
        sr["name"] = na  # 以基准(A)工作表名为准展示
        sheet_results.append(sr)
        if not sr["equal"]:
            all_equal = False

    if only_a or only_b:
        all_equal = False

    return {
        "ok": True,
        "type": "excel",
        "all_equal": all_equal and not only_a and not only_b,
        "only_a": only_a,    # 仅基准存在的工作表
        "only_b": only_b,    # 仅对照存在的工作表
        "sheets": sheet_results,
    }


# ====================== 两个非表格文件比对 ======================
def compare_generic(path_a, path_b):
    """对无法做内容解析的文件，按大小 + MD5 比对。"""
    size_a = os.path.getsize(path_a)
    size_b = os.path.getsize(path_b)
    if size_a == size_b and _md5(path_a) == _md5(path_b):
        return {"ok": True, "type": "generic", "all_equal": True,
                "size_a": size_a, "size_b": size_b}
    return {"ok": True, "type": "generic", "all_equal": False,
            "size_a": size_a, "size_b": size_b}


# ====================== 单文件（按扩展名分发） ======================
def compare_files(name, path_a, path_b, ignore_row_order=True):
    """对一对同名文件做比对，自动按扩展名选择比对方式。"""
    ext = os.path.splitext(name)[1].lower()
    if ext in EXCEL_EXTS:
        return compare_excel(path_a, path_b, ignore_row_order)
    return compare_generic(path_a, path_b)


# ====================== 文件夹级比对 ======================
def compare_folders(folder_a, folder_b, recursive=False, ignore_row_order=True,
                    progress_cb=None):
    """比对两个文件夹内同名文件，返回完整结果结构。

    progress_cb(done, total, name)：可用于界面进度反馈（在调用线程中执行）。
    """
    map_a = scan_folder(folder_a, recursive)
    map_b = scan_folder(folder_b, recursive)
    names = sorted(set(map_a) | set(map_b))
    total = len(names)

    files = []
    for i, name in enumerate(names, 1):
        in_a = name in map_a
        in_b = name in map_b
        entry = {"name": name, "in_a": in_a, "in_b": in_b}
        if in_a and in_b:
            entry["result"] = compare_files(name, map_a[name], map_b[name],
                                             ignore_row_order)
        files.append(entry)
        if progress_cb:
            try:
                progress_cb(i, total, name)
            except Exception:
                pass

    return {
        "folder_a": folder_a,
        "folder_b": folder_b,
        "recursive": recursive,
        "ignore_row_order": ignore_row_order,
        "total": total,
        "files": files,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


# ====================== 统计 + 报告 ======================
def summarize(data):
    """根据比对结果计算统计数字。"""
    s = {
        "total": data["total"],
        "equal": 0,
        "different": 0,
        "only_a": 0,
        "only_b": 0,
        "error": 0,
    }
    for f in data["files"]:
        if not f["in_a"] or not f["in_b"]:
            if f["in_a"] and not f["in_b"]:
                s["only_b"] += 1
            elif f["in_b"] and not f["in_a"]:
                s["only_a"] += 1
            else:
                s["only_a"] += 1
                s["only_b"] += 1
            continue
        r = f.get("result", {})
        if not r.get("ok", False):
            s["error"] += 1
        elif r.get("all_equal"):
            s["equal"] += 1
        else:
            s["different"] += 1
    return s


def _short_rows(rows):
    """把差异行列表截断到 MAX_DETAIL_ROWS，并返回 (展示行, 是否截断)。"""
    if len(rows) <= MAX_DETAIL_ROWS:
        return rows, False
    return rows[:MAX_DETAIL_ROWS], True


def format_report(data):
    """把比对结果结构渲染为纯文本报告。"""
    stat = summarize(data)
    L = []
    L.append("=" * 64)
    L.append("结果比对报告")
    L.append(f"生成时间: {data['generated_at']}")
    L.append(f"基准文件夹: {data['folder_a']}")
    L.append(f"对照文件夹: {data['folder_b']}")
    L.append(f"选项: 递归子目录={'是' if data['recursive'] else '否'} | "
             f"忽略行顺序={'是' if data['ignore_row_order'] else '否'}")
    L.append("=" * 64)
    L.append("")
    L.append("【总体统计】")
    L.append(f"  文件总数(按文件名匹配): {stat['total']}")
    L.append(f"  内容一致: {stat['equal']}")
    L.append(f"  内容不一致: {stat['different']}")
    L.append(f"  仅在基准(A): {stat['only_a']}")
    L.append(f"  仅在对面(B): {stat['only_b']}")
    L.append(f"  读取/比对错误: {stat['error']}")
    L.append("")

    # —— 文件明细 ——
    L.append("【文件明细】")
    for f in data["files"]:
        if not f["in_a"]:
            L.append(f"  [仅B]      {f['name']}")
        elif not f["in_b"]:
            L.append(f"  [仅A]      {f['name']}")
        else:
            r = f.get("result", {})
            if not r.get("ok", False):
                L.append(f"  [错误]     {f['name']}  ({r.get('error','')})")
            elif r.get("all_equal"):
                L.append(f"  [一致]     {f['name']}")
            else:
                if r.get("type") == "excel":
                    if r.get("ignore_row_order", data["ignore_row_order"]):
                        only_a = sum(len(s.get("removed", [])) for s in r.get("sheets", []))
                        only_b = sum(len(s.get("added", [])) for s in r.get("sheets", []))
                        summary = f"新增{only_b}行 / 减少{only_a}行"
                    else:
                        diffs = sum(len(s.get("cell_diffs", [])) for s in r.get("sheets", []))
                        summary = f"{diffs}处单元格差异"
                    if r.get("only_a") or r.get("only_b"):
                        extra = []
                        if r["only_a"]:
                            extra.append(f"缺失工作表{A}" if False else "缺少工作表:" + "/".join(map(str, r["only_a"])))
                        if r["only_b"]:
                            extra.append("多余工作表:" + "/".join(map(str, r["only_b"])))
                        summary += "；" + "，".join(extra)
                    L.append(f"  [不一致]   {f['name']}  ({summary})")
                else:
                    sa, sb = r.get("size_a"), r.get("size_b")
                    L.append(f"  [不一致]   {f['name']}  (大小 A={sa} / B={sb}，内容不同)")
    L.append("")

    # —— 不一致详情 ——
    detail_files = [
        f for f in data["files"]
        if f["in_a"] and f["in_b"] and (r := f.get("result", {}))
        and r.get("ok") and not r.get("all_equal")
    ]
    if detail_files:
        L.append("【不一致详情】")
        for f in detail_files:
            r = f["result"]
            L.append(f"  ▶ {f['name']}")
            if r.get("type") == "excel":
                # 仅存在于一侧的工作表
                if r.get("only_a"):
                    L.append(f"    仅基准存在的工作表: {', '.join(map(str, r['only_a']))}")
                if r.get("only_b"):
                    L.append(f"    仅对照存在的工作表: {', '.join(map(str, r['only_b']))}")
                for s in r.get("sheets", []):
                    if s.get("equal"):
                        continue
                    L.append(f"    · 工作表「{s['name']}」 "
                             f"(基准 {s['rows_a']} 行 / 对照 {s['rows_b']} 行)")
                    if s.get("mode") == "ignore_order":
                        removed, rtrunc = _short_rows(s.get("removed", []))
                        added, atrunc = _short_rows(s.get("added", []))
                        L.append(f"      - 减少的行 ({len(s.get('removed', []))}):")
                        for row in removed:
                            L.append("          " + " | ".join(row))
                        if rtrunc:
                            L.append("          ...(已截断)")
                        L.append(f"      + 新增的行 ({len(s.get('added', []))}):")
                        for row in added:
                            L.append("          " + " | ".join(row))
                        if atrunc:
                            L.append("          ...(已截断)")
                    else:
                        diffs = s.get("cell_diffs", [])
                        shown, trunc = _short_rows(diffs)
                        L.append(f"      单元格差异 ({len(diffs)}):")
                        for (rr, cc, va, vb) in shown:
                            if cc is None:
                                def _fmt(v):
                                    return "无" if v is None else " | ".join(str(x) for x in v)
                                L.append(f"        行{rr + 1}: 基准={_fmt(va)} / 对照={_fmt(vb)}")
                            else:
                                L.append(f"        行{rr + 1}列{cc + 1}: 基准={va!r} / 对照={vb!r}")
                        if trunc:
                            L.append("          ...(已截断)")
            else:
                L.append(f"    大小: 基准 {r.get('size_a')} / 对照 {r.get('size_b')}；MD5 不一致")
            L.append("")

    L.append("=" * 64)
    L.append("报告结束")
    return "\n".join(L)


def export_report(data, path):
    """把报告写入文件（UTF-8）。"""
    text = format_report(data)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


# 屏蔽 PyInstaller 打包时未安装 pandas 的告警噪音（运行期仍会给出可读错误）
if not HAS_PANDAS:  # pragma: no cover
    import logging
    logging.getLogger(__name__).warning("未检测到 pandas，Excel 比对功能不可用")
