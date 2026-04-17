import os
import warnings
from datetime import datetime, timedelta
import builtins
import pandas as pd
import pythoncom
import win32com.client as win32
from core.config import settings


warnings.filterwarnings("ignore")

"""
网关数据处理模块 - pywin32版
- 全面替换xlwings，避免多实例/读写冲突
- 批量公式备份/恢复
- 强制关键列文本格式，避免科学计数法
- 修复通报各区子表单列复制错位（如G25:G37 -> S25:S37）
"""

XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_XLUP = -4162
XL_XLTOLEFT = -4159
XL_PASTE_ALL = -4104


def get_file_modify_time(file_path):
    try:
        file_path_str = builtins.str(file_path)
        modify_timestamp = os.path.getmtime(file_path_str)
        modify_time = datetime.fromtimestamp(modify_timestamp)
        return modify_time
    except Exception:
        return None


def check_if_latest_file(file_path, time_threshold_hours=24):
    file_path_str = builtins.str(file_path)
    if not os.path.exists(file_path_str):
        return False
    file_modify_time = get_file_modify_time(file_path_str)
    if not file_modify_time:
        return False
    latest_valid_time = datetime.now() - timedelta(hours=time_threshold_hours)
    return file_modify_time >= latest_valid_time


def get_latest_gateway_file(base_path):
    """
    获取最新的边缘网关文件（优先选择修改时间最新的文件）
    同时检查xlsx和xls格式，选择最新的一个
    """
    base_path_str = builtins.str(base_path)
    xlsx_path = base_path_str + ".xlsx"
    xls_path = base_path_str + ".xls"

    candidates = []
    if os.path.exists(xlsx_path):
        mtime = get_file_modify_time(xlsx_path)
        if mtime:
            candidates.append((mtime, xlsx_path))
    if os.path.exists(xls_path):
        mtime = get_file_modify_time(xls_path)
        if mtime:
            candidates.append((mtime, xls_path))

    if not candidates:
        return None

    candidates.sort(reverse=True, key=lambda x: x[0])
    return candidates[0][1]


def load_excel_data(file_path, sheet_name=0):
    try:
        file_path_str = builtins.str(file_path)
        if not os.path.exists(file_path_str):
            raise FileNotFoundError(f"文件不存在：{file_path_str}")
        df = pd.read_excel(file_path_str, sheet_name=sheet_name, dtype=builtins.str)
        return df.fillna("")
    except Exception:
        return pd.DataFrame()


def preprocess_gateway_data(gateway_raw_path):
    df_gateway = load_excel_data(gateway_raw_path)
    if df_gateway.empty:
        return pd.DataFrame()

    core_fields = ["设备编码", "设备名称", "国家行政区县", "所属站址名称", "站址资源编码", "设备当前状态"]
    for field in core_fields:
        if field not in df_gateway.columns:
            df_gateway[field] = ""

    df_main = df_gateway[core_fields].copy()
    new_fields = ["是否临时入网", "整合站名与入临时入网", "分管维护员", "区域", "片区", "是否代维处理", "核减", "是否超7天"]
    for field in new_fields:
        df_main[field] = ""
    return df_main


def _normalize_datetime_text(value):
    if value is None:
        return ""
    text = builtins.str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    dt = pd.to_datetime(text, errors="coerce")
    if pd.isna(dt):
        return text
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def _iter_worksheets(wb):
    for i in range(1, wb.Worksheets.Count + 1):
        yield wb.Worksheets(i)


def _get_sheet_by_name(wb, target_name):
    target = builtins.str(target_name).strip()
    for sh in _iter_worksheets(wb):
        if builtins.str(sh.Name).strip() == target:
            return sh
    return None


def _normalize_matrix(values):
    if values is None:
        return []
    if isinstance(values, tuple):
        if not values:
            return []
        if isinstance(values[0], tuple):
            return [list(r) for r in values]
        return [list(values)]
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        if isinstance(values[0], tuple):
            return [list(r) for r in values]
        return [values]
    return [[values]]


def _normalize_row(values):
    matrix = _normalize_matrix(values)
    if not matrix:
        return []
    if len(matrix) == 1:
        return matrix[0]
    return [r[0] if r else "" for r in matrix]


def _to_column_matrix(values, row_count):
    if row_count <= 0:
        return []
    matrix = _normalize_matrix(values)
    if not matrix:
        return [[""] for _ in range(row_count)]

    # 单列范围通常返回 [[v1], [v2], ...]；这里统一成严格二维单列，避免写入错位。
    flat = []
    for row in matrix:
        if isinstance(row, list) and row:
            flat.append(row[0])
        elif isinstance(row, list):
            flat.append("")
        else:
            flat.append(row)

    if len(flat) < row_count:
        flat.extend([""] * (row_count - len(flat)))
    return [[flat[i]] for i in range(row_count)]


def _to_excel_matrix(rows):
    return tuple(tuple(r) for r in rows)


def _has_effective_reduce_value(value):
    if value is None:
        return False
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return False
        if text.startswith("#"):
            return False
        return True
    if isinstance(value, (int, float)):
        # Excel错误值在COM中可能以数值错误码返回（如 #N/A -> 2042）
        iv = int(value)
        if 2000 <= iv <= 2047:
            return False
        return True
    return True


def _normalize_lookup_key(value):
    if value is None:
        return ""
    if isinstance(value, str):
        text = value.strip()
        # Excel中常通过前导'强制文本，这里统一去掉以便键值匹配。
        if text.startswith("'"):
            text = text[1:].strip()
        return text
    if isinstance(value, (int, float)):
        # 统一 123.0 / 123 为同一键
        if float(value).is_integer():
            return builtins.str(int(value))
        return builtins.str(value).strip()
    return builtins.str(value).strip()


def _contiguous_ranges(row_numbers):
    if not row_numbers:
        return []
    sorted_rows = sorted(set(row_numbers))
    ranges = []
    start = sorted_rows[0]
    end = sorted_rows[0]
    for r in sorted_rows[1:]:
        if r == end + 1:
            end = r
        else:
            ranges.append((start, end))
            start = r
            end = r
    ranges.append((start, end))
    return ranges


def _open_excel_app():
    pythoncom.CoInitialize()
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    app.Interactive = False
    app.EnableEvents = False
    return app


def _close_excel_app(app):
    if app is None:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return

    try:
        app.Calculation = XL_CALCULATION_AUTOMATIC
        app.ScreenUpdating = True
        app.DisplayAlerts = True
        app.Interactive = True
        app.EnableEvents = True
        app.CutCopyMode = False
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


def _cell_addr(ws, row, col):
    # 某些pywin32环境中 Address 是字符串属性而非可调用方法
    # 统一取属性值并去掉 $，确保得到 A1 形式相对引用
    return ws.Cells(row, col).Address.replace("$", "")


def batch_backup_all_sheets_formulas(wb, exclude_sheet="网关总清单"):
    backup = {}
    for sh in _iter_worksheets(wb):
        if builtins.str(sh.Name) == builtins.str(exclude_sheet):
            continue
        try:
            used = sh.UsedRange
            backup[sh.Name] = {"range": used.Address, "formulas": used.Formula}
        except Exception:
            continue
    return backup


def batch_restore_all_sheets_formulas(wb, backup, exclude_sheet="网关总清单"):
    if not backup:
        return
    for name, data in backup.items():
        if builtins.str(name) == builtins.str(exclude_sheet):
            continue
        sh = _get_sheet_by_name(wb, name)
        if not sh:
            continue
        try:
            sh.Range(data["range"]).Formula = data["formulas"]
        except Exception:
            continue


def force_sheet_text_format(ws, key_columns=None, max_row=None):
    if key_columns is None:
        key_columns = ["设备编码", "站址资源编码", "设备编号", "入网设备编码", "站址编码"]

    try:
        last_col = ws.Cells(1, ws.Columns.Count).End(XL_XLTOLEFT).Column
        if last_col < 1:
            return
        if max_row is None:
            used = ws.UsedRange
            max_row = max(1, used.Row + used.Rows.Count - 1)
        headers = _normalize_row(ws.Range(ws.Cells(1, 1), ws.Cells(1, last_col)).Value)
        for col_idx, header in enumerate(headers, start=1):
            if header in key_columns:
                ws.Range(ws.Cells(1, col_idx), ws.Cells(max_row, col_idx)).NumberFormat = "@"
    except Exception:
        return


def force_all_sheets_text_format(wb, key_columns=None):
    for sh in _iter_worksheets(wb):
        force_sheet_text_format(sh, key_columns=key_columns)


def update_gateway_sheet_vol():
    """
    处理通报各区子表（支持同工作表多个独立表格）
    修复重点：所有单列复制统一写成二维单列，避免复制到S:AE错位
    """
    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print("通报各区子表处理失败：目标Excel不存在")
        return

    app = None
    wb = None
    formula_cache = {}

    try:
        app = _open_excel_app()
        app.Calculation = XL_CALCULATION_MANUAL
        wb = app.Workbooks.Open(target_excel_path)

        ws = _get_sheet_by_name(wb, "通报各区")
        if not ws:
            print("通报各区子表处理失败：未找到「通报各区」工作表")
            return

        used_range = ws.UsedRange
        used_formulas = _normalize_matrix(used_range.Formula)
        base_row = used_range.Row
        base_col = used_range.Column
        for r_idx, row in enumerate(used_formulas):
            for c_idx, cell_formula in enumerate(row):
                if isinstance(cell_formula, str) and cell_formula.startswith("="):
                    formula_cache[(base_row + r_idx, base_col + c_idx)] = cell_formula

        copy_operations = [
            (3, 15, 6, 15),   # F3-F15 -> O3-O15
            (3, 15, 7, 16),   # G3-G15 -> P3-P15
            (3, 15, 3, 10),   # C3-C15 -> J3-J15
            (3, 15, 24, 25),  # X3-X15 -> Y3-Y15
            (25, 37, 6, 18),  # F25-F37 -> R25-R37
            (25, 37, 4, 13),  # D25-D37 -> M25-M37
            (25, 37, 7, 19),  # G25-G37 -> S25-S37
        ]

        for start_row, end_row, src_col, dst_col in copy_operations:
            row_count = end_row - start_row + 1
            src_rng = ws.Range(ws.Cells(start_row, src_col), ws.Cells(end_row, src_col))
            dst_rng = ws.Range(ws.Cells(start_row, dst_col), ws.Cells(end_row, dst_col))
            src_data = src_rng.Value
            dst_rng.Value = _to_excel_matrix(_to_column_matrix(src_data, row_count))

        ws.Range(ws.Cells(3, 16), ws.Cells(15, 16)).NumberFormat = "0.00%"

        for (row, col), formula in formula_cache.items():
            ws.Cells(row, col).Formula = formula

        app.Calculate()
        wb.Save()
        print("通报各区子表处理完成")
    except Exception as e:
        print(f"通报各区子表处理失败：{builtins.str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=0)
            except Exception:
                pass
        _close_excel_app(app)


def update_gateway_sheet_with_pywin32(target_excel_path, sheet_name, df_updated):
    if df_updated.empty:
        print("网关总清单处理失败：无更新数据")
        return False

    app = None
    wb = None
    formula_backup = {}

    try:
        app = _open_excel_app()
        app.Calculation = XL_CALCULATION_MANUAL
        wb = app.Workbooks.Open(target_excel_path)

        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet=sheet_name)

        ws = _get_sheet_by_name(wb, sheet_name)
        if not ws:
            print(f"网关总清单处理失败：目标Sheet「{sheet_name}」不存在")
            return False

        start_row = 2
        last_row = ws.Cells(ws.Rows.Count, 1).End(XL_XLUP).Row
        if last_row >= start_row:
            ws.Range(f"A{start_row}:ZZ{last_row}").ClearContents()

        headers = _normalize_row(ws.Range("A1:ZZ1").Value)
        header_dict = {col: idx + 1 for idx, col in enumerate(headers) if col}
        valid_cols = [col for col in df_updated.columns if col in header_dict]
        if not valid_cols:
            print("网关总清单处理失败：无匹配的列名")
            return False

        df_valid = df_updated[valid_cols].copy()
        data_rows = len(df_valid)
        if data_rows == 0:
            return True

        for col_name in ["设备编码", "站址资源编码"]:
            if col_name in header_dict:
                col_idx = header_dict[col_name]
                ws.Range(ws.Cells(1, col_idx), ws.Cells(data_rows + 1, col_idx)).NumberFormat = "@"
                if col_name in df_valid.columns:
                    df_valid[col_name] = df_valid[col_name].apply(lambda x: f"'{x}" if builtins.str(x).strip() else x)

        sorted_cols = [c for c in valid_cols if c in header_dict]
        if sorted_cols:
            sorted_col_indices = [header_dict[c] for c in sorted_cols]
            min_col, max_col = min(sorted_col_indices), max(sorted_col_indices)
            data_matrix = [[""] * (max_col - min_col + 1) for _ in range(data_rows)]
            for col_idx, col_name in enumerate(sorted_cols):
                target_col_pos = sorted_col_indices[col_idx] - min_col
                col_data = df_valid[col_name].tolist()
                for row_idx in range(data_rows):
                    data_matrix[row_idx][target_col_pos] = col_data[row_idx]
            ws.Range(ws.Cells(2, min_col), ws.Cells(data_rows + 1, max_col)).Value = _to_excel_matrix(data_matrix)

        if "设备当前状态" in df_updated.columns and "设备当前状态" in header_dict:
            m_col_idx = header_dict["设备当前状态"]
            status_data = [[val] for val in df_updated["设备当前状态"].tolist()]
            ws.Range(ws.Cells(2, m_col_idx), ws.Cells(data_rows + 1, m_col_idx)).Value = _to_excel_matrix(status_data)
            ws.Range(ws.Cells(1, m_col_idx), ws.Cells(data_rows + 1, m_col_idx)).NumberFormat = "@"

        if "是否临时入网" in header_dict and "设备编码" in header_dict:
            code_col = header_dict["设备编码"]
            temp_col = header_dict["是否临时入网"]
            formula = f'=VLOOKUP({_cell_addr(ws, 2, code_col)},数据源!N:P,3,0)'
            ws.Range(ws.Cells(2, temp_col), ws.Cells(data_rows + 1, temp_col)).Formula = formula

        if all(col in header_dict for col in ["整合站名与入临时入网", "所属站址名称", "是否临时入网"]):
            integrate_col = header_dict["整合站名与入临时入网"]
            site_col = header_dict["所属站址名称"]
            temp_col = header_dict["是否临时入网"]
            formula = f'=IF(ISNA({_cell_addr(ws, 2, temp_col)}), {_cell_addr(ws, 2, site_col)}, {_cell_addr(ws, 2, temp_col)})'
            ws.Range(ws.Cells(2, integrate_col), ws.Cells(data_rows + 1, integrate_col)).Formula = formula

        if "整合站名与入临时入网" in header_dict:
            integrate_col = header_dict["整合站名与入临时入网"]
            integrate_cell = _cell_addr(ws, 2, integrate_col)
            if "分管维护员" in header_dict:
                col = header_dict["分管维护员"]
                ws.Range(ws.Cells(2, col), ws.Cells(data_rows + 1, col)).Formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,3,0)'
            if "区域" in header_dict:
                col = header_dict["区域"]
                ws.Range(ws.Cells(2, col), ws.Cells(data_rows + 1, col)).Formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,2,0)'
            if "片区" in header_dict:
                col = header_dict["片区"]
                ws.Range(ws.Cells(2, col), ws.Cells(data_rows + 1, col)).Formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,5,0)'

        if "是否超7天" in header_dict and "设备编码" in header_dict:
            over7_col = header_dict["是否超7天"]
            code_col = header_dict["设备编码"]
            code_cell = _cell_addr(ws, 2, code_col)
            formula = f'=IF(ISNA(VLOOKUP({code_cell},网关离线清单!A:B,2,0)),"",IF(VLOOKUP({code_cell},网关离线清单!A:B,2,0)>=7,"是",""))'
            ws.Range(ws.Cells(2, over7_col), ws.Cells(data_rows + 1, over7_col)).Formula = formula

        if "是否代维处理" in header_dict:
            dv_col = header_dict["是否代维处理"]
            ws.Range(ws.Cells(2, dv_col), ws.Cells(data_rows + 1, dv_col)).Value = "是"

        if "核减" in header_dict and "设备编码" in header_dict:
            reduce_col = header_dict["核减"]
            code_col = header_dict["设备编码"]
            code_cell = _cell_addr(ws, 2, code_col)
            ws.Range(ws.Cells(2, reduce_col), ws.Cells(data_rows + 1, reduce_col)).Formula = f"=VLOOKUP({code_cell},核减清单!A:C,3,0)"

        # 所有其他公式写完后，最后强制写入 H/I，避免被后续列公式覆盖。
        # 使用逐行A1公式矩阵，规避 AutoFill/FormulaR1C1 在不同环境下的不一致行为。
        # H列 = VLOOKUP(G2,数据源!V:Z,3,0)
        # I列 = VLOOKUP(G2,数据源!V:Z,2,0)
        h_formulas = [[f"=VLOOKUP(G{row_idx},数据源!V:Z,3,0)"] for row_idx in range(2, data_rows + 2)]
        i_formulas = [[f"=VLOOKUP(G{row_idx},数据源!V:Z,2,0)"] for row_idx in range(2, data_rows + 2)]
        ws.Range(ws.Cells(2, 8), ws.Cells(data_rows + 1, 8)).Formula = _to_excel_matrix(h_formulas)
        ws.Range(ws.Cells(2, 9), ws.Cells(data_rows + 1, 9)).Formula = _to_excel_matrix(i_formulas)
        print(f"H/I公式写入完成，行数：{data_rows}")

        # 固定列规则：若 L 列（核减）有有效值，则清空对应行 H/I；否则保留 H/I（含公式）。
        # 不再做Python侧键值匹配，直接以Excel计算后的L列显示结果为准。
        # 注意：这里必须逐行读单元格Text，避免多单元格Range.Text在COM下失真。
        app.Calculate()
        reduce_col = header_dict.get("核减", 12)
        print(f"网关总清单关键列：设备编码列={header_dict.get('设备编码', 'NOT_FOUND')}，核减列={reduce_col}")

        clear_rows = []
        blank_rows = 0
        na_rows = 0
        for i in range(data_rows):
            row_no = i + 2
            text = builtins.str(ws.Cells(row_no, reduce_col).Text).strip()
            if not text:
                blank_rows += 1
                continue
            if text.startswith("#"):
                na_rows += 1
                continue
            clear_rows.append(row_no)
        print(
            f"L列识别统计：空白={blank_rows}，错误值={na_rows}，有效匹配={len(clear_rows)}（将清空H/I）"
        )

        if clear_rows:
            for start_row, end_row in _contiguous_ranges(clear_rows):
                ws.Range(ws.Cells(start_row, 8), ws.Cells(end_row, 9)).ClearContents()

            backup_sheet_name = "网关总清单_核减清理备份"
            backup_ws = _get_sheet_by_name(wb, backup_sheet_name)
            if backup_ws:
                backup_ws.Delete()
            ws_backup = wb.Worksheets.Add(After=ws)
            ws_backup.Name = backup_sheet_name
            ws.UsedRange.Copy()
            ws_backup.Range("A1").PasteSpecial(Paste=XL_PASTE_ALL)
            ws_backup.Visible = False
            app.CutCopyMode = False

        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet=sheet_name)
        force_sheet_text_format(ws, max_row=data_rows + 1)

        app.Calculate()
        wb.Save()
        print("网关总清单处理完成")
        return True
    except Exception as e:
        print(f"网关总清单处理失败：{builtins.str(e)}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=0)
            except Exception:
                pass
        _close_excel_app(app)


def process_gateway_main_list():
    # 先处理通报各区，避免后续公式或筛选影响子表复制
    update_gateway_sheet_vol()

    gateway_base_path = settings.resolve_path("spider/script/station/down/边缘网关")
    gateway_raw_path = get_latest_gateway_file(gateway_base_path)
    if not gateway_raw_path:
        print("网关总清单处理失败：边缘网关数据文件不存在")
        return

    if not check_if_latest_file(gateway_raw_path, 24):
        print("网关总清单处理失败：边缘网关文件不是最新版本")
        return

    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print("网关总清单处理失败：目标Excel不存在")
        return

    df_updated = preprocess_gateway_data(gateway_raw_path)
    return update_gateway_sheet_with_pywin32(
        target_excel_path=target_excel_path,
        sheet_name="网关总清单",
        df_updated=df_updated,
    )


def update_gateway_offline_list():
    gateway_base_path = settings.resolve_path("spider/script/station/down/边缘网关")
    gateway_raw_path = get_latest_gateway_file(gateway_base_path)

    if not gateway_raw_path:
        print("网关离线清单处理失败：边缘网关数据文件不存在")
        return
    if not check_if_latest_file(gateway_raw_path, 24):
        print("网关离线清单处理失败：边缘网关文件不是最新版本")
        return

    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print("网关离线清单处理失败：目标Excel不存在")
        return

    df_gateway = load_excel_data(gateway_raw_path)
    if df_gateway.empty:
        print("网关离线清单处理失败：边缘网关数据加载失败")
        return

    required_fields = [
        "设备编码", "设备名称", "设备当前状态", "最近离线时间",
        "离线恢复时间", "设备入网状态", "国家行政区县", "所属站址名称", "站址资源编码"
    ]
    for field in required_fields:
        if field not in df_gateway.columns:
            df_gateway[field] = ""

    if "设备当前状态" in df_gateway.columns:
        df_gateway = df_gateway[df_gateway["设备当前状态"] == "离线"]

    df_updated = df_gateway[required_fields].copy()
    if "最近离线时间" in df_updated.columns:
        df_updated["最近离线时间"] = df_updated["最近离线时间"].apply(_normalize_datetime_text)

    data_rows = len(df_updated)
    if data_rows == 0:
        print("网关离线清单处理完成：无离线网关数据")
        return

    app = None
    wb = None
    try:
        app = _open_excel_app()
        app.Calculation = XL_CALCULATION_MANUAL
        wb = app.Workbooks.Open(target_excel_path)

        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet="网关离线清单")
        ws = _get_sheet_by_name(wb, "网关离线清单")
        if not ws:
            print("网关离线清单处理失败：目标Sheet「网关离线清单」不存在")
            return

        start_row = 2
        last_row = ws.Cells(ws.Rows.Count, 1).End(XL_XLUP).Row
        if last_row >= start_row:
            ws.Range(f"A{start_row}:ZZ{last_row}").ClearContents()

        headers = _normalize_row(ws.Range("A1:ZZ1").Value)
        header_dict = {col: idx + 1 for idx, col in enumerate(headers) if col}

        missing_required = [f for f in required_fields if f not in header_dict]
        if missing_required:
            print(f"网关离线清单处理失败：表头中缺少必填字段 {missing_required}")
            return

        sorted_cols = [field for field in required_fields if field in header_dict]
        sorted_col_indices = [header_dict[field] for field in sorted_cols]
        min_col, max_col = min(sorted_col_indices), max(sorted_col_indices)

        data_matrix = [[""] * (max_col - min_col + 1) for _ in range(data_rows)]
        for col_idx, field in enumerate(sorted_cols):
            target_col_pos = sorted_col_indices[col_idx] - min_col
            col_data = df_updated[field].tolist()
            for row_idx in range(data_rows):
                data_matrix[row_idx][target_col_pos] = col_data[row_idx]
        ws.Range(ws.Cells(2, min_col), ws.Cells(data_rows + 1, max_col)).Value = _to_excel_matrix(data_matrix)

        if "离线天数" in header_dict and "最近离线时间" in header_dict:
            offline_days_col = header_dict["离线天数"]
            recent_offline_col = header_dict["最近离线时间"]
            ws.Range(ws.Cells(2, offline_days_col), ws.Cells(data_rows + 1, offline_days_col)).Formula = (
                f'=IF({_cell_addr(ws, 2, recent_offline_col)}="","",TODAY()-INT({_cell_addr(ws, 2, recent_offline_col)}))'
            )

        if "是否代维处理" in header_dict and "设备编码" in header_dict:
            dv_col = header_dict["是否代维处理"]
            code_col = header_dict["设备编码"]
            ws.Range(ws.Cells(2, dv_col), ws.Cells(data_rows + 1, dv_col)).Formula = (
                f'=VLOOKUP({_cell_addr(ws, 2, code_col)},网关总清单!$A$2:$K$8876,11,0)'
            )

        if "分管维护员" in header_dict and "设备名称" in header_dict:
            leader_col = header_dict["分管维护员"]
            name_col = header_dict["设备名称"]
            ws.Range(ws.Cells(2, leader_col), ws.Cells(data_rows + 1, leader_col)).Formula = (
                f'=VLOOKUP({_cell_addr(ws, 2, name_col)},网关总清单!B:I,7,0)'
            )

        if "临时入网实际安装站" in header_dict and "设备编码" in header_dict:
            temp_site_col = header_dict["临时入网实际安装站"]
            code_col = header_dict["设备编码"]
            ws.Range(ws.Cells(2, temp_site_col), ws.Cells(data_rows + 1, temp_site_col)).Formula = (
                f'=VLOOKUP({_cell_addr(ws, 2, code_col)},网关总清单!$A$2:$F$8876,6,0)'
            )

        if "备注" in header_dict and "设备编码" in header_dict:
            remark_col = header_dict["备注"]
            code_col = header_dict["设备编码"]
            ws.Range(ws.Cells(2, remark_col), ws.Cells(data_rows + 1, remark_col)).Formula = (
                f'=VLOOKUP({_cell_addr(ws, 2, code_col)},核减清单!$A$2:$C$537,3,0)'
            )

        if "片区" in header_dict and "所属站址名称" in header_dict:
            region_col = header_dict["片区"]
            site_col = header_dict["所属站址名称"]
            ws.Range(ws.Cells(2, region_col), ws.Cells(data_rows + 1, region_col)).Formula = (
                f'=VLOOKUP({_cell_addr(ws, 2, site_col)},数据源!V:Z,5,0)'
            )

        if "考核核减" in header_dict and "备注" in header_dict:
            check_col = header_dict["考核核减"]
            remark_col = header_dict["备注"]
            ws.Range(ws.Cells(2, check_col), ws.Cells(data_rows + 1, check_col)).Formula = (
                f'=IF({_cell_addr(ws, 2, remark_col)}="","","核减")'
            )

        if "是否超7天" in header_dict and "离线天数" in header_dict:
            over7_col = header_dict["是否超7天"]
            offline_days_col = header_dict["离线天数"]
            ws.Range(ws.Cells(2, over7_col), ws.Cells(data_rows + 1, over7_col)).Formula = (
                f'=IF({_cell_addr(ws, 2, offline_days_col)}="","",IF({_cell_addr(ws, 2, offline_days_col)}>=7,"是","否"))'
            )

        if "国家行政区县" in header_dict and "设备名称" in header_dict:
            region_col = header_dict["国家行政区县"]
            name_col = header_dict["设备名称"]
            ws.Range(ws.Cells(2, region_col), ws.Cells(data_rows + 1, region_col)).Formula = (
                f'=VLOOKUP({_cell_addr(ws, 2, name_col)},网关总清单!$B:$J,8,0)'
            )

        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet="网关离线清单")
        force_sheet_text_format(ws, max_row=data_rows + 1)

        app.Calculate()
        wb.Save()
        print("网关离线清单处理完成")
    except Exception as e:
        print(f"网关离线清单处理失败：{builtins.str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=0)
            except Exception:
                pass
        _close_excel_app(app)


if __name__ == "__main__":
    process_gateway_main_list()
    update_gateway_offline_list()
