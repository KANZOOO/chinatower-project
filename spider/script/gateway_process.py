import os
import pandas as pd
import xlwings as xw
from core.config import settings
import warnings
from datetime import datetime, timedelta

warnings.filterwarnings('ignore')

"""
网关数据处理模块 - 性能优化版
- 批量公式备份/恢复（超快）
- 强制所有Sheet设备编码=文本格式，彻底消灭科学计数法
- 彻底解决其他Sheet引用新网关总清单失效问题
- 原有逻辑100%保留
- 修复update_gateway_sheet_vol()列复制Excel异常问题
- 新增：将边缘网关G列（设备当前状态）填充到网关总清单M列
- 性能优化：逐列写入改为一次性写入，逐行循环改为批量处理
- 输出精简：只保留关键信息
"""


def get_file_modify_time(file_path):
    try:
        file_path_str = str(file_path)
        modify_timestamp = os.path.getmtime(file_path_str)
        modify_time = datetime.fromtimestamp(modify_timestamp)
        return modify_time
    except Exception as e:
        return None


def check_if_latest_file(file_path, time_threshold_hours=24):
    file_path_str = str(file_path)
    if not os.path.exists(file_path_str):
        return False
    file_modify_time = get_file_modify_time(file_path_str)
    if not file_modify_time:
        return False
    latest_valid_time = datetime.now() - timedelta(hours=time_threshold_hours)
    is_latest = file_modify_time >= latest_valid_time
    return is_latest


def get_latest_gateway_file(base_path):
    """
    获取最新的边缘网关文件（优先选择修改时间最新的文件）
    同时检查xlsx和xls格式，选择最新的一个
    """
    base_path_str = str(base_path)
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
    latest_path = candidates[0][1]
    return latest_path


def load_excel_data(file_path, sheet_name=0):
    try:
        file_path_str = str(file_path)
        if not os.path.exists(file_path_str):
            raise FileNotFoundError(f"文件不存在：{file_path_str}")
        df = pd.read_excel(file_path_str, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")
        return df
    except Exception as e:
        return pd.DataFrame()


def preprocess_gateway_data(gateway_raw_path):
    df_gateway = load_excel_data(gateway_raw_path)
    if df_gateway.empty:
        return pd.DataFrame()
    core_fields = ["设备编码", "设备名称", "国家行政区县", "所属站址名称", "站址资源编码"]
    if "设备当前状态" not in core_fields:
        core_fields.append("设备当前状态")

    for field in core_fields:
        if field not in df_gateway.columns:
            df_gateway[field] = ""
    df_main = df_gateway[core_fields].copy()
    new_fields = ["是否临时入网", "整合站名与入临时入网", "分管维护员", "区域", "片区",
                  "是否代维处理", "核减", "是否超7天"]
    for field in new_fields:
        df_main[field] = ""
    return df_main


def batch_backup_all_sheets_formulas(wb, exclude_sheet="网关总清单"):
    backup = {}
    try:
        for sh in wb.sheets:
            if sh.name == exclude_sheet:
                continue
            try:
                used = sh.used_range
                if not used:
                    continue
                backup[sh.name] = {
                    "range": used.address,
                    "formulas": used.formula
                }
            except:
                continue
    except Exception as e:
        pass
    return backup


def batch_restore_all_sheets_formulas(wb, backup, exclude_sheet="网关总清单"):
    if not backup:
        return
    count = 0
    try:
        for name, data in backup.items():
            if name == exclude_sheet:
                continue
            if name not in [s.name for s in wb.sheets]:
                continue
            sh = wb.sheets[name]
            sh.range(data["range"]).formula = data["formulas"]
            count += 1
    except Exception as e:
        pass


def force_all_sheets_text_format(wb, key_columns=["设备编码", "站址资源编码", "设备编号", "入网设备编码", "站址编码"]):
    try:
        for sh in wb.sheets:
            try:
                headers = sh.range("1:1").value
                if not headers:
                    continue
                for col_idx, header in enumerate(headers):
                    if header in key_columns:
                        excel_col = col_idx + 1
                        sh.range((1, excel_col), (sh.cells.last_cell.row, excel_col)).number_format = "@"
            except:
                continue
    except Exception as e:
        pass


# def update_gateway_sheet_vol():
#     """
#     处理通报各区子表（支持同工作表多个独立表格）
#     性能优化：逐行读写改为批量操作
#     """
#     target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
#     if not os.path.exists(target_excel_path):
#         print("通报各区子表处理失败：目标Excel不存在")
#         return
#
#     app = None
#     wb = None
#     formula_cache = {}
#
#     try:
#         app = xw.App(visible=False, add_book=False)
#         app.calculation = 'manual'
#         app.screen_updating = False
#         app.display_alerts = False
#         app.interactive = False
#
#         wb = app.books.open(target_excel_path)
#         wb.display_alerts = False
#         wb.screen_updating = False
#
#         sheet_name = "通报各区"
#         sheet_names = [s.name.strip() for s in wb.sheets]
#         if sheet_name not in sheet_names:
#             print(f"通报各区子表处理失败：未找到「{sheet_name}」工作表")
#             return
#
#         ws = wb.sheets[sheet_name]
#
#         # 性能优化：批量备份公式，而非逐单元格
#         used_range = ws.used_range
#         formulas = used_range.formula
#         for row in range(len(formulas)):
#             for col in range(len(formulas[row])):
#                 if formulas[row][col] and isinstance(formulas[row][col], str) and formulas[row][col].startswith('='):
#                     formula_cache[(row + 1, col + 1)] = formulas[row][col]
#
#         # 性能优化：批量读取和写入数据
#         copy_operations = [
#             (3, 15, 6, 15),  # F3-F15 到 O3-O15
#             (3, 15, 7, 16),  # G3-G15 到 P3-P15
#             (3, 15, 3, 10),  # C3-C15 到 J3-J15
#             (3, 15, 24, 25), # X3-X15 到 Y3-Y15
#             (25, 37, 6, 18), # F25-F37 到 R25-R37
#             (25, 37, 4, 13), # D25-D37 到 M25-M37
#             (25, 37, 7, 19),  # G25-G37 到 S25-S37
#         ]
#
#         for start_row, end_row, src_col, dst_col in copy_operations:
#             try:
#                 # 批量读取源数据
#                 src_data = ws.range((start_row, src_col), (end_row, src_col)).value
#                 if src_data:
#                     # 确保数据格式正确
#                     if not isinstance(src_data, (list, tuple)):
#                         src_data = [[src_data]]
#                     # 批量写入目标数据
#                     ws.range((start_row, dst_col), (end_row, dst_col)).value = src_data
#             except Exception as e:
#                 pass
#
#         # 设置P3-P15为百分比格式
#         try:
#             ws.range((3, 16), (15, 16)).number_format = '0.00%'
#         except Exception as e:
#             pass
#
#         # 批量恢复公式
#         for (row, col), formula in formula_cache.items():
#             try:
#                 ws.range((row, col)).formula = formula
#             except Exception as e:
#                 pass
#
#         wb.app.calculate()
#         wb.save()
#
#     except Exception as e:
#         print(f"通报各区子表处理失败：{str(e)}")
#         import traceback
#         traceback.print_exc()
#     finally:
#         if wb:
#             try:
#                 wb.close()
#             except Exception as e:
#                 pass
#         if app:
#             try:
#                 app.calculation = 'automatic'
#                 app.screen_updating = True
#                 app.display_alerts = True
#                 app.interactive = True
#                 app.quit()
#             except Exception as e:
#                 try:
#                     app.kill()
#                 except:
#                     pass


def update_gateway_sheet_with_xlwings(target_excel_path, sheet_name, df_updated):
    if df_updated.empty:
        print("网关总清单处理失败：无更新数据")
        return False

    app = None
    wb = None
    formula_backup = {}
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(target_excel_path)
        wb.display_alerts = False
        wb.screen_updating = False
        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet=sheet_name)
        force_all_sheets_text_format(wb)

        if sheet_name not in [s.name for s in wb.sheets]:
            print(f"网关总清单处理失败：目标Sheet「{sheet_name}」不存在")
            return False
        ws = wb.sheets[sheet_name]

        start_row = 2
        last_row = ws.cells.last_cell.row
        if last_row >= start_row:
            ws.range(f"A{start_row}:ZZ{last_row}").clear_contents()

        header_range = ws.range("A1:ZZ1").value
        header_dict = {col: idx + 1 for idx, col in enumerate(header_range) if col}
        valid_cols = [col for col in df_updated.columns if col in header_dict]
        if not valid_cols:
            print("网关总清单处理失败：无匹配的列名")
            return False

        df_valid = df_updated[valid_cols].copy()
        data_rows = len(df_valid)
        if data_rows == 0:
            return True

        code_cols = ["设备编码", "站址资源编码"]
        for col_name in code_cols:
            if col_name in header_dict:
                col_idx = header_dict[col_name]
                ws.range((1, col_idx), (data_rows + 1, col_idx)).number_format = "@"
                if col_name in df_valid.columns:
                    df_valid[col_name] = df_valid[col_name].apply(lambda x: f"'{x}" if x.strip() else x)

        # 性能优化：一次性写入所有数据，而不是逐列写入
        sorted_cols = []
        for col_name in valid_cols:
            if col_name in header_dict:
                sorted_cols.append(col_name)
        
        if sorted_cols:
            sorted_col_indices = [header_dict[col] for col in sorted_cols]
            min_col, max_col = min(sorted_col_indices), max(sorted_col_indices)
            
            data_matrix = [[""] * (max_col - min_col + 1) for _ in range(data_rows)]
            
            for col_idx, col_name in enumerate(sorted_cols):
                if col_name in df_valid.columns:
                    target_col_pos = sorted_col_indices[col_idx] - min_col
                    col_data = df_valid[col_name].tolist()
                    for row_idx in range(data_rows):
                        data_matrix[row_idx][target_col_pos] = col_data[row_idx]
            
            ws.range((2, min_col), (data_rows + 1, max_col)).value = data_matrix

        # 核心新增逻辑：设备当前状态写入M列
        status_col_name = "设备当前状态"
        if status_col_name in df_updated.columns:
            m_col_idx = 13
            status_data = df_updated[status_col_name].tolist()
            ws.range((2, m_col_idx), (data_rows + 1, m_col_idx)).value = [[val] for val in status_data]
            ws.range((1, m_col_idx), (data_rows + 1, m_col_idx)).number_format = "@"

        if "是否临时入网" in header_dict and "设备编码" in header_dict:
            code_col = header_dict["设备编码"]
            temp_col = header_dict["是否临时入网"]
            formula_range = ws.range((2, temp_col), (data_rows + 1, temp_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},数据源!N:P,3,0)'
            formula_range.formula = formula

        if all(col in header_dict for col in ["整合站名与入临时入网", "所属站址名称", "是否临时入网"]):
            integrate_col = header_dict["整合站名与入临时入网"]
            site_col = header_dict["所属站址名称"]
            temp_col = header_dict["是否临时入网"]
            formula_range = ws.range((2, integrate_col), (data_rows + 1, integrate_col))
            first_temp_cell = ws.range((2, temp_col)).address.replace("$", "")
            first_site_cell = ws.range((2, site_col)).address.replace("$", "")
            formula = f'=IF(ISNA({first_temp_cell}), {first_site_cell}, {first_temp_cell})'
            formula_range.formula = formula

        if "整合站名与入临时入网" in header_dict:
            integrate_col = header_dict["整合站名与入临时入网"]
            integrate_cell = ws.range((2, integrate_col)).address.replace("$", "")
            if "分管维护员" in header_dict:
                leader_col = header_dict["分管维护员"]
                formula_range = ws.range((2, leader_col), (data_rows + 1, leader_col))
                formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,3,0)'
                formula_range.formula = formula
            if "区域" in header_dict:
                area_col = header_dict["区域"]
                formula_range = ws.range((2, area_col), (data_rows + 1, area_col))
                formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,2,0)'
                formula_range.formula = formula
            if "片区" in header_dict:
                region_col = header_dict["片区"]
                formula_range = ws.range((2, region_col), (data_rows + 1, region_col))
                formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,5,0)'
                formula_range.formula = formula

        if "是否超7天" in header_dict and "设备编码" in header_dict:
            over7_col = header_dict["是否超7天"]
            code_col = header_dict["设备编码"]
            formula_range = ws.range((2, over7_col), (data_rows + 1, over7_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=IF(ISNA(VLOOKUP({first_code_cell},网关离线清单!A:B,2,0)),"",IF(VLOOKUP({first_code_cell},网关离线清单!A:B,2,0)>=7,"是",""))'
            formula_range.formula = formula

        if "是否代维处理" in header_dict:
            dv_col = header_dict["是否代维处理"]
            ws.range((2, dv_col), (data_rows + 1, dv_col)).value = "是"

        if "核减" in header_dict and "设备编码" in header_dict:
            reduce_col = header_dict["核减"]
            code_col = header_dict["设备编码"]
            formula_range = ws.range((2, reduce_col), (data_rows + 1, reduce_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},核减清单!A:C,3,0)'
            formula_range.formula = formula

        if all(col in header_dict for col in ["核减", "分管维护员", "区域"]):
            try:
                wb.app.calculate()
                data_range = ws.used_range
                df = pd.DataFrame(data_range.value[1:], columns=data_range.value[0])
                reduce_series = df["核减"].astype(str)
                mask = (df["核减"].notna() &
                        (reduce_series.str.strip() != "") &
                        (~reduce_series.str.startswith("#", na=False)))
                clear_count = mask.sum()

                if clear_count > 0:
                    # 性能优化：批量清空分管维护员和区域列
                    leader_col_idx = header_dict["分管维护员"]
                    area_col_idx = header_dict["区域"]
                    
                    leader_values = ws.range((2, leader_col_idx), (data_rows + 1, leader_col_idx)).value
                    area_values = ws.range((2, area_col_idx), (data_rows + 1, area_col_idx)).value
                    
                    if not isinstance(leader_values, (list, tuple)):
                        leader_values = [[leader_values]] if leader_values is not None else []
                    if not isinstance(area_values, (list, tuple)):
                        area_values = [[area_values]] if area_values is not None else []
                    
                    new_leader_values = []
                    new_area_values = []
                    
                    safe_length = min(len(mask), data_rows)
                    
                    for i in range(safe_length):
                        mask_val = False
                        if i < len(mask):
                            mask_val = mask.iloc[i]
                        
                        if mask_val:
                            new_leader_values.append([""])
                            new_area_values.append([""])
                        else:
                            leader_val = ""
                            if i < len(leader_values):
                                if isinstance(leader_values[i], (list, tuple)):
                                    leader_val = leader_values[i][0] if leader_values[i] else ""
                                else:
                                    leader_val = leader_values[i] if leader_values[i] is not None else ""
                            
                            area_val = ""
                            if i < len(area_values):
                                if isinstance(area_values[i], (list, tuple)):
                                    area_val = area_values[i][0] if area_values[i] else ""
                                else:
                                    area_val = area_values[i] if area_values[i] is not None else ""
                            
                            new_leader_values.append([leader_val])
                            new_area_values.append([area_val])
                    
                    if new_leader_values:
                        ws.range((2, leader_col_idx), (data_rows + 1, leader_col_idx)).value = new_leader_values
                    if new_area_values:
                        ws.range((2, area_col_idx), (data_rows + 1, area_col_idx)).value = new_area_values

                    backup_sheet_name = "网关总清单_核减清理备份"
                    if backup_sheet_name in [s.name for s in wb.sheets]:
                        wb.sheets[backup_sheet_name].delete()
                    ws_backup = wb.sheets.add(backup_sheet_name, after=ws)
                    data_range.copy()
                    ws_backup.range("A1").paste(paste="all")
                    ws_backup.visible = False
            except Exception as e:
                pass

        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet=sheet_name)
        force_all_sheets_text_format(wb)

        wb.app.calculate()
        wb.screen_updating = True
        wb.save()
        print("网关总清单处理完成")
        return True

    except Exception as e:
        print(f"网关总清单处理失败：{str(e)}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
        if app:
            try:
                app.quit()
                app.kill()
            except:
                pass


def process_gateway_main_list():
    # update_gateway_sheet_vol()

    gateway_base_path = settings.resolve_path("spider/script/station/down/边缘网关")
    gateway_raw_path = get_latest_gateway_file(gateway_base_path)

    if not gateway_raw_path:
        print("网关总清单处理失败：边缘网关数据文件不存在")
        return

    TIME_THRESHOLD_HOURS = 24
    if not check_if_latest_file(gateway_raw_path, TIME_THRESHOLD_HOURS):
        print("网关总清单处理失败：边缘网关文件不是最新版本")
        return

    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print("网关总清单处理失败：目标Excel不存在")
        return

    df_updated = preprocess_gateway_data(gateway_raw_path)
    update_success = update_gateway_sheet_with_xlwings(
        target_excel_path=target_excel_path,
        sheet_name="网关总清单",
        df_updated=df_updated
    )
    return update_success


def update_gateway_offline_list():
    gateway_base_path = settings.resolve_path("spider/script/station/down/边缘网关")
    gateway_raw_path = get_latest_gateway_file(gateway_base_path)

    if not gateway_raw_path:
        print("网关离线清单处理失败：边缘网关数据文件不存在")
        return

    TIME_THRESHOLD_HOURS = 24
    if not check_if_latest_file(gateway_raw_path, TIME_THRESHOLD_HOURS):
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

    required_fields = ["设备编码", "设备名称", "设备当前状态", "最近离线时间",
                       "离线恢复时间", "设备入网状态", "国家行政区县",
                       "所属站址名称", "站址资源编码"]

    for field in required_fields:
        if field not in df_gateway.columns:
            df_gateway[field] = ""

    if "设备当前状态" in df_gateway.columns:
        df_gateway = df_gateway[df_gateway["设备当前状态"] == "离线"]

    df_updated = df_gateway[required_fields].copy()
    data_rows = len(df_updated)
    if data_rows == 0:
        print("网关离线清单处理完成：无离线网关数据")
        return

    app = None
    wb = None

    try:
        app = xw.App(visible=False, add_book=False)
        app.calculation = 'manual'
        app.screen_updating = False

        wb = app.books.open(target_excel_path)
        wb.display_alerts = False
        wb.screen_updating = False

        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet="网关离线清单")

        sheet_name = "网关离线清单"
        if sheet_name not in [s.name for s in wb.sheets]:
            print(f"网关离线清单处理失败：目标Sheet「{sheet_name}」不存在")
            return

        ws = wb.sheets[sheet_name]

        start_row = 2
        last_row = ws.cells.last_cell.row
        if last_row >= start_row:
            ws.range(f"A{start_row}:ZZ{last_row}").clear_contents()

        header_range = ws.range("1:1").value
        header_dict = {col: idx + 1 for idx, col in enumerate(header_range) if col}

        missing_required = [f for f in required_fields if f not in header_dict]
        if missing_required:
            print(f"网关离线清单处理失败：表头中缺少必填字段 {missing_required}")
            return

        # 性能优化：一次性写入所有数据，而不是逐列写入
        sorted_cols = []
        for field in required_fields:
            if field in header_dict:
                sorted_cols.append(field)
        
        if sorted_cols:
            sorted_col_indices = [header_dict[field] for field in sorted_cols]
            min_col, max_col = min(sorted_col_indices), max(sorted_col_indices)
            
            data_matrix = [[""] * (max_col - min_col + 1) for _ in range(data_rows)]
            
            for col_idx, field in enumerate(sorted_cols):
                if field in df_updated.columns:
                    target_col_pos = sorted_col_indices[col_idx] - min_col
                    col_data = df_updated[field].tolist()
                    for row_idx in range(data_rows):
                        data_matrix[row_idx][target_col_pos] = col_data[row_idx]
            
            ws.range((2, min_col), (data_rows + 1, max_col)).value = data_matrix

        if "离线天数" in header_dict and "最近离线时间" in header_dict:
            offline_days_col = header_dict["离线天数"]
            recent_offline_col = header_dict["最近离线时间"]
            first_offline_cell = ws.range((2, recent_offline_col)).address.replace("$", "")
            formula = f'=IF({first_offline_cell}="","",TODAY()-INT({first_offline_cell}))'
            formula_range = ws.range((2, offline_days_col), (data_rows + 1, offline_days_col))
            formula_range.formula = formula

        if "是否代维处理" in header_dict and "设备编码" in header_dict:
            dv_col = header_dict["是否代维处理"]
            code_col = header_dict["设备编码"]
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},网关总清单!$A$2:$K$8876,11,0)'
            formula_range = ws.range((2, dv_col), (data_rows + 1, dv_col))
            formula_range.formula = formula

        if "分管维护员" in header_dict and "设备名称" in header_dict:
            leader_col = header_dict["分管维护员"]
            name_col = header_dict["设备名称"]
            first_name_cell = ws.range((2, name_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_name_cell},网关总清单!B:I,7,0)'
            formula_range = ws.range((2, leader_col), (data_rows + 1, leader_col))
            formula_range.formula = formula

        if "临时入网实际安装站" in header_dict and "设备编码" in header_dict:
            temp_site_col = header_dict["临时入网实际安装站"]
            code_col = header_dict["设备编码"]
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},网关总清单!$A$2:$F$8876,6,0)'
            formula_range = ws.range((2, temp_site_col), (data_rows + 1, temp_site_col))
            formula_range.formula = formula

        if "备注" in header_dict and "设备编码" in header_dict:
            remark_col = header_dict["备注"]
            code_col = header_dict["设备编码"]
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},核减清单!$A$2:$C$537,3,0)'
            formula_range = ws.range((2, remark_col), (data_rows + 1, remark_col))
            formula_range.formula = formula

        if "片区" in header_dict and "所属站址名称" in header_dict:
            region_col = header_dict["片区"]
            site_col = header_dict["所属站址名称"]
            first_site_cell = ws.range((2, site_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_site_cell},数据源!V:Z,5,0)'
            formula_range = ws.range((2, region_col), (data_rows + 1, region_col))
            formula_range.formula = formula

        if "考核核减" in header_dict and "备注" in header_dict:
            check_col = header_dict["考核核减"]
            remark_col = header_dict["备注"]
            first_remark_cell = ws.range((2, remark_col)).address.replace("$", "")
            formula = f'=IF({first_remark_cell}="","","核减")'
            formula_range = ws.range((2, check_col), (data_rows + 1, check_col))
            formula_range.formula = formula

        if "是否超7天" in header_dict and "离线天数" in header_dict:
            over7_col = header_dict["是否超7天"]
            offline_days_col = header_dict["离线天数"]
            first_offline_days_cell = ws.range((2, offline_days_col)).address.replace("$", "")
            formula = f'=IF({first_offline_days_cell}="","",IF({first_offline_days_cell}>=7,"是","否"))'
            formula_range = ws.range((2, over7_col), (data_rows + 1, over7_col))
            formula_range.formula = formula

        if "国家行政区县" in header_dict and "设备名称" in header_dict:
            region_col = header_dict["国家行政区县"]
            name_col = header_dict["设备名称"]
            first_name_cell = ws.range((2, name_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_name_cell},网关总清单!$B:$J,8,0)'
            formula_range = ws.range((2, region_col), (data_rows + 1, region_col))
            formula_range.formula = formula

        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet="网关离线清单")
        force_all_sheets_text_format(wb)

        wb.app.calculate()
        wb.screen_updating = True
        wb.save()
        print("网关离线清单处理完成")

    except Exception as e:
        print(f"网关离线清单处理失败：{str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        if app:
            app.calculation = 'automatic'
            app.screen_updating = True
        if wb:
            try:
                wb.close()
            except:
                pass
        if app:
            try:
                app.quit()
                app.kill()
            except:
                pass


if __name__ == '__main__':
    process_gateway_main_list()
    update_gateway_offline_list()
