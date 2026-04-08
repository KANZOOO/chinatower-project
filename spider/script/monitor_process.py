import os
import pandas as pd
import xlwings as xw
from core.config import settings
import warnings
from datetime import datetime, timedelta

warnings.filterwarnings('ignore')

"""
摄像头数据处理模块（精简输出版）
- 批量公式备份/恢复（超快）→ 解决公式引用表名错误问题
- 强制所有Sheet设备编码=文本格式，彻底消灭科学计数法
- 彻底解决其他Sheet引用摄像头总清单失效问题
- 核减清理逻辑优化：保留原Sheet公式，仅清理指定单元格数据
- 原有逻辑100%保留，仅精简输出信息
"""


def get_file_modify_time(file_path):
    try:
        modify_timestamp = os.path.getmtime(file_path)
        modify_time = datetime.fromtimestamp(modify_timestamp)
        return modify_time
    except Exception as e:
        print(f"获取文件修改时间失败：{file_path} -> {e}")
        return None


def check_if_latest_file(file_path, time_threshold_hours=24):
    if not os.path.exists(file_path):
        return False
    file_modify_time = get_file_modify_time(file_path)
    if not file_modify_time:
        return False
    latest_valid_time = datetime.now() - timedelta(hours=time_threshold_hours)
    is_latest = file_modify_time >= latest_valid_time
    return is_latest


def load_excel_data(file_path, sheet_name=0):
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在：{file_path}")
        dtype_spec = {"设备编码": str, "站址资源编码": str}
        df_temp = pd.read_excel(file_path, sheet_name=sheet_name, nrows=0)
        valid_dtype = {k: v for k, v in dtype_spec.items() if k in df_temp.columns}
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=valid_dtype, na_filter=False)
        df = df.fillna("")
        for col in ["设备编码", "站址资源编码"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        return df
    except Exception as e:
        print(f"加载Excel失败：{file_path} -> {e}")
        return pd.DataFrame()


def preprocess_camera_data(camera_raw_path):
    TIME_THRESHOLD_HOURS = 24
    if not check_if_latest_file(camera_raw_path, TIME_THRESHOLD_HOURS):
        print("❌ 摄像头原始数据文件不是最新版本")
        return pd.DataFrame()
    df_camera = load_excel_data(camera_raw_path)
    if df_camera.empty:
        print("❌ 原始摄像头数据加载失败")
        return pd.DataFrame()
    core_fields = ["设备编码", "设备名称", "摄像头类型", "摄像头安装位置", "站址名称", "站址资源编码"]
    for field in core_fields:
        if field not in df_camera.columns:
            df_camera[field] = ""
    df_main = df_camera[core_fields].copy()
    new_fields = ["是否临时入网", "整合站名与入临时入网", "区域", "分管维护员", "片区", "是否代维处理", "核减",
                  "是否超7天"]
    for field in new_fields:
        df_main[field] = ""
    for col in ["设备编码", "站址资源编码"]:
        df_main[col] = df_main[col].astype(str).str.strip()
    return df_main


def batch_backup_all_sheets_formulas(wb, exclude_sheet="摄像头总清单"):
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
        print(f"批量备份失败：{e}")
    return backup


def batch_restore_all_sheets_formulas(wb, backup, exclude_sheet="摄像头总清单"):
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
        print(f"批量恢复失败：{e}")


def force_all_sheets_text_format(wb, key_columns=["设备编码", "站址资源编码", "设备编号", "入网设备编码", "站址编码",
                                                  "编号"]):
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
        print(f"格式修复失败：{e}")


def update_camera_sheet_with_xlwings(target_excel_path, sheet_name, df_updated):
    if df_updated.empty:
        print("❌ 无更新数据，终止写入")
        return False

    app = None
    wb = None
    formula_backup = {}

    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(target_excel_path)
        wb.display_alerts = False
        wb.screen_updating = False

        # 批量备份公式
        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet=sheet_name)

        # 强制文本格式
        force_all_sheets_text_format(wb)

        # 原有业务逻辑
        if sheet_name not in [s.name for s in wb.sheets]:
            print(f"❌ 目标Sheet「{sheet_name}」不存在")
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
            print("❌ 无匹配的列名，无法写入数据")
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

        for col_name in valid_cols:
            col_idx = header_dict[col_name]
            col_data = df_valid[col_name].tolist()
            ws.range((2, col_idx), (data_rows + 1, col_idx)).value = [[val] for val in col_data]

        # 3.1 是否临时入网
        if "是否临时入网" in header_dict and "设备编码" in header_dict:
            code_col = header_dict["设备编码"]
            temp_col = header_dict["是否临时入网"]
            formula_range = ws.range((2, temp_col), (data_rows + 1, temp_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},数据源!N:P,3,0)'
            formula_range.formula = formula

        # 3.2 整合站名
        if all(col in header_dict for col in ["整合站名与入临时入网", "站址名称", "是否临时入网"]):
            integrate_col = header_dict["整合站名与入临时入网"]
            site_col = header_dict["站址名称"]
            temp_col = header_dict["是否临时入网"]
            formula_range = ws.range((2, integrate_col), (data_rows + 1, integrate_col))
            first_temp_cell = ws.range((2, temp_col)).address.replace("$", "")
            first_site_cell = ws.range((2, site_col)).address.replace("$", "")
            formula = f'=IF(ISNA({first_temp_cell}), {first_site_cell}, {first_temp_cell})'
            formula_range.formula = formula

        # 3.3 区域/维护员/片区
        if "整合站名与入临时入网" in header_dict:
            integrate_cell = ws.range((2, header_dict["整合站名与入临时入网"])).address.replace("$", "")
            if "区域" in header_dict:
                formula_range = ws.range((2, header_dict["区域"]), (data_rows + 1, header_dict["区域"]))
                formula_range.formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,2,0)'
            if "分管维护员" in header_dict:
                formula_range = ws.range((2, header_dict["分管维护员"]), (data_rows + 1, header_dict["分管维护员"]))
                formula_range.formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,3,0)'
            if "片区" in header_dict:
                formula_range = ws.range((2, header_dict["片区"]), (data_rows + 1, header_dict["片区"]))
                formula_range.formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,5,0)'

        # 3.4 是否超7天
        if "是否超7天" in header_dict and "设备编码" in header_dict:
            over7_col = header_dict["是否超7天"]
            code_col = header_dict["设备编码"]
            formula_range = ws.range((2, over7_col), (data_rows + 1, over7_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=IF(ISNA(VLOOKUP({first_code_cell},摄像头离线清单!A:B,2,0)),"",IF(VLOOKUP({first_code_cell},摄像头离线清单!A:B,2,0)>=7,"是",""))'
            formula_range.formula = formula

        # 3.5 是否代维处理
        if "是否代维处理" in header_dict:
            dv_col = header_dict["是否代维处理"]
            ws.range((2, dv_col), (data_rows + 1, dv_col)).value = "是"

        # 3.6 核减
        if "核减" in header_dict and "设备编码" in header_dict:
            reduce_col = header_dict["核减"]
            code_col = header_dict["设备编码"]
            formula_range = ws.range((2, reduce_col), (data_rows + 1, reduce_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},核减清单!A:C,3,0)'
            formula_range.formula = formula

        # 3.7 核减清理
        if all(col in header_dict for col in ["核减", "分管维护员", "区域"]):
            try:
                wb.app.calculate()
                data_range = ws.used_range
                df = pd.DataFrame(data_range.value[1:], columns=data_range.value[0])
                df = df.fillna("")
                reduce_col = df["核减"].astype(str)

                mask = (df["核减"].notna() &
                        (reduce_col.str.strip() != "") &
                        (~reduce_col.str.startswith("#", na=False)))
                clear_count = mask.sum()

                if mask.any():
                    for row_idx in mask[mask].index:
                        row_num = row_idx + 2
                        if "分管维护员" in header_dict:
                            ws.range((row_num, header_dict["分管维护员"])).value = ""
                        if "区域" in header_dict:
                            ws.range((row_num, header_dict["区域"])).value = ""

                    # 备份
                    backup_sheet = "摄像头总清单_核减备份"
                    if backup_sheet in [s.name for s in wb.sheets]:
                        wb.sheets[backup_sheet].delete()
                    ws_bak = wb.sheets.add(backup_sheet, after=ws)
                    data_range.copy()
                    ws_bak.range("A1").paste(paste="all")
                    ws_bak.visible = False

            except Exception as e:
                print(f"核减批量清理失败：{e}")

        # 恢复公式
        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet=sheet_name)

        # 再次强制文本格式
        force_all_sheets_text_format(wb)

        # 保存
        wb.app.calculate()
        wb.screen_updating = True
        wb.save()

        print(f"✅ 摄像头总清单处理完成：{data_rows} 行")
        return True

    except Exception as e:
        print(f"❌ 执行失败：{e}")
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


def update_camera_offline_list(target_excel_path):
    """更新摄像头离线清单"""
    # 1. 确定文件路径
    gateway_xlsx_path = settings.resolve_path("spider/script/station/down/边缘网关.xlsx")
    gateway_xls_path = os.path.splitext(gateway_xlsx_path)[0] + ".xls"
    gateway_raw_path = gateway_xlsx_path if os.path.exists(gateway_xlsx_path) else gateway_xls_path

    camera_xlsx_path = settings.resolve_path("spider/script/station/down/边缘摄像头.xlsx")
    camera_xls_path = os.path.splitext(camera_xlsx_path)[0] + ".xls"
    camera_raw_path = camera_xlsx_path if os.path.exists(camera_xlsx_path) else camera_xls_path

    if not os.path.exists(gateway_raw_path):
        print(f"❌ 边缘网关文件不存在")
        return

    if not os.path.exists(camera_raw_path):
        print(f"❌ 边缘摄像头文件不存在")
        return

    if not os.path.exists(target_excel_path):
        print(f"❌ 目标模板不存在")
        return

    # 2. 加载数据
    df_gateway = load_excel_data(gateway_raw_path)
    df_camera = load_excel_data(camera_raw_path)

    if df_gateway.empty or df_camera.empty:
        print("❌ 数据加载失败")
        return

    # 3. 匹配数据
    gateway_codes = set(df_gateway["设备编码"].astype(str).str.strip())
    df_camera["设备编码"] = df_camera["设备编码"].astype(str).str.strip()
    df_matched = df_camera[df_camera["设备编码"].isin(gateway_codes)].copy()

    # 4. 保留字段
    keep_fields = ["设备编码", "设备名称", "通道当前状态", "通道最近离线时间", "摄像头安装位置", "站址名称",
                   "站址资源编码"]
    for field in keep_fields:
        if field not in df_matched.columns:
            df_matched[field] = ""

    df_processed = df_matched[keep_fields].copy()

    # 5. 筛选离线数据
    if "通道当前状态" in df_processed.columns:
        df_processed = df_processed[df_processed["通道当前状态"] == "离线"]

    data_rows = len(df_processed)
    if data_rows == 0:
        print("⚠️ 没有离线摄像头数据")
        return

    # 6. 更新Excel
    app = None
    wb = None

    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(target_excel_path)
        wb.display_alerts = False
        wb.screen_updating = False

        sheet_name = "摄像头离线清单"
        if sheet_name not in [s.name for s in wb.sheets]:
            print(f"❌ 目标Sheet「{sheet_name}」不存在")
            return

        ws = wb.sheets[sheet_name]

        # 清空数据
        start_row = 2
        last_row = ws.cells.last_cell.row
        if last_row >= start_row:
            ws.range(f"A{start_row}:ZZ{last_row}").clear_contents()

        # 获取表头
        header_range = ws.range("1:1").value
        header_dict = {col: idx + 1 for idx, col in enumerate(header_range) if col}

        # 写入数据
        for field in keep_fields:
            if field in header_dict:
                col_idx = header_dict[field]
                col_data = df_processed[field].tolist()
                ws.range((2, col_idx), (data_rows + 1, col_idx)).value = [[val] for val in col_data]

        # 生成公式字段
        # 11.1 离线天数公式
        if "离线天数" in header_dict and "通道最近离线时间" in header_dict:
            offline_days_col = header_dict["离线天数"]
            recent_offline_col = header_dict["通道最近离线时间"]
            recent_offline_col_letter = ws.range((1, recent_offline_col)).address.replace("$1", "").replace("$", "")
            formula = f'=IF({recent_offline_col_letter}2="","",TODAY()-INT({recent_offline_col_letter}2))'
            ws.range((2, offline_days_col)).formula = formula
            if data_rows > 1:
                ws.range((2, offline_days_col)).autofill(
                    ws.range((2, offline_days_col), (data_rows + 1, offline_days_col)))

        # 11.2 区域公式
        if "区域" in header_dict:
            area_col = header_dict["区域"]
            formula = f'=VLOOKUP(B2,摄像头总清单!B:K,8,0)'
            ws.range((2, area_col)).formula = formula
            if data_rows > 1:
                ws.range((2, area_col)).autofill(ws.range((2, area_col), (data_rows + 1, area_col)))

        # 11.3 网关在线情况公式
        if "网关在线情况" in header_dict and "设备编码" in header_dict:
            gateway_status_col = header_dict["网关在线情况"]
            code_col = header_dict["设备编码"]
            code_col_letter = ws.range((1, code_col)).address.replace("$1", "").replace("$", "")
            formula = f'=IF(ISNA(VLOOKUP({code_col_letter}2,网关离线清单!A:A,1,0)),"在线","网关离线")'
            ws.range((2, gateway_status_col)).formula = formula
            if data_rows > 1:
                ws.range((2, gateway_status_col)).autofill(
                    ws.range((2, gateway_status_col), (data_rows + 1, gateway_status_col)))

        # 11.4 是否代维处理公式
        if "是否代维处理" in header_dict and "设备编码" in header_dict:
            dv_col = header_dict["是否代维处理"]
            code_col = header_dict["设备编码"]
            code_col_letter = ws.range((1, code_col)).address.replace("$1", "").replace("$", "")
            formula = f'=VLOOKUP({code_col_letter}2,摄像头总清单!$A$2:$L$44393,12,0)'
            ws.range((2, dv_col)).formula = formula
            if data_rows > 1:
                ws.range((2, dv_col)).autofill(ws.range((2, dv_col), (data_rows + 1, dv_col)))

        # 11.5 分管维护员公式
        if "分管维护员" in header_dict:
            leader_col = header_dict["分管维护员"]
            formula = f'=VLOOKUP(B2,摄像头总清单!B:K,9,0)'
            ws.range((2, leader_col)).formula = formula
            if data_rows > 1:
                ws.range((2, leader_col)).autofill(ws.range((2, leader_col), (data_rows + 1, leader_col)))

        # 11.6 临时入网实际安装站公式
        if "临时入网实际安装站" in header_dict and "设备编码" in header_dict:
            temp_site_col = header_dict["临时入网实际安装站"]
            code_col = header_dict["设备编码"]
            code_col_letter = ws.range((1, code_col)).address.replace("$1", "").replace("$", "")
            formula = f'=VLOOKUP({code_col_letter}2,摄像头总清单!$A$2:$G$44393,7,0)'
            ws.range((2, temp_site_col)).formula = formula
            if data_rows > 1:
                ws.range((2, temp_site_col)).autofill(ws.range((2, temp_site_col), (data_rows + 1, temp_site_col)))

        # 11.7 备注公式
        if "备注" in header_dict and "设备编码" in header_dict:
            remark_col = header_dict["备注"]
            code_col = header_dict["设备编码"]
            code_col_letter = ws.range((1, code_col)).address.replace("$1", "").replace("$", "")
            formula = f'=VLOOKUP({code_col_letter}2,核减清单!$A$1:$C$10537,3,0)'
            ws.range((2, remark_col)).formula = formula
            if data_rows > 1:
                ws.range((2, remark_col)).autofill(ws.range((2, remark_col), (data_rows + 1, remark_col)))

        # 11.8 片区公式
        if "片区" in header_dict:
            region_col = header_dict["片区"]
            formula = f'=VLOOKUP(B2,摄像头总清单!B:K,10,0)'
            ws.range((2, region_col)).formula = formula
            if data_rows > 1:
                ws.range((2, region_col)).autofill(ws.range((2, region_col), (data_rows + 1, region_col)))

        # 11.9 考核核减公式
        if "考核核减" in header_dict and "备注" in header_dict:
            check_col = header_dict["考核核减"]
            remark_col = header_dict["备注"]
            remark_col_letter = ws.range((1, remark_col)).address.replace("$1", "").replace("$", "")
            formula = f'=IF({remark_col_letter}2="","","核减")'
            ws.range((2, check_col)).formula = formula
            if data_rows > 1:
                ws.range((2, check_col)).autofill(ws.range((2, check_col), (data_rows + 1, check_col)))

        # 11.10 是否超7天公式
        if "是否超7天" in header_dict and "离线天数" in header_dict:
            over7_col = header_dict["是否超7天"]
            offline_days_col = header_dict["离线天数"]
            formula_range = ws.range((2, over7_col), (data_rows + 1, over7_col))
            first_offline_days_cell = ws.range((2, offline_days_col)).address.replace("$", "")
            formula = f'=IF({first_offline_days_cell}="","",IF({first_offline_days_cell}>=7,"是","否"))'
            formula_range.formula = formula

        # 强制文本格式
        force_all_sheets_text_format(wb)

        # 保存
        wb.app.calculate()
        wb.screen_updating = True
        wb.save()
        print(f"✅ 摄像头离线清单处理完成：{data_rows} 行")

    except Exception as e:
        print(f"❌ 离线清单更新失败：{e}")
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


def process_camera_main_list():
    camera_xlsx_path = settings.resolve_path("spider/script/station/down/边缘摄像头.xlsx")
    camera_xls_path = os.path.splitext(camera_xlsx_path)[0] + ".xls"
    camera_raw_path = camera_xlsx_path if os.path.exists(camera_xlsx_path) else camera_xls_path

    if not os.path.exists(camera_raw_path):
        print(f"❌ 摄像头文件不存在")
        return

    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print(f"❌ 目标模板不存在")
        return

    df_updated = preprocess_camera_data(camera_raw_path)
    update_camera_sheet_with_xlwings(
        target_excel_path=target_excel_path,
        sheet_name="摄像头总清单",
        df_updated=df_updated
    )
    update_camera_offline_list(target_excel_path)


# 修复原代码中重复调用的问题
def process_camera_offline_list():
    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    update_camera_offline_list(target_excel_path)


if __name__ == '__main__':
    process_camera_main_list()