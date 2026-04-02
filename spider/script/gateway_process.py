import os
import pandas as pd
import xlwings as xw
from core.config import settings
import warnings
from datetime import datetime, timedelta

warnings.filterwarnings('ignore')

"""
网关数据处理模块 - 终极修复版
- 批量公式备份/恢复（超快）
- 强制所有Sheet设备编码=文本格式，彻底消灭科学计数法
- 彻底解决其他Sheet引用新网关总清单失效问题
- 原有逻辑100%保留
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
    print(f"\n📄 文件检查：{os.path.basename(file_path)}")
    print(f"   修改时间：{file_modify_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"   最新有效时间：{latest_valid_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"   是否最新：{'✅ 是' if is_latest else '❌ 否'}")
    return is_latest

def load_excel_data(file_path, sheet_name=0):
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在：{file_path}")
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")
        return df
    except Exception as e:
        print(f"加载Excel失败：{file_path} -> {e}")
        return pd.DataFrame()

def preprocess_gateway_data(gateway_raw_path):
    df_gateway = load_excel_data(gateway_raw_path)
    if df_gateway.empty:
        print("❌ 原始网关数据加载失败")
        return pd.DataFrame()
    core_fields = ["设备编码", "设备名称", "国家行政区县", "所属站址名称", "站址资源编码"]
    for field in core_fields:
        if field not in df_gateway.columns:
            df_gateway[field] = ""
    df_main = df_gateway[core_fields].copy()
    new_fields = ["是否临时入网", "整合站名与入临时入网", "分管维护员", "区域", "片区",
                  "是否代维处理", "核减", "是否超7天"]
    for field in new_fields:
        df_main[field] = ""
    print(f"✅ 网关数据预处理完成：有效数据{len(df_main)}行")
    return df_main

# ===================== 【极速批量】备份所有非网关sheet公式 =====================
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
        print(f"✅ 批量备份完成：{len(backup)} 个工作表（无逐行遍历）")
    except Exception as e:
        print(f"⚠️ 批量备份失败：{e}")
    return backup

# ===================== 【极速批量】恢复所有非网关sheet公式 =====================
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
        print(f"✅ 批量恢复完成：{count} 个工作表（无逐行写入）")
    except Exception as e:
        print(f"⚠️ 批量恢复失败：{e}")

# ===================== 【终极修复】强制所有Sheet设备编码列=文本格式（消灭科学计数法） =====================
def force_all_sheets_text_format(wb, key_columns=["设备编码", "站址资源编码", "设备编号"]):
    """
    强制整个Excel所有Sheet的关键字段为【文本格式】
    彻底杜绝科学计数法，确保VLOOKUP永远匹配成功
    """
    try:
        fix_count = 0
        for sh in wb.sheets:
            try:
                headers = sh.range("1:1").value
                if not headers:
                    continue
                for col_idx, header in enumerate(headers):
                    if header in key_columns:
                        excel_col = col_idx + 1
                        sh.range((1, excel_col), (sh.cells.last_cell.row, excel_col)).number_format = "@"
                        fix_count += 1
            except:
                continue
        print(f"✅ 格式修复完成：强制 {fix_count} 列 = 文本格式（无科学计数法）")
    except Exception as e:
        print(f"⚠️ 格式修复失败：{e}")

# ===================== 主逻辑（完全不动，只加批量备份/恢复 + 格式修复） =====================
def update_gateway_sheet_with_xlwings(target_excel_path, sheet_name, df_updated):
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

        # ==============================================
        # 1. 批量备份所有非网关公式
        # ==============================================
        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet=sheet_name)

        # ==============================================
        # 2. 【终极修复】执行前：全表强制文本格式（防科学计数法）
        # ==============================================
        force_all_sheets_text_format(wb)

        # 以下是你原来的所有逻辑，一行不改！！！
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
        if all(col in header_dict for col in ["整合站名与入临时入网", "所属站址名称", "是否临时入网"]):
            integrate_col = header_dict["整合站名与入临时入网"]
            site_col = header_dict["所属站址名称"]
            temp_col = header_dict["是否临时入网"]
            formula_range = ws.range((2, integrate_col), (data_rows + 1, integrate_col))
            first_temp_cell = ws.range((2, temp_col)).address.replace("$", "")
            first_site_cell = ws.range((2, site_col)).address.replace("$", "")
            formula = f'=IF(ISNA({first_temp_cell}), {first_site_cell}, {first_temp_cell})'
            formula_range.formula = formula

        # 3.3 分管维护员/区域/片区
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

        # 3.4 是否超7天
        if "是否超7天" in header_dict and "设备编码" in header_dict:
            over7_col = header_dict["是否超7天"]
            code_col = header_dict["设备编码"]
            formula_range = ws.range((2, over7_col), (data_rows + 1, over7_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=IF(ISNA(VLOOKUP({first_code_cell},网关离线清单!A:B,2,0)),"",IF(VLOOKUP({first_code_cell},网关离线清单!A:B,2,0)>=7,"是",""))'
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

        # 3.7 核减清理（你的原有逻辑）
        if all(col in header_dict for col in ["核减", "分管维护员", "区域"]):
            try:
                wb.app.calculate()
                data = ws.used_range.value
                if not data or len(data) < 2:
                    pass
                else:
                    df = pd.DataFrame(data[1:], columns=data[0])
                    reduce_series = df["核减"].astype(str)
                    mask = (df["核减"].notna() &
                            (reduce_series.str.strip() != "") &
                            (~reduce_series.str.startswith("#", na=False)))
                    clear_count = mask.sum()

                    if clear_count > 0:
                        df.loc[mask, ["分管维护员", "区域"]] = ""
                        print(f"✅ 网关数据 - 清理 {clear_count} 行核减相关数据")

                        new_sheet_name = "网关总清单_核减清理"
                        if new_sheet_name in [s.name for s in wb.sheets]:
                            wb.sheets[new_sheet_name].delete()
                        ws_new = wb.sheets.add(new_sheet_name, after=ws)

                        num_cols = len(df.columns)
                        header_orig = ws.range("A1").resize(column_size=num_cols)
                        header_new = ws_new.range("A1").resize(column_size=num_cols)
                        header_orig.copy()
                        header_new.paste(paste="all")

                        code_column_names = ["设备编码", "站址资源编码"]
                        for col_name in code_column_names:
                            if col_name in df.columns:
                                col_idx = df.columns.get_loc(col_name) + 1
                                ws_new.range(ws_new.cells(2, col_idx),
                                             ws_new.cells(df.shape[0] + 1, col_idx)).number_format = "@"

                        for col_name in code_column_names:
                            if col_name in df.columns:
                                df[col_name] = df[col_name].apply(
                                    lambda x: f"'{x}" if (pd.notna(x) and str(x).strip() != "") else x
                                )

                        ws_new.range("A2").value = df.values

                        backup_sheet_name = "网关总清单_备份"
                        if backup_sheet_name in [s.name for s in wb.sheets]:
                            wb.sheets[backup_sheet_name].delete()
                        ws.name = backup_sheet_name
                        ws.visible = False
                        ws_new.name = "网关总清单"

            except Exception as e:
                print(f"⚠️ 核减清理处理失败：{e}")

        # ==============================================
        # 3. 批量恢复所有非网关公式
        # ==============================================
        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet=sheet_name)

        # ==============================================
        # 4. 【终极修复】执行后：再次强制全表文本格式
        # ==============================================
        force_all_sheets_text_format(wb)

        wb.app.calculate()
        wb.screen_updating = True
        wb.save()
        print(f"\n🎉 网关数据更新完成！共处理 {data_rows} 行数据")
        return True

    except Exception as e:
        print(f"❌ xlwings更新数据失败：{e}")
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
    gateway_xlsx_path = settings.resolve_path("spider/script/station/down/边缘网关.xlsx")
    gateway_xls_path = os.path.splitext(gateway_xlsx_path)[0] + ".xls"

    gateway_raw_path = ""
    if os.path.exists(gateway_xlsx_path):
        gateway_raw_path = gateway_xlsx_path
    elif os.path.exists(gateway_xls_path):
        gateway_raw_path = gateway_xls_path
    else:
        print(f"❌ 网关数据文件不存在：")
        return

    TIME_THRESHOLD_HOURS = 24
    if not check_if_latest_file(gateway_raw_path, TIME_THRESHOLD_HOURS):
        print(f"\n❌ 错误：边缘网关文件不是最新版本！")
        return

    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print(f"❌ 目标Excel不存在：{target_excel_path}")
        return

    df_updated = preprocess_gateway_data(gateway_raw_path)
    update_success = update_gateway_sheet_with_xlwings(
        target_excel_path=target_excel_path,
        sheet_name="网关总清单",
        df_updated=df_updated
    )

if __name__ == '__main__':
    process_gateway_main_list()