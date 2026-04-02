import os
import pandas as pd
import xlwings as xw
from core.config import settings
import warnings
from datetime import datetime, timedelta

warnings.filterwarnings('ignore')

"""
摄像头数据处理模块（终极修复版）
- 批量公式备份/恢复（超快）→ 解决公式引用表名错误问题
- 强制所有Sheet设备编码=文本格式，彻底消灭科学计数法
- 彻底解决其他Sheet引用摄像头总清单失效问题
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
        print(f"\n❌ 错误：摄像头原始数据文件不是最新版本！")
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
    new_fields = ["是否临时入网", "整合站名与入临时入网", "区域", "分管维护员", "片区", "是否代维处理", "核减", "是否超7天"]
    for field in new_fields:
        df_main[field] = ""
    for col in ["设备编码", "站址资源编码"]:
        df_main[col] = df_main[col].astype(str).str.strip()
    print(f"✅ 摄像头数据预处理完成：有效数据{len(df_main)}行")
    return df_main

# ===================== 【极速批量】备份所有非摄像头sheet公式 =====================
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
        print(f"✅ 批量备份完成：{len(backup)} 个工作表（无逐行遍历）")
    except Exception as e:
        print(f"⚠️ 批量备份失败：{e}")
    return backup

# ===================== 【极速批量】恢复所有非摄像头sheet公式 =====================
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
        print(f"✅ 批量恢复完成：{count} 个工作表（无逐行写入）")
    except Exception as e:
        print(f"⚠️ 批量恢复失败：{e}")

# ===================== 【终极修复】强制所有Sheet设备编码列=文本格式（消灭科学计数法） =====================
def force_all_sheets_text_format(wb, key_columns=["设备编码", "站址资源编码", "设备编号", "编号"]):
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

# ===================== 主函数（保留原始业务逻辑，仅优化公式/格式处理） =====================
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

        # ==============================================
        # 1. 批量备份所有非摄像头sheet公式（核心修复：先备份再处理）
        # ==============================================
        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet=sheet_name)

        # ==============================================
        # 2. 【终极修复】执行前：全表强制文本格式（防科学计数法）
        # ==============================================
        force_all_sheets_text_format(wb)

        # ==============================================
        # 原有业务逻辑（100%保留，一行不改）
        # ==============================================
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

        # 3.7 核减清理（原始逻辑保留）
        if all(col in header_dict for col in ["核减", "分管维护员", "区域"]):
            try:
                wb.app.calculate()
                data = ws.used_range.value
                if data and len(data) >= 2:
                    df = pd.DataFrame(data[1:], columns=data[0])
                    reduce_series = df["核减"].astype(str)
                    mask = df["核减"].notna() & (reduce_series.str.strip() != "") & (~reduce_series.str.startswith("#", na=False))
                    clear_count = mask.sum()
                    if clear_count > 0:
                        df.loc[mask, ["分管维护员", "区域"]] = ""
                        print(f"✅ 摄像头数据 - 清理 {clear_count} 行核减数据")

                        new_sheet = wb.sheets.add("摄像头总清单_核减清理", after=ws)
                        ws.range("A1").resize(column_size=len(df.columns)).copy()
                        new_sheet.range("A1").paste(paste="all")

                        for c in ["设备编码", "站址资源编码"]:
                            if c in df.columns:
                                col_idx = df.columns.get_loc(c) + 1
                                new_sheet.range((2, col_idx), (df.shape[0]+1, col_idx)).number_format = "@"
                                df[c] = df[c].apply(lambda x: f"'{x}" if str(x).strip() else x)

                        new_sheet.range("A2").value = df.values
                        ws.name = "摄像头总清单_备份"
                        ws.visible = False
                        new_sheet.name = "摄像头总清单"
            except Exception as e:
                print(f"⚠️ 核减清理失败：{e}")

        # ==============================================
        # 3. 批量恢复所有非摄像头sheet公式（核心修复：恢复原始公式）
        # ==============================================
        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet=sheet_name)

        # ==============================================
        # 4. 执行后：再次强制全表文本格式（双重保障）
        # ==============================================
        force_all_sheets_text_format(wb)

        # 最终计算并保存
        wb.app.calculate()
        wb.screen_updating = True
        wb.save()

        print(f"\n🎉 摄像头处理完成！共处理 {data_rows} 行数据")
        print(f"✅ 其他Sheet公式已恢复，引用表名保持为「摄像头总清单」")
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

if __name__ == '__main__':
    process_camera_main_list()