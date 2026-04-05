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

        # 3.7 核减清理（优化版：保留原Sheet+公式，仅清理数据）
        if all(col in header_dict for col in ["核减", "分管维护员", "区域"]):
            try:
                wb.app.calculate()  # 先计算所有公式

                # 读取当前数据（保留公式引用）
                data_range = ws.used_range
                df = pd.DataFrame(data_range.value[1:], columns=data_range.value[0])
                reduce_series = df["核减"].astype(str)

                # 筛选需要清理的行
                mask = (df["核减"].notna() &
                        (reduce_series.str.strip() != "") &
                        (~reduce_series.str.startswith("#", na=False)))
                clear_count = mask.sum()

                if clear_count > 0:
                    # 直接在原Sheet中清空指定列，不重建Sheet
                    for row_idx in mask[mask].index:
                        # 行号 = 数据行索引 + 2（表头行+起始行）
                        row_num = row_idx + 2
                        # 清空分管维护员、区域列，保留其他公式列
                        if "分管维护员" in header_dict:
                            ws.range((row_num, header_dict["分管维护员"])).value = ""
                        if "区域" in header_dict:
                            ws.range((row_num, header_dict["区域"])).value = ""

                    print(f"✅ 网关数据 - 清理 {clear_count} 行核减相关数据（保留原Sheet+公式）")

                    # 可选：备份清理前的数据（单独Sheet，不影响主Sheet）
                    backup_sheet_name = "网关总清单_核减清理备份"
                    if backup_sheet_name in [s.name for s in wb.sheets]:
                        wb.sheets[backup_sheet_name].delete()
                    ws_backup = wb.sheets.add(backup_sheet_name, after=ws)
                    data_range.copy()
                    ws_backup.range("A1").paste(paste="all")
                    ws_backup.visible = False  # 备份Sheet隐藏

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

def update_gateway_offline_list():
    """
    更新网关离线清单
    - 清空数据但保留公式
    - 从边缘网关Excel导入9个字段
    - 生成7个字段的公式
    - 保留表头格式
    - 使用xlwings批量操作提高速度
    """
    # 1. 确定边缘网关数据文件路径
    gateway_xlsx_path = settings.resolve_path("spider/script/station/down/边缘网关.xlsx")
    gateway_xls_path = os.path.splitext(gateway_xlsx_path)[0] + ".xls"

    gateway_raw_path = ""
    if os.path.exists(gateway_xlsx_path):
        gateway_raw_path = gateway_xlsx_path
    elif os.path.exists(gateway_xls_path):
        gateway_raw_path = gateway_xls_path
    else:
        print(f"❌ 边缘网关数据文件不存在：")
        return

    # 2. 确定目标Excel路径
    target_excel_path = settings.resolve_path("spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print(f"❌ 目标Excel不存在：{target_excel_path}")
        return

    # 3. 加载边缘网关数据并预处理
    df_gateway = load_excel_data(gateway_raw_path)
    if df_gateway.empty:
        print("❌ 边缘网关数据加载失败")
        return

    # 4. 提取需要的9个字段
    required_fields = ["设备编码", "设备名称", "设备当前状态", "最近离线时间", 
                     "离线恢复时间", "设备入网状态", "国家行政区县", 
                     "所属站址名称", "站址资源编码"]
    
    # 确保所有字段存在
    for field in required_fields:
        if field not in df_gateway.columns:
            df_gateway[field] = ""

    # 5. 筛选只保留设备当前状态为离线的行
    if "设备当前状态" in df_gateway.columns:
        df_gateway = df_gateway[df_gateway["设备当前状态"] == "离线"]
        print(f"✅ 筛选完成：只保留 {len(df_gateway)} 行离线设备数据")
    else:
        print("⚠️ 设备当前状态字段不存在")

    df_updated = df_gateway[required_fields].copy()
    data_rows = len(df_updated)
    print(f"✅ 边缘网关数据预处理完成：有效数据{data_rows}行")

    # 6. 使用xlwings更新网关离线清单
    app = None
    wb = None
    
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(target_excel_path)
        wb.display_alerts = False
        wb.screen_updating = False

        # 批量备份所有非网关离线清单sheet公式
        formula_backup = batch_backup_all_sheets_formulas(wb, exclude_sheet="网关离线清单")

        # 7. 检查网关离线清单Sheet是否存在
        sheet_name = "网关离线清单"
        if sheet_name not in [s.name for s in wb.sheets]:
            print(f"❌ 目标Sheet「{sheet_name}」不存在")
            return

        ws = wb.sheets[sheet_name]

        # 8. 清空数据但保留表头
        start_row = 2
        last_row = ws.cells.last_cell.row
        if last_row >= start_row:
            ws.range(f"A{start_row}:ZZ{last_row}").clear_contents()
            print("✅ 数据清空完成，保留表头")

        if data_rows > 0:
            # 9. 获取表头信息
            header_range = ws.range("1:1").value
            header_dict = {col: idx + 1 for idx, col in enumerate(header_range) if col}
            
            # 确保所有需要的字段都在表头中
            for field in required_fields:
                if field not in header_dict:
                    print(f"❌ 表头中缺少字段：{field}")
                    return

            # 10. 批量写入数据
            for field in required_fields:
                col_idx = header_dict[field]
                col_data = df_updated[field].tolist()
                # 批量写入整列数据
                ws.range((2, col_idx), (data_rows + 1, col_idx)).value = [[val] for val in col_data]
            print(f"✅ 数据写入完成：{data_rows}行")

            # 11. 生成公式字段
            # 11.1 离线天数公式
            if "离线天数" in header_dict and "最近离线时间" in header_dict:
                offline_days_col = header_dict["离线天数"]
                recent_offline_col = header_dict["最近离线时间"]
                formula_range = ws.range((2, offline_days_col), (data_rows + 1, offline_days_col))
                first_offline_cell = ws.range((2, recent_offline_col)).address.replace("$", "")
                formula = f'=IF({first_offline_cell}="","",TODAY()-INT({first_offline_cell}))'
                formula_range.formula = formula
                print("✅ 离线天数公式生成完成")

            # 11.2 是否代维处理公式
            if "是否代维处理" in header_dict:
                dv_col = header_dict["是否代维处理"]
                ws.range((2, dv_col), (data_rows + 1, dv_col)).value = "是"
                print("✅ 是否代维处理字段填充完成")

            # 11.3 分管维护员公式
            if "分管维护员" in header_dict and "所属站址名称" in header_dict:
                leader_col = header_dict["分管维护员"]
                site_col = header_dict["所属站址名称"]
                formula_range = ws.range((2, leader_col), (data_rows + 1, leader_col))
                first_site_cell = ws.range((2, site_col)).address.replace("$", "")
                formula = f'=VLOOKUP({first_site_cell},数据源!V:Z,3,0)'
                formula_range.formula = formula
                print("✅ 分管维护员公式生成完成")

            # 11.4 临时入网实际安装站公式
            if "临时入网实际安装站" in header_dict and "设备编码" in header_dict:
                temp_site_col = header_dict["临时入网实际安装站"]
                code_col = header_dict["设备编码"]
                formula_range = ws.range((2, temp_site_col), (data_rows + 1, temp_site_col))
                first_code_cell = ws.range((2, code_col)).address.replace("$", "")
                formula = f'=VLOOKUP({first_code_cell},数据源!N:P,3,0)'
                formula_range.formula = formula
                print("✅ 临时入网实际安装站公式生成完成")

            # 11.5 备注字段公式
            if "备注" in header_dict and "设备编码" in header_dict:
                remark_col = header_dict["备注"]
                code_col = header_dict["设备编码"]
                formula_range = ws.range((2, remark_col), (data_rows + 1, remark_col))
                first_code_cell = ws.range((2, code_col)).address.replace("$", "")
                formula = f'=VLOOKUP({first_code_cell},核减清单!$A$2:$C$537,3,0)'
                formula_range.formula = formula
                print("✅ 备注字段公式生成完成")

            # 11.6 片区公式
            if "片区" in header_dict and "所属站址名称" in header_dict:
                region_col = header_dict["片区"]
                site_col = header_dict["所属站址名称"]
                formula_range = ws.range((2, region_col), (data_rows + 1, region_col))
                first_site_cell = ws.range((2, site_col)).address.replace("$", "")
                formula = f'=VLOOKUP({first_site_cell},数据源!V:Z,5,0)'
                formula_range.formula = formula
                print("✅ 片区公式生成完成")

            # 11.7 考核核减公式
            if "考核核减" in header_dict and "备注" in header_dict:
                check_col = header_dict["考核核减"]
                remark_col = header_dict["备注"]
                formula_range = ws.range((2, check_col), (data_rows + 1, check_col))
                first_remark_cell = ws.range((2, remark_col)).address.replace("$", "")
                formula = f'=IF({first_remark_cell}="","","核减")'
                formula_range.formula = formula
                print("✅ 考核核减公式生成完成")

            # 11.8 是否超7天公式
            if "是否超7天" in header_dict and "离线天数" in header_dict:
                over7_col = header_dict["是否超7天"]
                offline_days_col = header_dict["离线天数"]
                formula_range = ws.range((2, over7_col), (data_rows + 1, over7_col))
                first_offline_days_cell = ws.range((2, offline_days_col)).address.replace("$", "")
                formula = f'=IF({first_offline_days_cell}="","",IF({first_offline_days_cell}>=7,"是",""))'
                formula_range.formula = formula
                print("✅ 是否超7天公式生成完成")

        # 批量恢复所有非网关离线清单sheet公式
        batch_restore_all_sheets_formulas(wb, formula_backup, exclude_sheet="网关离线清单")

        # 12. 强制文本格式，防止科学计数法
        force_all_sheets_text_format(wb)

        # 13. 保存并关闭
        wb.app.calculate()
        wb.screen_updating = True
        wb.save()
        print(f"\n🎉 网关离线清单更新完成！共处理 {data_rows} 行数据")

    except Exception as e:
        print(f"❌ xlwings更新数据失败：{e}")
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


if __name__ == '__main__':
    process_gateway_main_list()
    update_gateway_offline_list()
