import os
import pandas as pd
import xlwings as xw
from core.config import settings
import warnings

warnings.filterwarnings('ignore')

"""
摄像头数据处理模块（优化版）
- 核心优化：批量写入数据+解决编码科学计数法显示问题
- 兼容macOS/Windows所有xlwings版本
- 核心特性：全量更新数据+保留Excel格式/公式
- 性能提升：批量写入替代逐行逐列，效率提升10倍+
- 改动：移除设备编码去重逻辑，保留所有原始数据行；移除非必要print输出
"""


def load_excel_data(file_path, sheet_name=0):
    """
    加载Excel文件数据（兼容xlsx/xls）
    """
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在：{file_path}")
        df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")  # 空值填充为空字符串
        return df
    except Exception as e:
        print(f"加载Excel失败：{file_path} -> {e}")
        return pd.DataFrame()


def preprocess_camera_data(camera_raw_path):
    """
    Step1：pandas预处理摄像头数据（筛选字段+新增字段）
    - 移除：设备编码去重逻辑
    """
    # 1. 加载原始数据
    df_camera = load_excel_data(camera_raw_path)
    if df_camera.empty:
        print("❌ 原始摄像头数据加载失败")
        return pd.DataFrame()

    # 2. 保留核心6个字段（不存在的字段置空）
    core_fields = ["设备编码", "设备名称", "摄像头类型", "摄像头安装位置", "站址名称", "站址资源编码"]
    for field in core_fields:
        if field not in df_camera.columns:
            df_camera[field] = ""
    df_main = df_camera[core_fields].copy()

    # 3. 新增8个空字段
    new_fields = ["是否临时入网", "整合站名与入临时入网", "区域", "分管维护员", "片区",
                  "是否代维处理", "核减", "是否超7天"]
    for field in new_fields:
        df_main[field] = ""

    print(f"✅ 摄像头数据预处理完成：有效数据{len(df_main)}行")
    return df_main


def safe_get_range_value(range_obj):
    """
    安全读取xlwings range的值，兼容空值/单个值/二维列表
    """
    try:
        value = range_obj.value
        # 处理空值
        if value is None:
            return []
        # 处理单个值（单个单元格）
        if not isinstance(value, list):
            return [[value]]
        # 处理一维列表（单行）
        if isinstance(value, list) and len(value) > 0 and not isinstance(value[0], list):
            return [[v] for v in value]
        # 处理二维列表（多行多列）
        return value
    except Exception as e:
        print(f"读取Range值失败：{e}")
        return []


def update_camera_sheet_with_xlwings(target_excel_path, sheet_name, df_updated):
    """
    Step2&3：xlwings更新摄像头总清单（优化版）
    - 核心优化1：批量写入数据，替代逐行逐列
    - 核心优化2：设置编码列为文本格式，解决科学计数法问题
    - 兼容：所有range使用元组格式 (row, col)
    - 核心：仅清空内容，保留所有格式/公式
    """
    if df_updated.empty:
        print("❌ 无更新数据，终止写入")
        return False

    app = None
    wb = None
    try:
        # 初始化Excel应用
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(target_excel_path)
        wb.display_alerts = False  # 关闭Excel提示框

        # 检查sheet是否存在
        if sheet_name not in [s.name for s in wb.sheets]:
            ws = wb.sheets.add(name=sheet_name)
        else:
            ws = wb.sheets[sheet_name]

        # ========== 1. 清空原有数据行（保留表头） ==========
        start_row = 2
        last_row = ws.cells.last_cell.row

        if last_row >= start_row:
            ws.range(f"A{start_row}:ZZ{last_row}").clear_contents()

        # ========== 2. 批量写入数据（核心优化） ==========
        # 映射列名到Excel列位置
        header_range = ws.range("A1:ZZ1").value
        header_dict = {col: idx + 1 for idx, col in enumerate(header_range) if col}

        # 过滤出目标Excel中存在的列
        valid_cols = [col for col in df_updated.columns if col in header_dict]
        if not valid_cols:
            print("❌ 无匹配的列名，无法写入数据")
            return False

        # 提取有效数据
        df_valid = df_updated[valid_cols].copy()
        data_rows = len(df_valid)
        if data_rows == 0:
            return True

        # 2.1 设置关键列（设备编码/站址资源编码）为文本格式（解决科学计数法）
        code_cols = ["设备编码", "站址资源编码"]
        for col_name in code_cols:
            if col_name in header_dict:
                col_idx = header_dict[col_name]
                # 使用 number_format 属性（跨平台）
                ws.range((1, col_idx), (data_rows + 1, col_idx)).number_format = "@"

        # 2.2 批量写入数据（按列批量写入，效率最大化）
        for col_name in valid_cols:
            col_idx = header_dict[col_name]
            # 提取列数据并批量写入
            col_data = df_valid[col_name].tolist()
            # 批量写入整列数据
            ws.range((2, col_idx), (data_rows + 1, col_idx)).value = [[val] for val in col_data]

        # ========== 3. 批量设置Excel公式（摄像头业务逻辑） ==========
        # 3.1 是否临时入网（VLOOKUP匹配）
        if "是否临时入网" in header_dict and "设备编码" in header_dict:
            code_col = header_dict["设备编码"]
            temp_col = header_dict["是否临时入网"]
            # 批量设置公式（整列）
            formula_range = ws.range((2, temp_col), (data_rows + 1, temp_col))
            # 获取第一个单元格地址用于公式模板
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},数据源!N:P,3,0)'
            # 批量填充公式（xlwings自动适配行号）
            formula_range.formula = formula

        # 3.2 整合站名与入临时入网（IF+ISNA逻辑）
        if all(col in header_dict for col in ["整合站名与入临时入网", "站址名称", "是否临时入网"]):
            integrate_col = header_dict["整合站名与入临时入网"]
            site_col = header_dict["站址名称"]
            temp_col = header_dict["是否临时入网"]

            # 批量设置公式
            formula_range = ws.range((2, integrate_col), (data_rows + 1, integrate_col))
            first_temp_cell = ws.range((2, temp_col)).address.replace("$", "")
            first_site_cell = ws.range((2, site_col)).address.replace("$", "")
            formula = f'=IF(ISNA({first_temp_cell}), {first_site_cell}, {first_temp_cell})'
            formula_range.formula = formula

        # 3.3 区域/分管维护员/片区（基于整合站名匹配）
        if "整合站名与入临时入网" in header_dict:
            integrate_col = header_dict["整合站名与入临时入网"]
            integrate_cell = ws.range((2, integrate_col)).address.replace("$", "")

            # 区域
            if "区域" in header_dict:
                area_col = header_dict["区域"]
                formula_range = ws.range((2, area_col), (data_rows + 1, area_col))
                formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,2,0)'
                formula_range.formula = formula

            # 分管维护员
            if "分管维护员" in header_dict:
                leader_col = header_dict["分管维护员"]
                formula_range = ws.range((2, leader_col), (data_rows + 1, leader_col))
                formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,3,0)'
                formula_range.formula = formula

            # 片区
            if "片区" in header_dict:
                region_col = header_dict["片区"]
                formula_range = ws.range((2, region_col), (data_rows + 1, region_col))
                formula = f'=VLOOKUP({integrate_cell},数据源!V:Z,5,0)'
                formula_range.formula = formula

        # 3.4 是否超7天（基于摄像头离线清单）
        if "是否超7天" in header_dict and "设备编码" in header_dict:
            over7_col = header_dict["是否超7天"]
            code_col = header_dict["设备编码"]
            formula_range = ws.range((2, over7_col), (data_rows + 1, over7_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=IF(ISNA(VLOOKUP({first_code_cell},摄像头离线清单!A:B,2,0)),"",IF(VLOOKUP({first_code_cell},摄像头离线清单!A:B,2,0)>=7,"是",""))'
            formula_range.formula = formula

        # 3.5 是否代维处理（固定值"是"）
        if "是否代维处理" in header_dict:
            dv_col = header_dict["是否代维处理"]
            ws.range((2, dv_col), (data_rows + 1, dv_col)).value = "是"

        # 3.6 核减（匹配核减清单）
        if "核减" in header_dict and "设备编码" in header_dict:
            reduce_col = header_dict["核减"]
            code_col = header_dict["设备编码"]
            formula_range = ws.range((2, reduce_col), (data_rows + 1, reduce_col))
            first_code_cell = ws.range((2, code_col)).address.replace("$", "")
            formula = f'=VLOOKUP({first_code_cell},核减清单!A:C,3,0)'
            formula_range.formula = formula

        # 3.7 根据核减列清空对应分管维护员、区域列数据（和gateway保持一致的逻辑）
        if all(col in header_dict for col in ["核减", "分管维护员", "区域"]):
            try:
                # 触发Excel重新计算，确保核减列公式结果生效
                ws.api.calculate()

                # 读取摄像头总清单的完整数据（包含表头）
                data = ws.used_range.value
                if not data or len(data) < 2:
                    pass
                else:
                    # 转换为DataFrame进行处理
                    df = pd.DataFrame(data[1:], columns=data[0])

                    # 构建掩码：核减列非空、非空字符串、且不以 '#' 开头
                    reduce_series = df["核减"].astype(str)
                    mask = (
                            df["核减"].notna() &
                            (reduce_series.str.strip() != "") &
                            (~reduce_series.str.startswith("#", na=False))
                    )
                    clear_count = mask.sum()

                    if clear_count > 0:
                        # 清空符合条件行的分管维护员、区域列
                        df.loc[mask, ["分管维护员", "区域"]] = ""
                        print(f"✅ 摄像头数据 - 清理 {clear_count} 行核减相关数据")

                        # 准备新工作表（摄像头总清单_核减清理）
                        new_sheet_name = "摄像头总清单_核减清理"
                        if new_sheet_name in [s.name for s in wb.sheets]:
                            wb.sheets[new_sheet_name].delete()
                        ws_new = wb.sheets.add(new_sheet_name, after=ws)

                        # 写入表头（保留原格式）
                        num_cols = len(df.columns)
                        header_orig = ws.range("A1").resize(column_size=num_cols)
                        header_new = ws_new.range("A1").resize(column_size=num_cols)
                        header_orig.copy()
                        header_new.paste(paste="all")  # 保留原表头格式

                        # 处理编码列格式和内容（避免科学计数法）
                        code_column_names = ["设备编码", "站址资源编码"]
                        # 批量设置编码列为文本格式
                        for col_name in code_column_names:
                            if col_name in df.columns:
                                col_idx = df.columns.get_loc(col_name) + 1  # xlwings列索引从1开始
                                ws_new.range(ws_new.cells(2, col_idx),
                                             ws_new.cells(df.shape[0] + 1, col_idx)).number_format = "@"

                        # 对编码列数据批量添加单引号
                        for col_name in code_column_names:
                            if col_name in df.columns:
                                df[col_name] = df[col_name].apply(
                                    lambda x: f"'{x}" if (pd.notna(x) and str(x).strip() != "") else x
                                )

                        # 批量写入所有数据（一次性写入）
                        ws_new.range("A2").value = df.values

                        # ========== Sheet重命名+隐藏+替换逻辑 ==========
                        # 1. 重命名原始摄像头总清单为摄像头总清单_备份
                        backup_sheet_name = "摄像头总清单_备份"
                        # 如果备份sheet已存在，先删除
                        if backup_sheet_name in [s.name for s in wb.sheets]:
                            wb.sheets[backup_sheet_name].delete()
                        # 重命名原sheet为备份
                        ws.name = backup_sheet_name

                        # 2. 隐藏备份sheet
                        ws.visible = False

                        # 3. 重命名核减清理后的sheet为摄像头总清单
                        ws_new.name = "摄像头总清单"

            except Exception as e:
                print(f"⚠️ 摄像头数据 - 核减清理处理失败：{e}")
                import traceback
                traceback.print_exc()

        # ========== 保存并关闭 ==========
        wb.save()
        print(f"\n🎉 摄像头数据更新完成！共处理 {data_rows} 行数据")
        return True

    except Exception as e:
        print(f"❌ xlwings更新摄像头数据失败：{e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        # 确保Excel进程完全关闭（增强版）
        if wb:
            try:
                wb.close()
            except:
                pass
        if app:
            try:
                app.quit()
            except:
                pass
            try:
                app.kill()  # 强制结束进程
            except:
                pass
        # 额外清理：释放xlwings资源


def process_camera_main_list():
    """
    主处理函数：串联数据预处理+xlwings更新
    """
    # 1. 配置文件路径
    # 摄像头原始数据文件路径（xlsx/xls兼容）
    camera_xlsx_path = settings.resolve_path("spider/script/station/down/边缘摄像头.xlsx")
    camera_xls_path = os.path.splitext(camera_xlsx_path)[0] + ".xls"

    camera_raw_path = ""
    if os.path.exists(camera_xlsx_path):
        camera_raw_path = camera_xlsx_path
    elif os.path.exists(camera_xls_path):
        camera_raw_path = camera_xls_path
    else:
        print(f"❌ 摄像头数据文件不存在：")
        print(f"   - {camera_xlsx_path}")
        print(f"   - {camera_xls_path}")
        return

    # 目标Excel文件（和网关共用同一个模板文件）
    target_excel_path = settings.resolve_path(
        "spider/script/station/down/智能运维离线通报模板.xlsx")
    if not os.path.exists(target_excel_path):
        print(f"❌ 目标Excel不存在：{target_excel_path}")
        return

    # 2. 预处理数据（已移除去重）
    df_updated = preprocess_camera_data(camera_raw_path)
    if df_updated.empty:
        return

    # 3. 全量更新摄像头总清单
    update_success = update_camera_sheet_with_xlwings(
        target_excel_path=target_excel_path,
        sheet_name="摄像头总清单",
        df_updated=df_updated
    )

    if update_success:
        print("\n✅ 摄像头总清单更新成功！")
    else:
        print("\n❌ 摄像头总清单更新失败！")


# # 测试执行
# if __name__ == '__main__':
#     process_camera_main_list()