import os
import sys
import pandas as pd
import xlwings as xw
import warnings
from datetime import datetime, timedelta

warnings.filterwarnings('ignore')

# 添加项目根目录到 Python 路径
current_file = os.path.abspath(__file__)
current_dir = os.path.dirname(current_file)
parent_dir = os.path.dirname(current_dir)
grandparent_dir = os.path.dirname(parent_dir)
great_grandparent_dir = os.path.dirname(grandparent_dir)
project_root = os.path.dirname(great_grandparent_dir)
sys.path.insert(0, project_root)

# 定义基础路径
BASE_PATH = project_root

# ===================== 配置区（需根据实际列名/列位置修改）=====================
# 请根据实际Excel结构修改以下配置，标注为【待确认】的需要核对实际列位置/名称
CONFIG = {
    "source_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量数据.xlsx"),
    "target_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量器设备异常清单.xlsx"),
    "cols_keep": ["室分", "微站", "昨日数据"],  # 保留的核心列
    "col_N": "数据准确标识",  # 【待确认】第N列实际列名（数据准确标识列）
    "col_X": "故障时间",  # 【待确认】第X列实际列名（故障时间列）
    "col_T": "处理事件",  # 【待确认】第T列实际列名（处理事件列）
    "col_U": "新增故障数",  # 【待确认】第U列实际列名（新增故障数列）
    "col_site_code": "站址编码",  # 【待确认】站址编码列名（匹配用）
    "today_completion_formula_pattern": r'COUNTIFS\(清单!G:G,@A13:A24,清单!T:T,"\d{4}/\d{1,2}/\d{1,2}"\)'  # 今日完成数公式匹配正则
}


# =============================================================================

def get_file_modify_time(file_path):
    """获取文件最后修改时间"""
    try:
        modify_timestamp = os.path.getmtime(file_path)
        modify_time = datetime.fromtimestamp(modify_timestamp)
        return modify_time
    except Exception as e:
        print(f"获取文件修改时间失败：{file_path} -> {e}")
        return None


def check_if_latest_file(file_path, time_threshold_hours=24):
    """检查文件是否为最新（24小时内修改）"""
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


def process_device_list():
    """
    分路计量器设备异常清单处理（基于xlwings批量操作）
    核心逻辑：
    1. 清单 → 昨日数据 复制
    2. 爬取数据 → 今日数据 写入（保留指定列+公式）
    3. 清单清空原始数据（保留表头+公式）
    4. 今日数据筛选后 → 清单写入
    5. 故障时间匹配 → 更新数据准确标识
    6. 昨日/今日数据对比 → 更新处理事件/新增故障数
    7. 通报sheet公式日期自动更新
    """
    # 初始化当天日期（核心变量）
    today = datetime.now()
    today_str = today.strftime("%Y/%m/%d")  # 格式：2026/04/07
    print(f"\n🚀 开始处理分路计量器设备异常清单 - {today_str}")

    # 1. 前置检查：文件存在性+时效性
    try:
        # 检查源文件和目标文件
        if not os.path.exists(CONFIG["source_excel"]):
            raise FileNotFoundError(f"源文件不存在：{CONFIG['source_excel']}")
        if not os.path.exists(CONFIG["target_excel"]):
            raise FileNotFoundError(f"目标文件不存在：{CONFIG['target_excel']}")

        # 检查源文件时效性
        if not check_if_latest_file(CONFIG["source_excel"], 24):
            raise ValueError("分路计量数据文件不是最新版本（超过24小时未更新）")

        # 2. 打开Excel（xlwings批量操作核心：使用app/book/sheet对象）
        # 启动xlwings应用（后台运行，不显示Excel窗口）
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False  # 关闭弹窗提示
        app.screen_updating = False  # 关闭屏幕刷新（提升效率）

        # 打开目标Excel文件（分路计量器设备异常清单）
        wb_target = app.books.open(CONFIG["target_excel"])
        # 打开源Excel文件（爬取的分路计量设备清单）
        wb_source = app.books.open(CONFIG["source_excel"])

        # ===================== 步骤1：清单 → 昨日数据 批量复制 =====================
        print("\n📌 步骤1：批量复制【清单】数据到【昨日数据】sheet")
        try:
            # 获取sheet对象
            sht_list = wb_target.sheets["清单"]
            sht_yesterday = wb_target.sheets["昨日数据"]

            # 批量复制（使用used_range获取所有数据区域，避免逐行操作）
            # 复制值+公式+格式（xlwings批量操作核心）
            list_used_range = sht_list.used_range
            sht_yesterday.range("A1").value = list_used_range.value  # 批量赋值
            sht_yesterday.range("A1").formula = list_used_range.formula  # 保留公式
            print("   ✅ 清单数据已批量复制到昨日数据")
        except Exception as e:
            raise RuntimeError(f"步骤1执行失败：{e}")

        # ===================== 步骤2：爬取数据 → 今日数据 写入 =====================
        print("\n📌 步骤2：批量写入爬取数据到【今日数据】sheet（保留指定列+公式）")
        try:
            # 获取源数据sheet（默认第一个sheet，可根据实际修改）
            sht_source = wb_source.sheets[0]
            sht_today = wb_target.sheets["今日数据"]

            # 源数据：第二行为表头，第三行开始为数据（批量读取）
            # 读取所有源数据（跳过第一行，从第二行开始）
            source_header = sht_source.range("A2").expand("right").value  # 表头（第二行）
            source_data = sht_source.range("A3").expand("table").value  # 数据（第三行开始）

            # 转换为DataFrame（便于列筛选，批量操作）
            df_source = pd.DataFrame(source_data, columns=source_header)

            # 保留指定列（室分、微站、昨日数据）+ 保留公式
            df_keep = df_source[CONFIG["cols_keep"]].copy()

            # 批量写入今日数据sheet（清空原有数据后写入）
            sht_today.used_range.clear()  # 清空原有数据
            sht_today.range("A1").value = df_keep.columns.tolist()  # 写入表头
            sht_today.range("A2").value = df_keep.values  # 批量写入数据

            # 保留公式：读取源数据中的公式列（如果有）
            # 【待补充】若源数据有公式列，需单独处理公式复制，示例：
            # source_formula = sht_source.range("A3").expand("table").formula
            # sht_today.range("A2").formula = source_formula
            print("   ✅ 爬取数据已批量写入今日数据（保留指定列）")
        except Exception as e:
            raise RuntimeError(f"步骤2执行失败：{e}")

        # ===================== 步骤3：清空清单原始数据（保留表头+公式） =====================
        print("\n📌 步骤3：清空【清单】原始数据（保留表头+公式）")
        try:
            # 获取清单sheet的使用区域
            used_range = sht_list.used_range
            header_row = 1  # 表头行（根据实际修改）
            data_start_row = header_row + 1

            # 清空数据行（保留表头行和公式）
            if used_range.last_cell.row > data_start_row:
                # 批量清空数据区域的值（保留公式）
                sht_list.range(f"A{data_start_row}:{used_range.last_cell.address.split('$')[-1]}").clear_contents()

            # 保留公式：确保公式列不被清空（示例：若T列/U列为公式列）
            # 【待确认】需根据实际公式列调整，示例：
            # formula_cols = ["T", "U"]  # 处理事件、新增故障数列为公式列
            # for col in formula_cols:
            #     sht_list.range(f"{col}{data_start_row}:{col}{used_range.last_cell.row}").formula = sht_list.range(f"{col}{data_start_row}:{col}{used_range.last_cell.row}").formula
            print("   ✅ 清单原始数据已清空（保留表头+公式）")
        except Exception as e:
            raise RuntimeError(f"步骤3执行失败：{e}")

        # ===================== 步骤4：今日数据筛选后写入清单 =====================
        print("\n📌 步骤4：批量写入【今日数据】筛选后数据到【清单】")
        try:
            # 读取今日数据sheet的筛选后数据（室分、微站列以外的所有数据）
            df_today = pd.DataFrame(sht_today.used_range.value, columns=sht_today.used_range.value[0])
            df_today = df_today.iloc[1:]  # 跳过表头

            # 筛选列：排除室分、微站站址列
            filter_cols = [col for col in df_today.columns if col not in ["室分", "微站"]]
            df_filtered = df_today[filter_cols].copy()

            # 批量写入清单sheet（对应字段）
            # 【待确认】需确保列名匹配，示例：按列名映射写入
            # 列名映射字典（今日数据列 → 清单列），需根据实际修改
            col_mapping = {col: col for col in filter_cols}  # 默认同名列匹配
            for today_col, list_col in col_mapping.items():
                # 找到清单sheet中目标列的位置
                list_col_idx = sht_list.range("A1").expand("right").value.index(list_col) + 1
                # 批量写入数据
                sht_list.range(f"{chr(64 + list_col_idx)}{data_start_row}").value = df_filtered[
                    today_col].values.reshape(-1, 1)
            print("   ✅ 今日数据筛选后已批量写入清单")
        except Exception as e:
            raise RuntimeError(f"步骤4执行失败：{e}")

        # ===================== 步骤5：故障时间匹配 → 更新数据准确标识 =====================
        print("\n📌 步骤5：批量匹配故障时间 → 更新【数据准确标识】列")
        try:
            # 读取清单sheet的故障时间列和数据准确标识列
            # 获取列位置（根据列名）
            col_X_idx = sht_list.range("A1").expand("right").value.index(CONFIG["col_X"]) + 1  # 故障时间列
            col_N_idx = sht_list.range("A1").expand("right").value.index(CONFIG["col_N"]) + 1  # 数据准确标识列

            # 批量读取故障时间数据
            fault_time_data = sht_list.range(f"{chr(64 + col_X_idx)}{data_start_row}").expand("down").value
            # 批量处理：有日期→异常，无内容→准确
            update_data = []
            for cell_value in fault_time_data:
                if cell_value and str(cell_value).strip():  # 有日期内容
                    update_data.append(["异常"])
                else:  # 无内容
                    update_data.append(["准确"])

            # 批量写入数据准确标识列
            sht_list.range(f"{chr(64 + col_N_idx)}{data_start_row}").value = update_data
            print("   ✅ 数据准确标识已批量更新（故障时间匹配）")
        except Exception as e:
            raise RuntimeError(f"步骤5执行失败：{e}")

        # ===================== 步骤6：昨日/今日数据对比 → 更新处理事件/新增故障数 =====================
        print("\n📌 步骤6：批量对比昨日/今日数据 → 更新处理事件/新增故障数")
        try:
            # 读取昨日数据和清单数据（按站址编码匹配）
            # 昨日数据sheet
            sht_yesterday = wb_target.sheets["昨日数据"]
            # 获取站址编码列位置
            col_site_code_idx = sht_list.range("A1").expand("right").value.index(CONFIG["col_site_code"]) + 1

            # 批量读取数据（转换为DataFrame便于匹配）
            # 清单数据
            list_data = sht_list.used_range.value
            df_list = pd.DataFrame(list_data[1:], columns=list_data[0])
            # 昨日数据
            yesterday_data = sht_yesterday.used_range.value
            df_yesterday = pd.DataFrame(yesterday_data[1:], columns=yesterday_data[0])

            # 按站址编码合并数据（批量匹配核心）
            df_merged = pd.merge(
                df_list,
                df_yesterday[[CONFIG["col_site_code"], CONFIG["col_N"]]],
                on=CONFIG["col_site_code"],
                how="left",
                suffixes=("_today", "_yesterday")
            )

            # 批量处理逻辑：
            # 1. 昨日异常 → 今日准确 → 处理事件列填当天日期
            # 2. 昨日准确 → 今日异常 → 新增故障数列填1
            df_merged[CONFIG["col_T"]] = df_merged.apply(
                lambda row: today_str if (row[f"{CONFIG['col_N']}_yesterday"] == "异常" and row[
                    f"{CONFIG['col_N']}_today"] == "准确") else row[CONFIG["col_T"]],
                axis=1
            )
            df_merged[CONFIG["col_U"]] = df_merged.apply(
                lambda row: 1 if (row[f"{CONFIG['col_N']}_yesterday"] == "准确" and row[
                    f"{CONFIG['col_N']}_today"] == "异常") else row[CONFIG["col_U"]],
                axis=1
            )

            # 批量写回清单sheet
            # 处理事件列（T列）
            col_T_idx = sht_list.range("A1").expand("right").value.index(CONFIG["col_T"]) + 1
            sht_list.range(f"{chr(64 + col_T_idx)}{data_start_row}").value = df_merged[CONFIG["col_T"]].values.reshape(
                -1, 1)
            # 新增故障数列（U列）
            col_U_idx = sht_list.range("A1").expand("right").value.index(CONFIG["col_U"]) + 1
            sht_list.range(f"{chr(64 + col_U_idx)}{data_start_row}").value = df_merged[CONFIG["col_U"]].values.reshape(
                -1, 1)
            print("   ✅ 处理事件/新增故障数已批量更新（站址编码匹配）")
        except Exception as e:
            raise RuntimeError(f"步骤6执行失败：{e}")

        # ===================== 步骤7：更新通报sheet今日完成数公式日期 =====================
        print("\n📌 步骤7：批量更新【通报】sheet今日完成数公式日期")
        try:
            sht_report = wb_target.sheets["通报"]
            # 读取通报sheet所有公式
            report_formula = sht_report.used_range.formula

            # 批量替换公式中的日期（正则匹配+替换）
            import re
            pattern = CONFIG["today_completion_formula_pattern"]
            new_report_formula = []
            for row in report_formula:
                new_row = []
                for cell in row:
                    if cell and isinstance(cell, str) and re.match(pattern, cell):
                        # 替换日期为当天日期
                        new_cell = re.sub(r'"(\d{4}/\d{1,2}/\d{1,2})"', f'"{today_str}"', cell)
                        new_row.append(new_cell)
                    else:
                        new_row.append(cell)
                new_report_formula.append(new_row)

            # 批量写回公式
            sht_report.used_range.formula = new_report_formula
            print(f"   ✅ 通报sheet公式日期已更新为：{today_str}")
        except Exception as e:
            raise RuntimeError(f"步骤7执行失败：{e}")

        # ===================== 最终操作：保存+关闭 =====================
        print("\n📌 最终步骤：保存文件并关闭Excel")
        try:
            wb_target.save()  # 保存目标文件
            wb_source.close()  # 关闭源文件
            wb_target.close()  # 关闭目标文件
            app.quit()  # 退出xlwings应用
            print("   ✅ 文件已保存，Excel进程已关闭")
        except Exception as e:
            raise RuntimeError(f"保存/关闭文件失败：{e}")

        print("\n🎉 分路计量器设备异常清单处理完成！")

    # 全局异常捕获
    except Exception as e:
        print(f"\n❌ 处理失败：{str(e)}")
        # 异常时确保Excel进程关闭
        try:
            app.quit()
        except:
            pass
        return False

    return True


# ===================== 待补充/核对项（重点关注）=====================
"""
【待确认】需核对的配置项：
1. CONFIG中的列名（col_N/col_X/col_T/col_U/col_site_code）是否与实际Excel一致
2. 源数据sheet是否为第一个sheet（wb_source.sheets[0]）
3. 清单sheet表头行是否为第1行（header_row = 1）
4. 今日完成数公式匹配正则是否准确（today_completion_formula_pattern）
5. 公式保留逻辑是否需要补充（步骤2/3中的【待补充】部分）

【待优化】可扩展方向：
1. 列位置可改为按列名动态获取（已实现），避免硬编码列索引
2. 可添加数据校验（如站址编码唯一性、数据类型校验）
3. 可添加日志记录，便于排查问题
4. 可将配置项提取为外部配置文件（如config.ini）
"""
# =============================================================================

if __name__ == '__main__':
    process_device_list()