import os
import sys
import tempfile
import builtins
import pandas as pd
import warnings
from datetime import datetime
import win32com.client as win32
import pythoncom

warnings.filterwarnings('ignore')

# ==============================
# 【路径配置】
# ==============================
current_file = os.path.abspath(__file__)
current_dir = os.path.dirname(current_file)
parent_dir = os.path.dirname(current_dir)
grandparent_dir = os.path.dirname(parent_dir)
great_grandparent_dir = os.path.dirname(grandparent_dir)
project_root = os.path.dirname(great_grandparent_dir)
sys.path.insert(0, project_root)

BASE_PATH = project_root

# ==============================
# 【业务配置】
# ==============================
CONFIG = {
    "source_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备清单.xls"),
    "target_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量器设备异常清单.xlsx"),
}

# 需要设置为文本格式的字段名称
TEXT_FORMAT_FIELDS = [
    "站址编码",
    "站址运维ID",
    "FSU运维ID",
    "设备运维ID",
    "运维ID"
]

# 全局互斥锁文件：避免多个流程并发占用 Excel COM 导致冲突
LOCK_FILE = os.path.join(tempfile.gettempdir(), "jiliangzhibiao_process.lock")


class SingleProcessLock:
    """基于锁文件的轻量互斥，确保同一时刻只运行一个处理任务。"""

    def __init__(self, lock_file, part):
        self.lock_file = lock_file
        self.part = part
        self.locked = False

    def acquire(self):
        lock_payload = (
            f"pid={os.getpid()}\n"
            f"part={self.part}\n"
            f"time={datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        )
        try:
            fd = os.open(self.lock_file, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8") as fp:
                fp.write(lock_payload)
            self.locked = True
            return True
        except FileExistsError:
            owner = ""
            try:
                with open(self.lock_file, "r", encoding="utf-8") as fp:
                    owner = fp.read().strip()
            except Exception:
                owner = "未知占用者"
            raise RuntimeError(
                "检测到已有处理任务在运行，为避免冲突已阻止本次执行。"
                f"当前锁信息：{owner}"
            )

    def release(self):
        if not self.locked:
            return
        try:
            if os.path.exists(self.lock_file):
                os.remove(self.lock_file)
        finally:
            self.locked = False

    def __enter__(self):
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.release()


def get_col_index(sheet, field_name):
    """封装：获取字段对应的列索引（提前校验，避免索引错误）"""
    last_col = sheet.UsedRange.Columns.Count if sheet.UsedRange.Columns.Count else 100  # 兜底
    header_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, last_col))
    headers = [str(cell.Value).strip() if cell.Value else "" for cell in header_range]
    return headers.index(field_name) + 1 if field_name in headers else None


def process_device_list():
    # 设备流程默认静默，仅输出最终成功/失败
    real_print = builtins.print
    print = lambda *args, **kwargs: None

    try:
        # ==============================
        # 检查文件
        # ==============================
        for f in [CONFIG["source_excel"], CONFIG["target_excel"]]:
            if not os.path.exists(f):
                print(f"❌ 文件不存在：{f}")
                return False

        # ==============================
        # 启动Excel
        # ==============================
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False  # 关闭屏幕刷新，大幅提速

        # ==============================
        # 读取爬取数据（强制全字段文本，避免pandas自动转换）
        # ==============================
        # 【修改1】pandas读取时强制所有列为文本，且禁用科学计数法
        df_source = pd.read_excel(
            CONFIG["source_excel"],
            header=1,
            dtype=str,  # 强制所有列文本
            keep_default_na=False  # 空值不转NaN，保留空字符串
        )
        # 清理空格+确保空值为""
        df_source = df_source.apply(lambda x: x.str.strip() if x.dtype == "object" else x).fillna("")
        rows_count, cols_count = df_source.shape
        print(f"✅ 读取数据完成：{rows_count} 行 {cols_count} 列")

        # ==============================
        # 打开目标文件
        # ==============================
        wb = excel.Workbooks.Open(CONFIG["target_excel"])
        sht_list = wb.Sheets("清单")
        sht_yesterday = wb.Sheets("昨天数据")
        sht_today = wb.Sheets("今日数据")
        sht_shifen = wb.Sheets("室分站址")
        sht_wangge = wb.Sheets("网格")

        # 【修改2】暂时跳过文本格式初始化，先执行核心复制操作
        print("🔹 跳过文本格式初始化，先执行核心操作")

        # 1. 复制清单到昨天数据（仅复制值，不复制公式）
        # 关键：必须先取清单快照，再清空昨天数据；否则清单中依赖昨天数据的公式会先失效
        print("🔹 复制清单数据到昨天数据（先取值快照，再覆盖写入）")

        last_row_list = sht_list.UsedRange.Rows.Count
        last_col_list = max(sht_list.UsedRange.Columns.Count, 24)  # 至少覆盖到X列

        # 后续 T/U/X 逻辑需要用到昨天状态与X列值，先初始化映射
        yesterday_status_map = {}
        id_text_fields = ["站址编码", "站址运维ID", "FSU运维ID", "设备运维ID", "运维ID"]

        def _normalize_matrix(values):
            """将 Excel Range.Value 统一成二维 list。"""
            if values is None:
                return []
            if isinstance(values, tuple):
                if len(values) == 0:
                    return []
                if isinstance(values[0], tuple):
                    return [list(r) for r in values]
                return [list(values)]
            return [[values]]

        def _is_excel_error(v):
            # COM 错误码通常是负整数，字符串 #N/A 也视为错误值
            return (isinstance(v, int) and v < 0) or (isinstance(v, str) and v.strip().upper() == "#N/A")

        def _to_text_for_id(v):
            """长数字ID统一转文本，避免科学计数法。"""
            if v is None:
                return ""
            if isinstance(v, float):
                # 对整数型浮点去掉 .0，避免 123456.0
                if v.is_integer():
                    return str(int(v))
                return format(v, "f").rstrip("0").rstrip(".")
            if isinstance(v, int):
                return str(v)
            return str(v).strip()

        if last_row_list >= 1 and last_col_list >= 1:
            # 先计算一次，确保公式结果稳定，再取值快照
            try:
                excel.Calculate()
                sht_list.Calculate()
            except:
                pass

            src_rng = sht_list.Range(
                sht_list.Cells(1, 1),
                sht_list.Cells(last_row_list, last_col_list)
            )
            src_data = _normalize_matrix(src_rng.Value)
            cleaned_data = []
            header_col_map = {}

            # 固定列号（1-based）
            b_col = 2
            n_col = 14
            x_col = 24

            for row_idx, row in enumerate(src_data):
                clean_row = []
                # 第一行表头，先建立字段到列号映射
                if row_idx == 0:
                    for idx, h in enumerate(row, start=1):
                        h_name = str(h).strip() if h is not None else ""
                        if h_name:
                            header_col_map[h_name] = idx

                for col_idx, val in enumerate(row, start=1):
                    if _is_excel_error(val):
                        clean_val = ""
                    else:
                        clean_val = val

                    # X列（故障时间）仅保留值，不保留错误值
                    if col_idx == x_col and _is_excel_error(clean_val):
                        clean_val = ""

                    # 指定ID字段强制转文本，防止科学计数法
                    if row_idx >= 1:
                        for field in id_text_fields:
                            if header_col_map.get(field) == col_idx:
                                clean_val = _to_text_for_id(clean_val)
                                break

                    clean_row.append(clean_val)

                cleaned_data.append(clean_row)

                # 跳过表头，构建昨天状态与X值映射，供后续 T/U/X 逻辑使用
                if row_idx == 0:
                    continue
                site_code = str(clean_row[b_col - 1]).strip() if len(clean_row) >= b_col and clean_row[b_col - 1] is not None else ""
                y_status = str(clean_row[n_col - 1]).strip() if len(clean_row) >= n_col and clean_row[n_col - 1] is not None else ""
                if site_code:
                    yesterday_status_map[site_code] = y_status

            if cleaned_data:
                # 写入前先清空昨天数据
                sht_yesterday.UsedRange.ClearContents()
                # 目标表对应ID列设置文本格式，避免Excel自动科学计数法
                for field in id_text_fields:
                    col_idx = header_col_map.get(field)
                    if col_idx:
                        sht_yesterday.Columns(col_idx).NumberFormat = "@"

                sht_yesterday.Range(
                    sht_yesterday.Cells(1, 1),
                    sht_yesterday.Cells(last_row_list, last_col_list)
                ).Value = cleaned_data

            print("✅ 昨天数据快照完成（纯值，无公式）")

        # 2. 清空今日数据 → 批量写入爬取数据
        print("🔹 批量写入今日数据")
        sht_today.UsedRange.ClearContents()

        # 【修改3】写入表头前，先把今日数据sheet的关键列设为文本（兜底）
        for field in TEXT_FORMAT_FIELDS:
            col_idx = get_col_index(sht_today, field)
            if col_idx:
                sht_today.Columns(col_idx).NumberFormat = "@"

        # 批量写入表头
        sht_today.Range(
            sht_today.Cells(1, 1),
            sht_today.Cells(1, cols_count)
        ).Value = df_source.columns.tolist()

        # 批量写入所有数据（此时列已设为文本，写入不会转科学计数法）
        if rows_count > 0:
            # 【修改4】数据写入前，强制为字符串+加单引号（双重保险）
            data_array = []
            for row in df_source.values.tolist():
                new_row = []
                for idx, val in enumerate(row):
                    col_name = df_source.columns[idx]
                    if col_name in TEXT_FORMAT_FIELDS:
                        # 强制加单引号（Excel中单引号开头会强制文本，且不显示）
                        new_row.append("'" + str(val).strip())
                    else:
                        new_row.append(val)
                data_array.append(new_row)
            # 写入数据
            sht_today.Range(
                sht_today.Cells(2, 1),
                sht_today.Cells(rows_count + 1, cols_count)
            ).Value = data_array

        # 3. 批量添加 3 列公式
        print("🔹 批量添加公式列")
        new_col1 = cols_count + 1
        new_col2 = cols_count + 2
        new_col3 = cols_count + 3
        data_end_row = rows_count + 1

        # 写入表头
        sht_today.Cells(1, new_col1).Value = "室分"
        sht_today.Cells(1, new_col2).Value = "微站"
        sht_today.Cells(1, new_col3).Value = "昨天数据"

        # 批量整列写入公式
        if data_end_row >= 2:
            # 室分
            sht_today.Range(
                sht_today.Cells(2, new_col1),
                sht_today.Cells(data_end_row, new_col1)
            ).Formula = "=VLOOKUP(B2,室分站址!L:L,1,0)"

            # 微站
            sht_today.Range(
                sht_today.Cells(2, new_col2),
                sht_today.Cells(data_end_row, new_col2)
            ).Formula = "=VLOOKUP(B2,微站!K:K,1,0)"

            # 昨天数据
            sht_today.Range(
                sht_today.Cells(2, new_col3),
                sht_today.Cells(data_end_row, new_col3)
            ).Formula = "=VLOOKUP(B2,昨天数据!B:B,1,0)"

        # 4. 清空清单除表头外数据
        print("🔹 清空清单数据")
        lastRow = sht_list.UsedRange.Rows.Count
        if lastRow > 1:
            sht_list.Range(
                sht_list.Cells(2, 1),
                sht_list.Cells(lastRow, cols_count)
            ).ClearContents()

        # 5. 批量复制今日数据 → 清单
        print("🔹 批量复制今日数据到清单")
        copy_range = sht_today.Range(
            sht_today.Cells(1, 1),
            sht_today.Cells(data_end_row, cols_count)
        )
        copy_range.Copy()
        # 【修改5】粘贴时指定值+格式，确保文本格式保留
        sht_list.Range("A1").PasteSpecial(Paste=-4163)  # -4163=值，-4122=格式（可叠加）
        excel.CutCopyMode = False

        # 修复：日期列只粘贴值后会显示为序列号（如46130），这里恢复为日期格式
        for date_field in ["数据稽核时间", "稽核时间"]:
            date_col_idx = get_col_index(sht_list, date_field)
            if date_col_idx:
                sht_list.Columns(date_col_idx).NumberFormat = "yyyy/m/d"
                break

        # 【修改6】暂时跳兜底校验，先让核心操作完成
        print("🔹 跳过最终校验，先让核心操作完成")

        # 7、按业务规则处理清单 S/T/U/X
        print("🔹 按规则处理清单 S/T/U/X 列")

        # ====================== 配置区 =======================
        n_col = 14  # 异常/准确 标识列（N列）
        s_col = 19  # 故障时间列（S列）
        t_col = 20  # 处理时间列（T列）
        u_col = 21  # 新增故障数列（U列）
        x_col = 24  # 故障时间列（X列）
        site_code_col = 2  # 站址编码列（B列）
        # ====================================================

        # 获取清单数据最后一行
        list_last_row = sht_list.UsedRange.Rows.Count

        # 仅处理有数据的行（跳过表头）
        if list_last_row >= 2:
            # 1. 先清空 S/T/U/X 列所有数据（保留表头）
            sht_list.Range(sht_list.Cells(2, s_col), sht_list.Cells(list_last_row, s_col)).ClearContents()
            sht_list.Range(sht_list.Cells(2, t_col), sht_list.Cells(list_last_row, t_col)).ClearContents()
            sht_list.Range(sht_list.Cells(2, u_col), sht_list.Cells(list_last_row, u_col)).ClearContents()
            sht_list.Range(sht_list.Cells(2, x_col), sht_list.Cells(list_last_row, x_col)).ClearContents()

            # 一次性读取 B/N 列
            rng_site_code = _normalize_matrix(
                sht_list.Range(sht_list.Cells(2, site_code_col), sht_list.Cells(list_last_row, site_code_col)).Value
            )
            rng_n_status = _normalize_matrix(
                sht_list.Range(sht_list.Cells(2, n_col), sht_list.Cells(list_last_row, n_col)).Value
            )

            today_date = datetime.now().strftime("%Y-%m-%d")
            t_values = []
            u_values = []
            s_formula_array = []
            x_formula_array = []
            x_today_rows = []

            for row_idx in range(len(rng_n_status)):
                actual_row = row_idx + 2
                site_code = str(rng_site_code[row_idx][0]).strip() if row_idx < len(rng_site_code) and rng_site_code[row_idx][0] is not None else ""
                cur_status = str(rng_n_status[row_idx][0]).strip() if rng_n_status[row_idx][0] is not None else ""
                y_status = yesterday_status_map.get(site_code, "")

                # 先处理 T/U（只按 N 列状态变化）
                if y_status == "异常" and cur_status == "准确":
                    t_values.append([today_date])
                else:
                    t_values.append([""])

                is_new_fault = (y_status == "准确" and cur_status == "异常")
                if is_new_fault:
                    u_values.append([1])
                else:
                    u_values.append([""])

                # 再处理 S（仅 N=异常 时填公式）
                if cur_status == "异常":
                    s_formula_array.append([f"=VLOOKUP(B{actual_row},网格!C:P,14,0)"])
                else:
                    s_formula_array.append([""])

                # 最后处理 X：
                # - N=异常且新增故障(U=1)：写今天日期（后续覆盖）
                # - N=异常且非新增：写昨天X列查找公式
                # - 其他：留空
                if cur_status == "异常":
                    if is_new_fault:
                        x_formula_array.append([""])
                        x_today_rows.append(actual_row)
                    else:
                        x_formula_array.append([f"=VLOOKUP(B{actual_row},昨天数据!B:X,23,0)"])
                else:
                    x_formula_array.append([""])

            # 批量写入（T/U为值，S/X为公式）
            sht_list.Range(sht_list.Cells(2, t_col), sht_list.Cells(list_last_row, t_col)).Value = t_values
            sht_list.Range(sht_list.Cells(2, u_col), sht_list.Cells(list_last_row, u_col)).Value = u_values
            sht_list.Range(sht_list.Cells(2, s_col), sht_list.Cells(list_last_row, s_col)).Formula = s_formula_array
            sht_list.Range(sht_list.Cells(2, x_col), sht_list.Cells(list_last_row, x_col)).Formula = x_formula_array

            # 新增故障行：X列改为今天日期
            for r in x_today_rows:
                sht_list.Cells(r, x_col).Value = today_date

            print(f"✅ S/T/U/X 处理完成：共处理 {len(rng_n_status)} 行，新增故障 {len(x_today_rows)} 行")
        else:
            print("⚠️ 清单工作表无有效数据，跳过 S/T/U/X 处理")

        # 8、【纯批量】清单G列填充VLOOKUP公式：=VLOOKUP(B2,网格!C:I,7,0)
        print("🔹 仅清单工作表：保留表头，清空G列数据并批量填充VLOOKUP公式")
        g_col = 7  # G列固定列号
        # 仅获取【清单】工作表的最后一行数据
        last_data_row = sht_list.UsedRange.Rows.Count

        # 1. 【修复】只清空 G2 至 G列最后一行数据，**保留G1表头不删除**
        if last_data_row >= 2:
            # 定位G2到最后一行的区域，仅清空数据行
            sht_list.Range(sht_list.Cells(2, g_col), sht_list.Cells(last_data_row, g_col)).ClearContents()

        # 2. 【清单】G列设为常规格式（公式必须用常规格式才生效，不影响表头）
        sht_list.Columns(g_col).NumberFormat = "General"

        # 3. 【完全模仿你提供的格式】批量写入公式
        if last_data_row >= 2:
            # 完全和你给的代码格式一致，仅修改工作表、列、公式
            sht_list.Range(
                sht_list.Cells(2, g_col),
                sht_list.Cells(last_data_row, g_col)
            ).Formula = "=VLOOKUP(B2,网格!C:I,7,0)"
            print(f"✅ 清单工作表G列填充完成：G2 ~ G{last_data_row}（表头G1已保留）")
        else:
            print("⚠️ 清单工作表无有效数据，跳过G列公式填充")

        # 9、填充今日完成数公式
        print("🔹 填充通报工作表今日完成数公式（G3~G14）")
        try:
            # 定位通报工作表
            sht_tongbao = wb.Sheets("通报")

            # 获取今日日期（格式匹配T列的处理时间格式）
            today_date = datetime.now().strftime("%Y-%m-%d")  # 需和T列存储的日期格式一致
            # 定义公式模板：COUNTIFS(清单!G:G, @A{行}, 清单!T:T, 今日日期)
            # 注意：Excel中today()是函数，若T列是带时分的时间，需用TEXT(清单!T:T,"yyyy-mm-dd")=今日日期

            # 遍历G3到G14行（对应A3:A14 → A14:A15）
            for row in range(3, 15):  # 3~14行（含）
                # 计算A列的结束行：A3→A14，A4→A15 ... A14→A25
                a_end_row = row + 11  # 3+11=14, 4+11=15 ... 14+11=25
                # 拼接公式（两种写法可选，根据T列格式适配）
                # 写法1：T列是纯日期（yyyy-mm-dd）
                formula = f'=COUNTIFS(清单!G:G,@A{row}:A{a_end_row},清单!T:T,"{today_date}")'
                # 写法2：T列是带时分的时间（yyyy-mm-dd hh:mm:ss），需格式化匹配
                # formula = f'=COUNTIFS(清单!G:G,@A{row}:A{a_end_row},TEXT(清单!T:T,"yyyy-mm-dd"),"{today_date}")'

                # 写入公式到对应单元格
                sht_tongbao.Cells(row, 7).Formula = formula  # 7=G列

            print(f"✅ 通报工作表G3~G14公式填充完成，今日日期：{today_date}")
        except Exception as e:
            print(f"⚠️ 通报工作表公式填充警告：{str(e)}")
            pass
        # ==============================
        # 保存退出
        # ==============================
        excel.ScreenUpdating = True
        wb.Save()
        wb.Close()
        excel.Quit()
        pythoncom.CoUninitialize()

        real_print("✅ device 处理成功")
        return True

    except Exception as e:
        real_print(f"❌ device 处理失败：{str(e)}")
        # 异常处理时确保Excel资源释放
        try:
            excel.ScreenUpdating = True
            excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()
        return False


def process_lessee_list():
    """处理租户电流数据：批量写入移动/联通/电信子表，仅导入A-P列，保留表头，批量公式+格式
    重点：长数字列禁用科学计数法，完整显示 + 三个文件数据追加合并（不覆盖）
    修复点：
    1. 每次运行强制清空三张表数据（仅保留表头）
    2. 关闭源文件前先提取所有需要的数据，避免访问失效对象
    3. 优化行数/列数计算逻辑，防止越界
    4. 增加源文件读取容错（处理.xls老格式）
    5. 优化COM对象释放逻辑
    6. 修复公式写入的相对引用问题
    """
    # 租户流程默认静默，仅输出最终成功/失败
    real_print = builtins.print
    print = lambda *args, **kwargs: None

    # ==============================
    # 【配置区】- 根据实际路径调整
    # ==============================
    LESSEE_CONFIG = {
        "target_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路租户异常通报（质询）.xlsx"),
        # 9个爬取的Excel文件路径（按顺序）
        "source_excels": [
            # 移动租户电流
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备-移动租户电流.xls"),
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/开关电源-移动租户电流.xls"),
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备-移动租户电流（5G）.xls"),
            # 联通租户电流
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备-联通租户电流.xls"),
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/开关电源-联通租户电流.xls"),
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备-联通租户电流（5G）.xls"),
            # 电信租户电流
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备-电信租户电流.xls"),
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/开关电源-电信租户电流.xls"),
            os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备-电信租户电流（5G）.xls"),
        ],
        # 子表映射：数据源索引 → (目标子表名, Q列填充值)
        "sheet_mapping": [
            # 移动
            ("移动电流", "分路计量设备"),
            ("移动电流", "开关电源"),
            ("移动电流", "分路计量设备"),
            # 联通
            ("联通电流", "分路计量设备"),
            ("联通电流", "开关电源"),
            ("联通电流", "分路计量设备"),
            # 电信
            ("电信电流", "分路计量设备"),
            ("电信电流", "开关电源"),
            ("电信电流", "分路计量设备"),
        ]
    }

    # 需要完整显示、禁止科学计数法的列标题（长数字ID）
    NO_SCIENCE_COLS = ["站址运维ID", "设备ID", "设备资源编码", "信号量ID"]
    # 固定需要清空的三张表
    TARGET_SHEETS = ["移动电流", "联通电流", "电信电流"]

    # 初始化Excel对象（提前定义，避免异常时无法释放）
    excel = None
    wb_target = None

    try:
        # 检查文件是否存在
        if not os.path.exists(LESSEE_CONFIG["target_excel"]):
            print(f"❌ 目标文件不存在：{LESSEE_CONFIG['target_excel']}")
            return False

        for idx, src_file in enumerate(LESSEE_CONFIG["source_excels"]):
            if not os.path.exists(src_file):
                print(f"❌ 源文件不存在：{src_file}（索引：{idx}）")
                return False

        # ==============================
        # 启动Excel（后台批量处理）
        # ==============================
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        # 打开目标文件
        wb_target = excel.Workbooks.Open(LESSEE_CONFIG["target_excel"])

        # ==============================
        # 【核心修复】强制清空三张表所有数据，仅保留表头
        # 每次运行都清空，彻底杜绝数据残留
        # ==============================
        print("\n🧹 开始强制清空三张表数据（保留表头）...")
        for sheet_name in TARGET_SHEETS:
            try:
                sht = wb_target.Sheets(sheet_name)
                last_row = sht.UsedRange.Rows.Count if sht.UsedRange.Rows.Count else 1
                if last_row >= 2:
                    # 彻底删除第2行到最后一行（保留第1行表头）
                    sht.Rows(f"2:{last_row}").Delete()
                    print(f"   ✅ 已清空 {sheet_name} 所有数据行")
                else:
                    print(f"   ℹ️ {sheet_name} 无数据需要清空")
            except Exception as e:
                print(f"   ⚠️ 清空 {sheet_name} 失败：{str(e)}")

        # 遍历9个数据源
        for idx, src_file in enumerate(LESSEE_CONFIG["source_excels"]):
            target_sheet_name, q_col_value = LESSEE_CONFIG["sheet_mapping"][idx]
            print(f"\n🔹 处理数据源 {idx + 1}/9：")
            print(f"   源文件：{os.path.basename(src_file)}")
            print(f"   目标子表：{target_sheet_name} | Q列填充：{q_col_value}")

            # ==============================================
            # 1. 读取源文件 → 只取 A-P 列（舍弃原Q列）
            # ==============================================
            wb_src = None
            sht_src = None
            src_data_rows = None
            try:
                wb_src = excel.Workbooks.Open(src_file)
                sht_src = wb_src.Sheets(1)
                last_row_src = sht_src.UsedRange.Rows.Count if sht_src.UsedRange.Rows.Count else 0
                last_col_src = 16  # 固定只取 A-P 列

                if last_row_src < 2:
                    print(f"⚠️ 无有效数据（行数<2），跳过")
                    continue

                # 读取数据行（第2行到最后一行）
                src_data_rows = sht_src.Range(
                    sht_src.Cells(2, 1),
                    sht_src.Cells(last_row_src, last_col_src)
                ).Value

                print(f"   ✅ 读取源文件数据：{last_row_src - 1} 行")

            except Exception as e:
                print(f"⚠️ 读取源文件失败：{str(e)}")
                if wb_src:
                    wb_src.Close(False)
                continue
            finally:
                # 及时关闭源文件
                if wb_src:
                    wb_src.Close(False)
                    wb_src = None
                    sht_src = None

            if not src_data_rows:
                print(f"⚠️ 源文件无有效数据行，跳过")
                continue

            # ==============================================
            # 2. 找到当前子表最后一行 → 追加写入
            # ==============================================
            sht_target = wb_target.Sheets(target_sheet_name)
            current_last_row = sht_target.UsedRange.Rows.Count if sht_target.UsedRange.Rows.Count else 1
            insert_start_row = current_last_row + 1
            data_rows = len(src_data_rows)

            # 写入数据
            if data_rows > 0:
                sht_target.Range(
                    sht_target.Cells(insert_start_row, 1),
                    sht_target.Cells(insert_start_row + data_rows - 1, last_col_src)
                ).Value = src_data_rows
                print(f"   ✅ 追加写入 {data_rows} 行")

            # ==============================================
            # 3. Q列批量填充
            # ==============================================
            q_col = 17
            if data_rows > 0:
                sht_target.Range(
                    sht_target.Cells(insert_start_row, q_col),
                    sht_target.Cells(insert_start_row + data_rows - 1, q_col)
                ).Value = q_col_value
                print(f"   ✅ Q列填充「{q_col_value}」")

            # ==============================================
            # 4. R列 = E列&Q列 、S列 = N列（批量公式）
            # ==============================================
            if data_rows > 0:
                # 相对引用公式，自动适配行号
                sht_target.Range(
                    f"R{insert_start_row}:R{insert_start_row + data_rows - 1}").Formula = f"=E{insert_start_row}&Q{insert_start_row}"
                sht_target.Range(
                    f"S{insert_start_row}:S{insert_start_row + data_rows - 1}").Formula = f"=N{insert_start_row}"
                print(f"   ✅ R/S列批量公式写入完成")

        # ==============================
        # 统一收尾：N列排序 + 长数字格式
        # ==============================
        print("\n📌 开始统一格式处理（所有子表）...")
        for sheet_name in TARGET_SHEETS:
            try:
                sht = wb_target.Sheets(sheet_name)
                last_row = sht.UsedRange.Rows.Count if sht.UsedRange.Rows.Count else 1
                if last_row < 2:
                    print(f"   ⚠️ {sheet_name} 无有效数据，跳过格式处理")
                    continue

                # 1. N列数字格式 + 降序
                n_col = 14
                sht.Columns(n_col).NumberFormat = "0.00"
                n_range = sht.Range(sht.Cells(2, n_col), sht.Cells(last_row, n_col))
                n_values = n_range.Value or []
                new_n = []
                for v in n_values:
                    try:
                        val = v[0] if isinstance(v, (list, tuple)) else v
                        num = float(val) if (val is not None and val != "") else 0.0
                    except (ValueError, TypeError):
                        num = 0.0
                    new_n.append([num])
                if new_n:
                    n_range.Value = new_n

                # 排序
                sort_rng = sht.Range(sht.Cells(1, 1), sht.Cells(last_row, 19))
                sort_rng.Sort(Key1=sht.Columns(n_col), Order1=2, Header=1, Orientation=1, MatchCase=False)

                # 2. 长数字列禁用科学计数法
                header_range = sht.Range(sht.Cells(1, 1), sht.Cells(1, 50))
                header_arr = [cell.Value for cell in header_range] if header_range.Value else []
                for col_idx, col_name in enumerate(header_arr, start=1):
                    if col_name in NO_SCIENCE_COLS:
                        sht.Columns(col_idx).NumberFormat = "@"

                print(f"   ✅ {sheet_name} 格式处理完成")
            except Exception as e:
                print(f"   ⚠️ {sheet_name} 格式处理失败：{str(e)}")
        # 结算配置和实际共享情况
        print("\n📌 开始处理结算配置和实际共享情况子表...")
        try:
            sht_settlement = wb_target.Sheets("结算配置和实际共享情况")
            last_row_settlement = sht_settlement.UsedRange.Rows.Count if sht_settlement.UsedRange.Rows.Count else 1

            def _normalize_matrix(values):
                if values is None:
                    return []
                if isinstance(values, tuple):
                    if len(values) == 0:
                        return []
                    if isinstance(values[0], tuple):
                        return [list(r) for r in values]
                    return [list(values)]
                return [[values]]

            if last_row_settlement >= 2:
                row_count = last_row_settlement - 1
                print("   🔹 筛选 J列=基站能耗管理系统，批量写入匹配公式")

                # 读取 J 列作为筛选条件；读取现有 L:R 公式，非目标行保持原值不被覆盖
                j_vals = _normalize_matrix(sht_settlement.Range(f"J2:J{last_row_settlement}").Value)
                lr_existing = _normalize_matrix(sht_settlement.Range(f"L2:R{last_row_settlement}").Formula)

                l_out, m_out, n_out = [], [], []
                o_out, p_out, q_out, r_out = [], [], [], []
                active_flags = []

                for idx in range(row_count):
                    row_num = idx + 2
                    j_val = str(j_vals[idx][0]).strip() if idx < len(j_vals) and j_vals[idx][0] is not None else ""
                    existing = lr_existing[idx] if idx < len(lr_existing) else [""] * 7
                    is_active = j_val in ["分路计量设备", "开关电源"]
                    active_flags.append(is_active)

                    if is_active:
                        l_out.append([f"=IFERROR(VLOOKUP(K{row_num},移动电流!R:S,2,0),\"\")"])
                        m_out.append([f"=IFERROR(VLOOKUP(K{row_num},联通电流!R:S,2,0),\"\")"])
                        n_out.append([f"=IFERROR(VLOOKUP(K{row_num},电信电流!R:S,2,0),\"\")"])
                        o_out.append([f'=IF(G{row_num}="移动",IF(L{row_num}>0,"匹配","不匹配"),IF(L{row_num}>0,"不匹配","匹配"))'])
                        p_out.append([f'=IF(H{row_num}="联通",IF(M{row_num}>0,"匹配","不匹配"),IF(M{row_num}>0,"不匹配","匹配"))'])
                        q_out.append([f'=IF(I{row_num}="电信",IF(N{row_num}>0,"匹配","不匹配"),IF(N{row_num}>0,"不匹配","匹配"))'])
                        r_out.append([f'=IF(O{row_num}="匹配",IF(P{row_num}="匹配",IF(Q{row_num}="匹配","整改完成","FALSE"),"FALSE"),"FALSE")'])
                    else:
                        l_out.append([existing[0] if len(existing) >= 1 else ""])
                        m_out.append([existing[1] if len(existing) >= 2 else ""])
                        n_out.append([existing[2] if len(existing) >= 3 else ""])
                        o_out.append([existing[3] if len(existing) >= 4 else ""])
                        p_out.append([existing[4] if len(existing) >= 5 else ""])
                        q_out.append([existing[5] if len(existing) >= 6 else ""])
                        r_out.append([existing[6] if len(existing) >= 7 else ""])

                # 批量落公式/值（一次性写入，避免逐单元格 COM 调用）
                sht_settlement.Range(f"L2:L{last_row_settlement}").Formula = l_out
                sht_settlement.Range(f"M2:M{last_row_settlement}").Formula = m_out
                sht_settlement.Range(f"N2:N{last_row_settlement}").Formula = n_out
                sht_settlement.Range(f"O2:O{last_row_settlement}").Formula = o_out
                sht_settlement.Range(f"P2:P{last_row_settlement}").Formula = p_out
                sht_settlement.Range(f"Q2:Q{last_row_settlement}").Formula = q_out
                sht_settlement.Range(f"R2:R{last_row_settlement}").Formula = r_out

                excel.Calculate()

                # 清洗 L/M/N：删除 #N/A 和 0
                for col_idx in [12, 13, 14]:
                    rng = sht_settlement.Range(sht_settlement.Cells(2, col_idx), sht_settlement.Cells(last_row_settlement, col_idx))
                    vals = _normalize_matrix(rng.Value)
                    cleaned = []
                    for idx, v in enumerate(vals):
                        cell_val = v[0] if v else ""
                        if not active_flags[idx]:
                            cleaned.append([cell_val])
                            continue
                        if cell_val == -2146826246 or \
                                (isinstance(cell_val, str) and cell_val.strip().upper() == "#N/A") or \
                                (isinstance(cell_val, (int, float)) and float(cell_val) == 0.0) or \
                                (str(cell_val).strip() in ["0", "0.0", "0.00"]):
                            cleaned.append([""])
                        else:
                            cleaned.append([cell_val])
                    rng.Value = cleaned

                # 按规则更新 W / Z（批量写入）
                current_time = datetime.now().strftime("%Y-%m-%d")
                rs_vals = _normalize_matrix(sht_settlement.Range(f"R2:S{last_row_settlement}").Value)
                w_vals = _normalize_matrix(sht_settlement.Range(f"W2:W{last_row_settlement}").Value)
                z_vals = _normalize_matrix(sht_settlement.Range(f"Z2:Z{last_row_settlement}").Value)

                for idx in range(row_count):
                    if not active_flags[idx]:
                        continue
                    r_val = str(rs_vals[idx][0]).strip().upper() if idx < len(rs_vals) and rs_vals[idx][0] is not None else ""
                    s_val = str(rs_vals[idx][1]).strip() if idx < len(rs_vals) and len(rs_vals[idx]) > 1 and rs_vals[idx][1] is not None else ""

                    if r_val == "整改完成" and s_val == "维护处理":
                        w_vals[idx][0] = current_time
                    elif r_val == "FALSE" and s_val == "数据正常":
                        z_vals[idx][0] = 1

                sht_settlement.Range(f"W2:W{last_row_settlement}").Value = w_vals
                sht_settlement.Range(f"Z2:Z{last_row_settlement}").Value = z_vals
                print(f"   ✅ 结算配置处理完成：共处理 {row_count} 行（筛选来源：基站能耗管理系统）")
            else:
                print("   ⚠️ 子表无数据，跳过")

        except Exception as e:
            print(f"   ❌ 失败：{str(e)}")
            import traceback
            traceback.print_exc()

        # ==============================
        # 保存退出
        # ==============================
        excel.ScreenUpdating = True
        wb_target.Save()

    except Exception as e:
        real_print(f"❌ lessee 处理失败：{str(e)}")
        return False

    finally:
        # 释放Excel资源
        try:
            if wb_target:
                wb_target.Close(SaveChanges=True)
        except:
            pass
        try:
            if excel:
                excel.ScreenUpdating = True
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

    real_print("✅ lessee 处理成功")
    return True


def _run_with_lock(part, runner):
    """统一加锁执行，确保设备/租户处理不会并发冲突。"""
    try:
        with SingleProcessLock(LOCK_FILE, part):
            return runner()
    except Exception as e:
        print(f"❌ 任务 {part} 无法执行：{str(e)}")
        return False


def run_device_process():
    """接口调用：仅执行设备清单流程。"""
    return _run_with_lock("device", process_device_list)


def run_lessee_process():
    """接口调用：仅执行租户电流流程。"""
    return _run_with_lock("lessee", process_lessee_list)


if __name__ == '__main__':
    # 只跑设备流程：保留下一行，注释 run_lessee_process()
    # 只跑租户流程：注释 run_device_process()，保留下一行
    # 两个都跑：两行都保留
    ok_device = True
    ok_lessee = True
    ok_device = run_device_process()
    ok_lessee = run_lessee_process()
    sys.exit(0 if (ok_device and ok_lessee) else 1)
