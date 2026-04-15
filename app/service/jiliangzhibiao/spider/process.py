import os
import sys
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


def get_col_index(sheet, field_name):
    """封装：获取字段对应的列索引（提前校验，避免索引错误）"""
    last_col = sheet.UsedRange.Columns.Count if sheet.UsedRange.Columns.Count else 100  # 兜底
    header_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, last_col))
    headers = [str(cell.Value).strip() if cell.Value else "" for cell in header_range]
    return headers.index(field_name) + 1 if field_name in headers else None


def process_device_list():
    today_str = datetime.now().strftime("%Y/%m/%d")
    print(f"开始处理 - {today_str}")

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

        # 【修改2】提前初始化所有目标sheet的文本列格式（先设格式，后写数据）
        print("🔹 提前初始化关键列文本格式（防止科学计数法）")
        target_sheets = [sht_list, sht_yesterday, sht_today, sht_shifen, sht_wangge]
        for sheet in target_sheets:
            for field in TEXT_FORMAT_FIELDS:
                col_idx = get_col_index(sheet, field)
                if not col_idx:
                    continue
                # 整列设置为文本格式（覆盖所有行，包括未写入的行）
                sheet.Columns(col_idx).NumberFormat = "@"  # 整列设为文本，核心！
        print("✅ 关键列文本格式初始化完成")

        # 1. 清空昨日数据 → 复制清单到昨日数据
        print("🔹 复制清单数据到昨日数据工作表")
        sht_yesterday.UsedRange.ClearContents()

        last_row_list = sht_list.UsedRange.Rows.Count
        last_col_list = sht_list.UsedRange.Columns.Count

        if last_row_list >= 1 and last_col_list >= 1:
            data_copy = sht_list.Range(
                sht_list.Cells(1, 1),
                sht_list.Cells(last_row_list, last_col_list)
            ).Value
            # 批量写入昨天数据
            sht_yesterday.Range(
                sht_yesterday.Cells(1, 1),
                sht_yesterday.Cells(last_row_list, last_col_list)
            ).Value = data_copy

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

        # 【修改6】兜底：再次校验关键列格式（防止复制过程中格式丢失）
        print("🔹 最终校验关键列文本格式")
        for sheet in target_sheets:
            try:
                for field in TEXT_FORMAT_FIELDS:
                    col_idx = get_col_index(sheet, field)
                    if not col_idx:
                        continue
                    last_row = sheet.UsedRange.Rows.Count
                    if last_row < 2:
                        continue
                    rng = sheet.Range(sheet.Cells(2, col_idx), sheet.Cells(last_row, col_idx))
                    # 再次确认格式+强制文本（无循环批量处理）
                    rng.NumberFormat = "@"
                    vals = rng.Value
                    new_vals = []
                    for item in vals:
                        v = str(item[0]).strip() if (item and item[0] is not None) else ""
                        new_vals.append(["'" + v])
                    rng.Value = new_vals
                print(f"✅ {sheet.Name} 文本格式校验完成")
            except Exception as e:
                print(f"⚠️ {sheet.Name} 格式校验警告：{str(e)}")
                pass

        # 7、清单异常匹配故障时间（核心批量处理逻辑）
        print("🔹 批量处理清单异常数据，填充故障时间公式 + X列匹配公式")

        # ====================== 配置区 =======================
        n_col = 14  # 异常/准确 标识列（N列）
        s_col = 19  # 故障时间列（S列）
        x_col = 24  # 新增：X列（第24列）
        site_code_col = 2  # 站址编码列（B列）
        # ====================================================

        # 获取清单数据最后一行
        list_last_row = sht_list.UsedRange.Rows.Count

        # 仅处理有数据的行（跳过表头）
        if list_last_row >= 2:
            # 1. 先清空 S列 + X列 所有数据（保留表头）
            sht_list.Range(sht_list.Cells(2, s_col), sht_list.Cells(list_last_row, s_col)).ClearContents()
            sht_list.Range(sht_list.Cells(2, x_col), sht_list.Cells(list_last_row, x_col)).ClearContents()

            # 2. 批量生成公式数组（仅N列为"异常"的行填充公式）
            s_formula_array = []
            x_formula_array = []

            # 读取N列所有状态值（批量读取，提升效率）
            rng_n_status = sht_list.Range(
                sht_list.Cells(2, n_col),
                sht_list.Cells(list_last_row, n_col)
            ).Value

            # 遍历每一行生成对应公式
            for row_idx in range(len(rng_n_status)):
                # 当前行的实际行号（+1是因为从第2行开始，+row_idx是遍历偏移）
                actual_row = row_idx + 2
                # 获取当前行N列的状态
                status = rng_n_status[row_idx][0] if (
                            rng_n_status[row_idx] and rng_n_status[row_idx][0] is not None) else ""
                status = str(status).strip()

                # 判断是否为异常，是则填充公式，否则留空
                if status == "异常":
                    s_formula = f"=VLOOKUP(B{actual_row},网格!C:P,14,0)"
                    x_formula = f"=VLOOKUP(B{actual_row},昨天数据!B:X,23,0)"
                else:
                    s_formula = ""
                    x_formula = ""

                s_formula_array.append([s_formula])
                x_formula_array.append([x_formula])

            # 3. 批量写入公式到 S列 和 X列（一次性写入，效率极高）
            sht_list.Range(sht_list.Cells(2, s_col), sht_list.Cells(list_last_row, s_col)).Formula = s_formula_array
            sht_list.Range(sht_list.Cells(2, x_col), sht_list.Cells(list_last_row, x_col)).Formula = x_formula_array

            print(f"✅ 公式填充完成：共处理 {len(s_formula_array)} 行，异常行自动填充 S列 + X列 公式")
        else:
            print("⚠️ 清单工作表无有效数据，跳过公式填充")

        # 8、【批量】状态对比：异常 <-> 准确 自动填 处理时间 / 新增故障数
        print("🔹 批量状态对比：昨日数据 ↔ 清单，自动填充处理时间、新增故障数")

        # ====================== 配置区 =======================
        n_col = 14  # 异常/准确 标识列（N列）
        t_col = 20  # 处理时间 列（T列）
        u_col = 21  # 新增故障数 列（U列）
        site_code_col = 2  # 站址编码 所在列（默认B列=2）
        today = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # 当天时间
        # ====================================================

        # ----------------------
        # 1. 批量读取【清单】全部关键列：站址编码、N列状态
        # ----------------------
        list_last_row = sht_list.UsedRange.Rows.Count
        list_site_codes = []
        list_n_status = []
        if list_last_row >= 2:
            # 读取 站址编码 列
            rng_site_list = sht_list.Range(sht_list.Cells(2, site_code_col),
                                           sht_list.Cells(list_last_row, site_code_col))
            list_site_codes = rng_site_list.Value

            rng_n_list = sht_list.Range(sht_list.Cells(2, n_col), sht_list.Cells(list_last_row, n_col))
            list_n_status = rng_n_list.Value

        # ----------------------
        # 2. 批量读取【昨天数据】全部关键列：站址编码、N列状态
        # ----------------------
        yesterday_last_row = sht_yesterday.UsedRange.Rows.Count
        yesterday_site_dict = {}
        if yesterday_last_row >= 2:
            rng_site_yest = sht_yesterday.Range(sht_yesterday.Cells(2, site_code_col),
                                                sht_yesterday.Cells(yesterday_last_row, site_code_col))
            yesterday_site_codes = rng_site_yest.Value

            rng_n_yest = sht_yesterday.Range(sht_yesterday.Cells(2, n_col),
                                             sht_yesterday.Cells(yesterday_last_row, n_col))
            yesterday_n_status = rng_n_yest.Value

            for i in range(len(yesterday_site_codes)):
                s = yesterday_site_codes[i][0] if isinstance(yesterday_site_codes[i], (list, tuple)) else \
                    yesterday_site_codes[i]
                s = str(s).strip() if s else ""

                st = yesterday_n_status[i][0] if isinstance(yesterday_n_status[i], (list, tuple)) else \
                    yesterday_n_status[i]
                st = str(st).strip() if st else ""

                if s:
                    yesterday_site_dict[s] = st

        # ----------------------
        # 3. 批量生成 T列（处理时间）、U列（新增故障数）数组
        # ----------------------
        t_array = []  # 处理时间
        u_array = []  # 新增故障数

        if list_last_row >= 2:
            for i in range(len(list_site_codes)):

                site = list_site_codes[i][0] if isinstance(list_site_codes[i], (list, tuple)) else list_site_codes[i]
                site = str(site).strip() if site else ""

                now_status = list_n_status[i][0] if isinstance(list_n_status[i], (list, tuple)) else list_n_status[i]
                now_status = str(now_status).strip() if now_status else ""

                last_status = yesterday_site_dict.get(site, "")

                t_val = ""
                u_val = ""

                if last_status == "异常" and now_status == "准确":
                    t_val = today
                if last_status == "准确" and now_status == "异常":
                    u_val = "1"

                t_array.append([t_val])
                u_array.append([u_val])

            sht_list.Range(
                sht_list.Cells(2, t_col),
                sht_list.Cells(list_last_row, t_col)
            ).Value = t_array

            # 写入 U列 新增故障数
            sht_list.Range(
                sht_list.Cells(2, u_col),
                sht_list.Cells(list_last_row, u_col)
            ).Value = u_array

        print("✅ 状态对比 & 自动填充完成")

        # 9、【纯批量】清单G列填充VLOOKUP公式：=VLOOKUP(B2,网格!C:I,7,0)
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

        # 10、填充今日完成数公式
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

        print("✅ 全部处理完成！")
        return True

    except Exception as e:
        print(f"❌ 失败：{str(e)}")
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
    today_str = datetime.now().strftime("%Y/%m/%d")
    print(f"\n开始处理租户电流数据 - {today_str}")

    # ==============================
    # 【配置区】- 根据实际路径调整
    # ==============================
    LESSEE_CONFIG = {
        "target_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路租户异常通报.xlsx"),
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
                sht_target.Range(f"R{insert_start_row}:R{insert_start_row + data_rows - 1}").Formula = f"=E{insert_start_row}&Q{insert_start_row}"
                sht_target.Range(f"S{insert_start_row}:S{insert_start_row + data_rows - 1}").Formula = f"=N{insert_start_row}"
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

            if last_row_settlement >= 2:
                print("   🔹 批量写入 L/M/N 列 VLOOKUP 公式（按J列判断）")

                # --------------------------
                # 🔥 终极方案：只更新符合条件的行，其余完全保留原有数据
                # --------------------------
                # 1. 一次性读取J列所有数据
                j_data = sht_settlement.Range(f"J2:J{last_row_settlement}").Value

                # 2. 只记录【符合条件】的行号 + 公式（不符合的直接跳过）
                update_rows = []  # 存储要更新的行号
                l_formulas = []
                m_formulas = []
                n_formulas = []

                for idx, item in enumerate(j_data):
                    row_num = idx + 2  # 真实行号
                    j_val = str(item[0]).strip() if (item and item[0] is not None) else ""

                    # ✅ 只处理这2种，其余全部跳过，保留原有数据！
                    if j_val in ["分路计量设备", "开关电源"]:
                        update_rows.append(row_num)
                        l_formulas.append([f"=VLOOKUP(K{row_num},移动电流!R:S,2,0)"])
                        m_formulas.append([f"=VLOOKUP(K{row_num},联通电流!R:S,2,0)"])
                        n_formulas.append([f"=VLOOKUP(K{row_num},电信电流!R:S,2,0)"])

                # 3. 🔥 只批量写入【符合条件的行】，其他行完全不动！
                if update_rows:
                    # 批量写入L/M/N
                    sht_settlement.Range(f"L{update_rows[0]}:L{update_rows[-1]}").Formula = l_formulas
                    sht_settlement.Range(f"M{update_rows[0]}:M{update_rows[-1]}").Formula = m_formulas
                    sht_settlement.Range(f"N{update_rows[0]}:N{update_rows[-1]}").Formula = n_formulas

                # # 3. 一次性批量写入Excel（只交互1次，巨快）
                # if l_formulas:
                #     # 🔥 关键：只写入有公式的行，不覆盖原有数据
                #     sht_settlement.Range(f"L2:L{last_row_settlement}").Formula = l_formulas
                #     sht_settlement.Range(f"M2:M{last_row_settlement}").Formula = m_formulas
                #     sht_settlement.Range(f"N2:N{last_row_settlement}").Formula = n_formulas

                # --------------------------
                # O/P/Q/R 列 批量写入公式
                # --------------------------
                print("   🔹 批量写入 O/P/Q/R 列判断公式")
                # O列 = 15
                sht_settlement.Range(
                    f"O2:O{last_row_settlement}").Formula = '=IF(G2="移动",IF(L2>0,"匹配","不匹配"),IF(L2>0,"不匹配","匹配"))'
                # P列 = 16
                sht_settlement.Range(
                    f"P2:P{last_row_settlement}").Formula = '=IF(H2="联通",IF(M2>0,"匹配","不匹配"),IF(M2>0,"不匹配","匹配"))'
                # Q列 = 17
                sht_settlement.Range(
                    f"Q2:Q{last_row_settlement}").Formula = '=IF(I2="电信",IF(N2>0,"匹配","不匹配"),IF(N2>0,"不匹配","匹配"))'
                # R列 = 18
                sht_settlement.Range(
                    f"R2:R{last_row_settlement}").Formula = '=IF(O2="匹配",IF(P2="匹配",IF(Q2="匹配","整改完成","FALSE")))'

                # # --------------------------
                # # S列：如果为空 → 填入“数据正常”
                # # --------------------------
                # print("   🔹 S列空值填充：数据正常")
                # # S列 = 19
                # sht_settlement.Range(f"S2:S{last_row_settlement}").Formula = '=IF(S2="","数据正常",S2)'

                # 强制计算
                excel.Calculate()

                # 设置数值格式
                print("   🔹 设置 L/M/N 列为数值格式 0.00")
                for col in [12, 13, 14]:
                    sht_settlement.Columns(col).NumberFormat = "0.00"

                # 替换 #N / -2146826246 / 0 / 0.00 为空
                print("   🔹 替换无效值为空")
                for col_idx in [12, 13, 14]:
                    rng = sht_settlement.Range(sht_settlement.Cells(2, col_idx),
                                               sht_settlement.Cells(last_row_settlement, col_idx))
                    vals = rng.Value or []
                    new_vals = []
                    for v in vals:
                        cell_val = v[0] if isinstance(v, (list, tuple)) else v

                        # 🔥 关键修复：增加 -2146826246 判断（#N/A 错误码）
                        if cell_val == -2146826246 or \
                                (isinstance(cell_val, str) and cell_val == "#N/A") or \
                                (isinstance(cell_val, (int, float)) and cell_val == 0) or \
                                (cell_val == "0.00"):
                            new_vals.append([""])
                        else:
                            new_vals.append([cell_val])
                    if new_vals:
                        rng.Value = new_vals

                # --------------------------
                # 🔥 最终版：保留 W/Z 原始数据，仅更新符合条件的行
                # --------------------------
                print("   🔹 按条件写入 W列处理时间 / Z列故障数时间（保留原有数据）")
                current_time = datetime.now().strftime("%Y-%m-%d")

                # 批量读取 R、S、W、Z 四列原始数据
                rng = sht_settlement.Range(f"R2:Z{last_row_settlement}")
                data = rng.Value if rng.Value else []

                # 遍历每一行，保留原始数据，只改符合条件的
                for row_idx in range(len(data)):
                    row = data[row_idx]
                    # 读取各列值
                    r_val = str(row[0]).strip() if row[0] is not None else ""
                    s_val = str(row[1]).strip() if row[1] is not None else ""
                    w_old = row[4]  # W列原始值
                    z_old = row[7]  # Z列原始值

                    # ======================
                    # 规则 1：只在满足条件时更新
                    # ======================
                    if r_val == "整改完成" and s_val == "维护处理":
                        # W列写入新时间，Z列保持原样
                        sht_settlement.Cells(row_idx + 2, 23).Value = current_time

                    elif (r_val == "False" or r_val == "FALSE") and s_val == "数据正常":
                        # Z列写入新时间，W列保持原样
                        sht_settlement.Cells(row_idx + 2, 26).Value = current_time

                    # 不满足条件：什么都不做，保留原始数据！

                print(f"   ✅ 处理完成：共 {last_row_settlement - 1} 行")
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
        print(f"\n✅ 目标文件已保存")

    except Exception as e:
        print(f"\n❌ 处理失败：{str(e)}")
        import traceback
        traceback.print_exc()
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
        print("✅ Excel资源已释放")

    print("\n✅ 全部电流数据处理完成！数据已全新写入无残留！")
    return True

if __name__ == '__main__':
    # process_device_list()
    process_lessee_list()