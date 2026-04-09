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
        print("🔹 批量处理清单异常数据，填充故障时间公式")

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


if __name__ == '__main__':
    process_device_list()