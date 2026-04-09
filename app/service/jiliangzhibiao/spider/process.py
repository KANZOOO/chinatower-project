import os
import sys
import pandas as pd
import xlwings as xw
import warnings
import re
from datetime import datetime, timedelta
import pywin32
warnings.filterwarnings('ignore')

# 路径
current_file = os.path.abspath(__file__)
current_dir = os.path.dirname(current_file)
parent_dir = os.path.dirname(current_dir)
grandparent_dir = os.path.dirname(parent_dir)
great_grandparent_dir = os.path.dirname(grandparent_dir)
project_root = os.path.dirname(great_grandparent_dir)
sys.path.insert(0, project_root)

BASE_PATH = project_root

CONFIG = {
    "source_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量设备清单.xls"),
    "target_excel": os.path.join(BASE_PATH, r"app/service/jiliangzhibiao/down/分路计量器设备异常清单.xlsx"),
    "cols_keep": ["室分", "微站", "昨日数据"],
    "col_N": "数据准确标识",
    "col_X": "故障时间",
    "col_T": "处理时间",
    "col_U": "新增故障数",
    "col_site_code": "站址编码",
    "today_completion_formula_pattern": r'COUNTIFS\(清单!G:G,@A13:A24,清单!T:T,"\d{4}/\d{1,2}/\d{1,2}"\)',
    "text_format_cols": ["站址编码", "站址运维ID", "FSU运维ID", "设备运维ID", "订单号", "资源铁塔编码"]
}

def get_file_modify_time(file_path):
    try:
        modify_timestamp = os.path.getmtime(file_path)
        return datetime.fromtimestamp(modify_timestamp)
    except:
        return None

def check_if_latest_file(file_path, hours=24):
    if not os.path.exists(file_path):
        return False
    t = get_file_modify_time(file_path)
    if not t:
        return False
    return t >= datetime.now() - timedelta(hours=hours)

def set_column_text_format(sheet, col_names, header_row=1):
    try:
        headers = sheet.range(f"A{header_row}").expand("right").value
        for name in col_names:
            if name in headers:
                idx = headers.index(name) + 1
                sheet.range(f"{chr(64+idx)}:{chr(64+idx)}").number_format = "@"
    except:
        pass

# def kill_excel_wps_processes():
#     """清理Excel/WPS进程，避免占用"""
#     try:
#         os.system("killall -9 'Microsoft Excel' > /dev/null 2>&1")
#         os.system("killall -9 'WPS Office' > /dev/null 2>&1")
#         os.system("killall -9 'wps' > /dev/null 2>&1")
#     except:
#         pass

def get_excel_app():
    """Mac 专用：配置 xlwings 使用正确的 Excel 路径"""
    app = None
    try:
        # 设置 xlwings 后端和 Excel 路径
        import os
        os.environ['XLWINGS_BACKEND'] = 'apple'
        
        # 显式指定 Excel 应用路径（如果需要）
        # xlwings 会自动查找 /Applications/Microsoft Excel.app
        
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        return app
    except Exception as e:
        print(f"⚠️ 启动失败: {str(e)}")
        if app:
            try:
                app.quit()
            except:
                pass
        return None

def process_device_list():
    today_str = datetime.now().strftime("%Y/%m/%d")
    print(f"🚀 开始处理 - {today_str}")

    # kill_excel_wps_processes()

    try:
        # 文件检查
        for f in [CONFIG["source_excel"], CONFIG["target_excel"]]:
            if not os.path.exists(f):
                print(f"❌ 文件不存在：{f}")
                return False

        if not check_if_latest_file(CONFIG["source_excel"]):
            print("❌ 源文件超过24小时未更新")
            return False

        # 启动
        app = get_excel_app()
        if not app:
            print("❌ 办公软件启动失败")
            return False

        # 读取源文件
        dtype = {c: str for c in CONFIG["text_format_cols"]}
        df_source = pd.read_excel(CONFIG["source_excel"], header=1, dtype=dtype)
        df_source = df_source.fillna("").astype(str).apply(lambda x: x.str.strip() if x.dtype == "object" else x)

        # 打开目标文件
        wb_target = app.books.open(CONFIG["target_excel"])

        # ===================== 核心处理 =====================
        sht_list = wb_target.sheets["清单"]
        sht_yesterday = wb_target.sheets["昨日数据"]
        sht_today = wb_target.sheets["今日数据"]

        for sh in [sht_list, sht_yesterday, sht_today]:
            set_column_text_format(sh, CONFIG["text_format_cols"])

        # 1 清单 → 昨日数据
        sht_yesterday.range("A1").value = sht_list.used_range.value
        sht_yesterday.range("A1").formula = sht_list.used_range.formula

        # 2 写入今日数据
        df_keep = df_source[CONFIG["cols_keep"]]
        sht_today.used_range.clear()
        sht_today.range("A1").value = df_keep.columns.tolist()
        sht_today.range("A2").value = df_keep.values

        # 3 清空清单
        used = sht_list.used_range
        start = 2
        if used.last_cell.row >= start:
            sht_list.range(f"A{start}:{used.last_cell.address.split('$')[-1]}").clear_contents()

        # 4 今日数据 → 清单
        df_today = pd.DataFrame(sht_today.used_range.value[1:], columns=sht_today.used_range.value[0])
        df_today = df_today.fillna("").astype(str)
        keep_cols = [c for c in df_today.columns if c not in ["室分", "微站"]]
        headers = sht_list.range("A1").expand("right").value

        for col in keep_cols:
            if col in headers:
                col_idx = headers.index(col) + 1
                sht_list.range(f"{chr(64+col_idx)}{start}").value = df_today[col].values.reshape(-1, 1)

        # 5 数据准确标识
        col_x = headers.index(CONFIG["col_X"]) + 1
        col_n = headers.index(CONFIG["col_N"]) + 1
        data = sht_list.range(f"{chr(64+col_x)}{start}").expand("down").value
        res = [["异常" if v and str(v).strip() else "准确"] for v in data]
        sht_list.range(f"{chr(64+col_n)}{start}").value = res

        # 6 对比昨日今日
        df_list = pd.DataFrame(sht_list.used_range.value[1:], columns=sht_list.used_range.value[0])
        df_yesterday = pd.DataFrame(sht_yesterday.used_range.value[1:], columns=sht_yesterday.used_range.value[0])
        site = CONFIG["col_site_code"]

        df_list[site] = df_list[site].astype(str).str.strip()
        df_yesterday[site] = df_yesterday[site].astype(str).str.strip()

        merge = pd.merge(df_list, df_yesterday[[site, CONFIG["col_N"]]], on=site, how="left", suffixes=("_t","_y"))
        merge[CONFIG["col_T"]] = merge.apply(lambda r: today_str if r[f"{CONFIG['col_N']}_y"]=="异常" and r[f"{CONFIG['col_N']}_t"]=="准确" else r[CONFIG["col_T"]], axis=1)
        merge[CONFIG["col_U"]] = merge.apply(lambda r: 1 if r[f"{CONFIG['col_N']}_y"]=="准确" and r[f"{CONFIG['col_N']}_t"]=="异常" else r[CONFIG["col_U"]], axis=1)

        col_t = headers.index(CONFIG["col_T"]) + 1
        col_u = headers.index(CONFIG["col_U"]) + 1
        sht_list.range(f"{chr(64+col_t)}{start}").value = merge[CONFIG["col_T"]].values.reshape(-1,1)
        sht_list.range(f"{chr(64+col_u)}{start}").value = merge[CONFIG["col_U"]].values.reshape(-1,1)

        # 7 更新通报公式
        sht_report = wb_target.sheets["通报"]
        formulas = sht_report.used_range.formula
        new_f = []
        for row in formulas:
            new_r = []
            for cell in row:
                if cell and isinstance(cell, str) and re.match(CONFIG["today_completion_formula_pattern"], cell):
                    cell = re.sub(r'"(\d{4}/\d{1,2}/\d{1,2})"', f'"{today_str}"', cell)
                new_r.append(cell)
            new_f.append(new_r)
        sht_report.used_range.formula = new_f

        # 保存关闭
        wb_target.save()
        wb_target.close()
        app.quit()

        print("\n✅ 处理完成！兼容Mac WPS")
        return True

    except Exception as e:
        print(f"\n❌ 失败：{str(e)}")
        try:
            if 'app' in locals() and app:
                app.quit()
        except:
            pass
        return False

if __name__ == '__main__':
    process_device_list()