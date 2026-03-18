import datetime
import pandas as pd
import numpy as np
from pathlib import Path
from core.msg.msg_text import AddressBookManagement
from core.config import settings
import re
class chaochangchaopin_fsu_msg():
    def send_msg(self, row, code):
        """发送短信通知"""
        data = {
            "city": row["市"],
            "site_name": row.get("站址名称", ""),
            "ratio":row.get("占比",""),
            "day":row["最近连续天数"]
        }

        manager_df = AddressBookManagement().get_address_book(
            city=row['市'],
            level=row["level"],
            tasks="超长超频fsu离线",
            businessCategory="一体"
        )
        if len(manager_df) > 0:
            manager_df = self.add_managers(manager_df)
            AddressBookManagement().send_msg(manager_df, data, code)

    def add_managers(self, manager_df):
        new_index = len(manager_df)
        manager_df.loc[new_index] = np.nan
        manager_df.loc[new_index, 'name'] = '陈华兴'
        manager_df.loc[new_index, 'phone'] = '13878785848'
        manager_df.loc[new_index, 'businessCategory'] = '一体'
        manager_df.loc[new_index, 'tasks'] = '超长超频fsu离线'

        return manager_df
    def run(self):
        FOLDER_PATH = settings.resolve_path(r"app/service/msg_freq_overtime_fsu_offline/data")
        def latest_consecutive_info(date_series):
            """
            输入：一个 pd.Series，元素为 date 对象、字符串或可转换为日期的类型
            输出：int —— 以“今天”为结束日的连续天数。
                  若“今天”不在数据中，返回 0。
            """
            if date_series.empty:
                return 0
            dates = pd.to_datetime(date_series).dt.normalize().drop_duplicates()
            if dates.empty:
                return 0

            today_ts =pd.Timestamp(now)
            date_set = set(dates)

            if today_ts not in date_set:
                return 0

            current = today_ts
            streak = 0
            while current in date_set:
                streak += 1
                current -= pd.Timedelta(days=1)
            return streak
        # ================== 主流程 ==================
        # 读取文件
        freq_list = []
        overtime_list = []
        now = datetime.datetime.now().date()
        # 编译正则表达式（提高效率）
        date_pattern = re.compile(r'fsu离线情况(\d{8})\.xlsx$', re.IGNORECASE)
        for file in FOLDER_PATH.glob("*.xlsx"):
            date_str = date_pattern.search(file.name).group(1)
            file_date = datetime.datetime.strptime(date_str, "%Y%m%d").date()
            if now - file_date > datetime.timedelta(days=5):
                continue
            sheet_names = pd.ExcelFile(file, engine='openpyxl').sheet_names
            for sheet_name in sheet_names:
                if sheet_name not in ['FSU离线情况明细_日(超频)', 'FSU离线情况明细_日(重要场景)']:
                    continue
                df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')
                df["市"]=df["市"].str.replace("市","").str.replace("分公司","")
                df['数据时间'] = pd.to_datetime(df['数据时间'], format='%Y-%m-%d').dt.date
                if sheet_name == 'FSU离线情况明细_日(超频)':
                    freq_list.append(df)
                elif sheet_name == 'FSU离线情况明细_日(重要场景)':
                    overtime_list.append(df)
        freq_all = pd.concat(freq_list, ignore_index=True)
        overtime_all = pd.concat(overtime_list, ignore_index=True)
        key_cols = ['省', '市', '站址名称', '站址运维ID']

        # ===== 处理超频清单：连续3/4/5天 =====
        freq_grouped = freq_all.groupby(key_cols)
        freq_summary = freq_grouped['数据时间'].apply(
            lambda x: latest_consecutive_info(x)
        ).reset_index(name='最近连续天数')
        # ===== 处理重要场景清单：24小时>离线≥12小时，比例前三连续2/4/5天 =====
        # 计算比例
        overtime_all['离线历时_数值'] = pd.to_numeric(overtime_all['离线历时（时）'], errors='coerce')
        overtime_all = overtime_all[
            (overtime_all['离线历时_数值'] >= 12) &
            (overtime_all['离线历时_数值'] < 24)
            ]
        counts = overtime_all.groupby(['市', '数据时间']).size().reset_index(name='count')
        time_totals = overtime_all.groupby('数据时间').size().reset_index(name='total')
        result = counts.merge(time_totals, on='数据时间')
        result['占比'] = result['count'] / result['total']
        result = result[['市', '数据时间', '占比']]
        result_top3 = result.groupby('数据时间').apply(
            lambda x: x.nlargest(3, '占比')
        ).reset_index(drop=True)
        result.to_excel("超长比例清单.xlsx")

        # 计算每个市连续天数
        overtime_grouped = result_top3.groupby("市")
        overtime_summary = overtime_grouped['数据时间'].apply(
            lambda x: latest_consecutive_info(x)
        ).reset_index(name='最近连续天数')
        overtime_summary=overtime_summary[overtime_summary["最近连续天数"]>0]
        result_top3['数据时间'] = pd.to_datetime(result_top3['数据时间'])
        now_ratio = (
            result_top3[result_top3['数据时间'].dt.date == now]
            [['市', '数据时间', '占比']]
        )
        overtime_summary = overtime_summary.merge(
            now_ratio,
            on=['市'],
            how='left'
        )
        overtime_summary = overtime_summary[['市', '最近连续天数', '占比']]
        overtime_summary['占比'] = (overtime_summary['占比'] * 100).round(2).astype(str) + '%'
        # 整理
        def map_level(days):
            if days == 2:
                return "二级督办对象"
            elif days == 3:
                return "二级督办对象"
            elif days == 4:
                return "三级督办对象"
            elif days == 5:
                return "四级督办对象"
            else:
                return ""

        freq_summary = freq_summary[
            freq_summary['最近连续天数'].isin([2, 3, 4, 5])
        ].reset_index(drop=True)
        overtime_summary = overtime_summary[
            overtime_summary['最近连续天数'].isin([2, 4, 5])
        ].reset_index(drop=True)
        freq_summary['level'] = freq_summary['最近连续天数'].apply(map_level)
        overtime_summary['level'] = overtime_summary['最近连续天数'].apply(map_level)
        # 发送信息
        freq_summary.apply(lambda row: self.send_msg(row, "SMS_500475369"), axis=1)
        overtime_summary.apply(lambda row: self.send_msg(row, "SMS_500530337"), axis=1)
def main():
    chaochangchaopin_fsu_msg().run()
if __name__ == '__main__':
    main()