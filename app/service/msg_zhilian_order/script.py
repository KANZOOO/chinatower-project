import re
import requests
import pandas as pd
import datetime
from core.sql import sql_orm
from core.msg.msg_ding import DingMsg
from core.msg.msg_text import AddressBookManagement
from core.config import settings
class ZhilianOrderMsg:
    def __init__(self):
        self.msg_to_code = 'SMS_485350644'
        self.now = datetime.datetime.now()
        self.level_list = ['未发', '一级督办对象', '二级督办对象',
                           '三级督办对象', '四级督办对象', '五级督办对象']

        self.path = settings.resolve_path(r"app/service/msg_zhilian_order/data/智联工单.xls")
        self.column = ['工单编码', '地市', '区县', '站址名称', '站址编码', '任务名称',
                       '创建时间', '超时预警时限(小时)', '工单处理人', '工单状态', '业务分类']
        self.column_rename = {
            '地市': '市',
            '创建时间': '建单时间',
            '工单处理人': '接单人',
            '业务分类': '业务类型',
            '超时预警时限(小时)': '回单超时时间'
        }
        self.pardon_list = [
            "451402000010002128", "451421908000000291", "45142400001327000005",
            "45142100831327000014", "45142100001327000001", "45142100831327000012",
            "1101021000597048", "45142100831327000011", "45142100004320000001",
            "45142100004320000002", "45140200831327000014"
        ]

    def _process_site_name(self, row):
        site_name = str(row.get('站址名称', ''))
        for char in ['(', '（', '/', '#', '_', '－']:
            if char in site_name:
                site_name = site_name.split(char)[0].strip()
        site_name = site_name.replace(' ', '')
        site_name = re.sub(r'[a-zA-Z]', '', site_name)
        return site_name

    def _update_send_status(self, df):
        levels = {
            '一级督办对象': 0.25, '二级督办对象': 0.5, '三级督办对象': 0.75,
            '四级督办对象': 0.85, '五级督办对象': 0.95
        }
        for level, factor in levels.items():
            df[f'接单{level}'] = df['建单时间'] + df['接单超时时间'] * factor
            df[f'回单{level}'] = df['建单时间'] + df['回单超时时间'] * factor

        df['接单超时时间'] = df['建单时间'] + df['接单超时时间']
        df['回单超时时间'] = df['建单时间'] + df['回单超时时间']
        df = df.loc[df.groupby('id')['接单超时时间'].idxmin()]

        df_ids = set(df['id'])
        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.msg_zhilian_order_level
            for site in sql.query(pojo).all():
                if site.id not in df_ids:
                    site.level = '接单未发'

        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.msg_zhilian_order_level
            for index, row in df.iterrows():
                res = sql.query(pojo).filter(pojo.id == row.id).first()
                order_status = row.工单状态[1:]  # "待接单" → "接单"
                if res is None:
                    temp = pojo()
                    temp.id = row.id
                    temp.level = f'{order_status}未发'
                    temp.send_time = row.建单时间
                    sql.merge(temp)
                else:
                    now_level = res.level
                    if order_status != now_level[:2]:
                        res.level = f"{order_status}{self.level_list[0]}"
                    elif '五级督办' not in now_level:
                        current_level_index = self.level_list.index(now_level[2:])
                        next_level = f"{order_status}{self.level_list[current_level_index + 1]}"
                        if res.send_time <= getattr(row, next_level) <= self.now:
                            res.level = next_level
                            res.send_time = self.now
                            sql.commit()
                            level = next_level[2:]
                            manager_df = self._send_msg(row, order_status, level)
                            self._send_dingding(manager_df, row, order_status, level)

    def _send_msg(self, row, order_status, level):
        data = {
            'city': row['市'],
            'area': row['区县'].replace('区', ''),
            'name': row['接单人'],
            'site_name': row['站址名称'],
            'type': row['业务类型'],
            'operate_type': order_status
        }

        timeout_col = f'{order_status}超时时间'
        last_time = round((row[timeout_col] - self.now).total_seconds() / 3600, 2)
        data['time'] = f"{max(last_time, 0)}小时"

        manager_df = AddressBookManagement().get_address_book(
            city=row['市'],
            area=row['区县'],
            businessCategory="拓展",
            level=level,
            tasks="智联工单提醒"
        )

        if len(manager_df) > 0:
            if row["站址编码"] in self.pardon_list:
                manager_df = manager_df[manager_df["level"].isin(["一级督办对象", "二级督办对象"])]
            AddressBookManagement().send_msg(manager_df, data, self.msg_to_code)
            return manager_df
        return pd.DataFrame()

    def _send_dingding(self, manager_df, row, order_status, level):
        timeout_col = f'{order_status}超时时间'
        last_time = round((row[timeout_col] - self.now).total_seconds() / 3600, 2)
        order_text = (f"{row['市']}{row['区县']}，{row['站址名称']}，{row['业务类型']}工单\n"
                      f"剩余{max(last_time, 0)}小时{order_status}超时，请地市监控中心抓紧督办处理。")

        phone_list = []
        if len(manager_df) > 0:
            phone_list.extend(manager_df['phone'].tolist())
            d = DingMsg()
            d.text_at(d.LYGD, order_text, phone_list, [])

    def _get_timeout(self, station_code):
        special_stations = [
            451002500000000069, 451022908000000385, 451024500000000034,
            451026500000000038, 451027908000000202, 451028908000000003,
            451030500000001359, 451031908000000264, 451025908000000093,
            450502500000000051, 450501908000000597, 45052101000085,
            451402500010001834, 45142301000034, 45142501000016,
            450602908000000303, 45060201000117, 450681908000000303,
            450821908000000696, 450881500000000353, 450324908001900283,
            450325600000000830, 450328908000000016, 451281500000000014,
            45122101000021, 451227908000000156, 451228500000000011,
            451102500000000195, 451121500000000118, 451302500000000066,
            450203908000000052, 45022101000046, 450224500000001545,
            45010201000036, 450103500000000190, 450107500000000325,
            450108908000000236, 450123500000000093, 450124700000053164,
            450125500000000106, 45012701000126, 450702500000000612,
            450701908000000014, 450721908000000655, 45072100000093,
            450722908000000088, 450722908000000107, 450403500000001533,
            450423908000000137, 450481908000000518, 45092201000029,
            450924908000000387, 450923908000000034, 450923908000000431,
            450981500000000043
        ]
        return pd.Timedelta(hours=4) if station_code in special_stations else pd.Timedelta(hours=8)

    def _down_order_msg(self):
        INTO_DATA = {'BUSI_TYPE': '2',
                      '_': '1749349542771',
                      'orgId': '0098364',
                      'pageName': 'taskListIndex',
                      'queryType': '1'}

        URL = 'http://omms.chinatowercom.cn:9000/portal/SelfTaskController/exportExcel'
        headers = {
            'Host': 'omms.chinatowercom.cn:9000',
            'Origin': 'http://omms.chinatowercom.cn:9000',
            'Referer': 'http://omms.chinatowercom.cn:9000/portal/iframe.html?modules/domian/views/listIndexTask',
            'Cookie': '',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
        }
        headers['Cookie'] = sql_orm().get_cookies("foura")["cookies_str"]
        res = requests.get(URL, headers=headers, params=INTO_DATA)
        with open(self.path, "wb") as f:
            f.write(res.content)

    def _process_df(self):
        df = pd.read_excel(self.path, dtype=str, usecols=self.column)
        df = df.rename(columns=self.column_rename)
        df = df.reset_index(drop=True)
        df['建单时间'] = pd.to_datetime(df['建单时间'])
        df = df[df['工单状态'].isin(['已领取', '待领取'])]
        df['工单状态'] = df['工单状态'].replace({'已领取': '待回单', '待领取': '待接单'})

        df['站址名称'] = df.apply(self._process_site_name, axis=1)
        df['id'] = df['工单编码'] + df['业务类型']

        df['接单人'] = df['接单人'].fillna('未知')
        df['市'] = df['市'].str.replace('市', '').str.replace('分公司', '')

        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.station
            site_codes = df['站址编码'].tolist()
            results = sql.query(pojo).filter(pojo.site_code.in_(site_codes)).all()
            area_map = {r.site_code: r.area for r in results}
            df['区县'] = df['站址编码'].map(area_map).fillna('')

        df['接单超时时间'] = pd.Timedelta(hours=2)
        df['回单超时时间'] = df['站址编码'].apply(lambda x: self._get_timeout(int(x) if x else 0))

        self._update_send_status(df)

    def run(self):
        self._down_order_msg()
        self._process_df()
def main():
    ZhilianOrderMsg()
    ZhilianOrderMsg().run()
if __name__ == "__main__":
    main()