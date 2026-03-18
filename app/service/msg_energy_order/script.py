# energy_order_msg.py
import re
import requests
import pandas as pd
import datetime
import json
from chinese_calendar import is_holiday, is_workday

from core.sql import sql_orm
from core.msg.msg_ding import DingMsg
from core.msg.msg_text import AddressBookManagement
from core.config import settings
class MsgEnergyOrder:
    def __init__(self):
        self.bustype = ['低速充电', '备电', '换电']
        self.msg_to_code = 'SMS_485350644'
        self.now = datetime.datetime.now()
        self.level_list = ['未发', '一级督办对象', '二级督办对象',
                           '三级督办对象', '四级督办对象', '五级督办对象']
        self.level = {
            '一级督办对象': 0.25, '二级督办对象': 0.5, '三级督办对象': 0.75,
            '四级督办对象': 0.85, '五级督办对象': 0.95
        }
        # 超时时间配置
        self.receive_to_timeout = {
            '一级告警': pd.Timedelta(hours=1),
            '二级告警': pd.Timedelta(hours=2)
        }
        self.response_to_timeout = {
            '一级告警': pd.Timedelta(hours=4),
            '二级告警': pd.Timedelta(hours=24)
        }
        # 白名单配置
        self.white_list = pd.read_excel(
            settings.resolve_path("app/service/msg_energy_order/data/能源设备人为原因离线台账-广西.xlsx"),
            dtype=str,
            usecols=['站址编码', '时间点', '节假日']
        )
    def _process_site_name(self, row):
        pattern = r'[a-zA-Z]'
        point_code = re.sub(pattern, '', str(row.get('点位编码', '')))
        if not point_code:
            return row.get('工单标题', '')
        return f'点位编码{point_code}'
    def update_and_send(self, df):
        for level, factor in self.level.items():
            df[f'接单{level}'] = df['建单时间'] + df['接单超时时间'] * factor
            df[f'回单{level}'] = df['建单时间'] + df['回单超时时间'] * factor
        df['接单超时时间'] = df['建单时间'] + df['接单超时时间']
        df['回单超时时间'] = df['建单时间'] + df['回单超时时间']
        df = df.loc[df.groupby('id')['接单超时时间'].idxmin()]

        df_ids = set(df['id'])
        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.msg_energy_order_level
            for site in sql.query(pojo).all():
                if any(bustype in site.id for bustype in self.bustype) and site.id not in df_ids:
                    site.level = '接单未发'

        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.msg_energy_order_level
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
                            self._send_dingding(manager_df, row, order_status)
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
            businessCategory="能源",
            specificBusiness=row['业务类型'],
            level=level,
            tasks="能源工单"
        )

        if len(manager_df) > 0:
            AddressBookManagement().send_msg(manager_df, data, self.msg_to_code)
            return manager_df
        return pd.DataFrame()
    def _send_dingding(self, manager_df, row, order_status):
        timeout_col = f'{order_status}超时时间'
        last_time = round((row[timeout_col] - self.now).total_seconds() / 3600, 2)
        order_text = (f"{row['市']}{row['区县']}，{row['站址名称']}，{row['业务类型']}工单\n"
                      f"剩余{max(last_time, 0)}小时{order_status}超时，请地市监控中心抓紧督办处理。")

        phone_list = []
        if len(manager_df) > 0:
            phone_list.extend(manager_df['phone'].tolist())
            d = DingMsg()
            d.text_at(d.LYGD, order_text, phone_list, [])
    def process_df(self, df):
        df = df.rename(columns={
            '管理城市': '市',
            '管理区县': '区县',
            '点位名称': '站址名称'
        })
        df = df.loc[df['业务类型'].isin(self.bustype)]

        df['建单时间'] = pd.to_datetime(df['建单时间'])
        df['点位编码'] = df['点位编码'].fillna('')
        df['接单人'] = 'AIOT用户'
        df['市'] = df['市'].str.replace('市', '').str.replace('分公司', '')

        df['建单时间分'] = df['建单时间'].dt.strftime('%Y-%m-%d %H:%M')
        df['站址名称']=df.apply(self._process_site_name,axis=1)
        df['id'] = df.apply(
            lambda x: f"{x['站址名称']}{x['业务类型']}{x['建单时间分']}"
            if pd.notnull(x['站址名称']) else f"{x['市']}{x['业务类型']}{x['建单时间分']}",
            axis=1
        )

        current_date = self.now.date()
        current_hour = self.now.hour
        self.white_list['时间点'] = self.white_list['时间点'].apply(lambda x: eval(x))
        self.white_list['节假日'] = self.white_list['节假日'].astype(int)
        df['删除标记'] = False

        for index, row in df.iterrows():
            if row['点位编码'] in self.white_list['站址编码'].values:
                white_row = self.white_list[self.white_list['站址编码'] == row['点位编码']].iloc[0]
                if current_hour in white_row['时间点']:
                    df.at[index, '删除标记'] = True
                if white_row['节假日'] == 1 and (is_holiday(current_date) or not is_workday(current_date)):
                    df.at[index, '删除标记'] = True

        df = df[df['删除标记'] == False].drop(columns=['删除标记'])

        df['接单超时时间'] = df['告警等级'].map(self.receive_to_timeout)
        df['回单超时时间'] = df['告警等级'].map(self.response_to_timeout)
        return df
    def down(self):
        Authorization=sql_orm().get_cookies("energy")["cookies_str"]
        cols_to_name = {
            '工单编码': 'workOrderCode',
            '工单标题': 'workOrderTitle',
            '管理城市': 'cityName',
            '管理区县': 'countyName',
            '点位名称': 'stationName',
            '点位编码': 'stationPubCode',
            '点位业务编码': 'stationCode',
            '告警等级': 'alarmLevelName',
            '建单时间': 'createTime',
            '工单状态': 'workOrderStatusName',
            '业务类型': 'businessTypeName',
            '接单考核': 'receiverAssessName',
            '回单考核': 'receiptAssessName'
        }
        headers = {
            'Content-Type': 'application/json;charset=UTF-8',
            'Authorization': Authorization
        }
        data = {
            "searchType": "1", "alarmClearTimes": [], "pageNum": 1,
            "pageSize": 10, "deptIds": [], "workType": "0",
            "timer": [], "businessType": "", "createTimes": [], "receiptTimes": []
        }
        df = pd.DataFrame(columns=cols_to_name.keys())
        while True:
            res = requests.post(
                'https://energy-iot.chinatowercom.cn/api/workorder/workOrder/page',
                headers=headers,
                data=json.dumps(data)
            )
            if res.status_code == 401:
                raise Exception("能源token过期")
            text = json.loads(res.text)
            rows = text.get('rows', [])
            if not rows:
                break

            for row in rows:
                temp_row = {k: row[v] for k, v in cols_to_name.items()}
                df = pd.concat([df, pd.DataFrame([temp_row])], ignore_index=True)

            data['pageNum'] += 1
        return df
def main():
    en=MsgEnergyOrder()
    df=en.down()
    if len(df) > 0:
        df=en.process_df(df)
    en.update_and_send(df)
if __name__ == "__main__":
    main()