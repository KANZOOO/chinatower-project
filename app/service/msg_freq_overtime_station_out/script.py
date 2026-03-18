import os
import re
import json
import datetime
import pandas as pd
import requests
from bs4 import BeautifulSoup
from sqlalchemy import text, and_
from sqlalchemy.testing.plugin.plugin_base import engines

from core.sql import sql_orm
from core.msg.msg_text import AddressBookManagement
from core.config import settings
from app.service.msg_freq_overtime_station_out import schema
class MsgFreqOvertimeStationOut:
    def __init__(self):
        self.down_path = settings.resolve_path("spider/down/oracle_data/data_alarm_his_out.csv")
        self.msg_to_code = {
            '退服即将超长': 'SMS_471455419',
            '退服已超长': 'SMS_471260485'
        }
        self.msg_to_code_chaopin = {
            '退服即将超频': {1: 'SMS_478530457', 2: 'SMS_478385497', 3: 'SMS_478590515', 4: 'SMS_478460544'},
            '退服已超频': {1: 'SMS_478620464', 2: 'SMS_478570484', 3: 'SMS_478445474', 4: 'SMS_478470484', 5: 'SMS_478380467'}
        }
        self.level_to_msg = {
            3: '退服即将超长', 6: '退服即将超长', 8: '退服已超长', 36: '退服已超长', 40: '退服已超长', 48: '退服已超长',
            '2d': '退服即将超频', '3d': '退服即将超频', '4d': '退服即将超频', '5d': '退服已超频'
        }
        self.level_to_manager = {
            0: [], 3: ['一级督办对象'], 6: ['二级督办对象'], 8: ['二级督办对象'], 36: ['三级督办对象'], 40: ['四级督办对象'], 48: ['五级督办对象'],
            '2d': ['一级督办对象', '二级督办对象'], '3d': ['三级督办对象'], '4d': ['四级督办对象'], '5d': ['五级督办对象']
        }
        self.time = [48, 40, 36, 8, 6, 3, 0]

    def get_alarm_history(self):
        cookies=sql_orm().get_cookies("foura")["cookies"]
        res = requests.get(schema.URL, headers=schema.HEADERS, cookies=cookies)
        javax = BeautifulSoup(res.text, 'html.parser', from_encoding='utf-8').find('input', id='javax.faces.ViewState')['value']

        for name in dir(schema):
            if '__' in name or 'INTO_DATA' not in name:
                continue
            into_data = getattr(schema, name)
            if 'INTO_DATA1' in name:
                into_data.update({
                    'queryForm:queryalarmName': '一级低压脱离告警',
                    'queryForm:firstendtimeInputCurrentDate': datetime.datetime.now().strftime("%m/%Y"),
                    'queryForm:firststarttimeInputCurrentDate': datetime.datetime.now().strftime("%m/%Y"),
                    'queryForm:firstendtimeInputDate': datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                    'queryForm:firststarttimeInputDate': datetime.datetime.now().replace(day=1, hour=0, minute=0, second=0).strftime("%Y-%m-%d %H:%M")
                })
            into_data['javax.faces.ViewState'] = javax
            res = requests.post(schema.URL, headers=schema.HEADERS, data=into_data, cookies=cookies)
            if 'INTO_DATA_FINAL' in name:
                with open(self.down_path, "wb") as f:
                    f.write(res.content)
        return 0

    def get_time_reduce(self, time_last):
        end_time = datetime.datetime.now()
        start_time = end_time - datetime.timedelta(minutes=time_last)
        days = end_time.day - start_time.day
        if days == 0: days = -1
        start_alter = (start_time.replace(hour=6, minute=0, second=0) - start_time).total_seconds() / 60 if start_time.hour < 6 else 0
        end_alter = (end_time - end_time.replace(hour=0, minute=0, second=0)).total_seconds() / 60 if end_time.hour < 6 else 0
        reduce = start_alter + end_alter + days * 360
        return time_last - reduce if reduce > 0 else time_last

    def process(self):
        #########################子函数#######################
        def process_site_name(row):
            name = row['site_name'] or f"{row['city']}{row['area']}"
            if pd.isnull(name):
                name = f"site_code{row['site_code']}"
            name = re.sub(r'[a-zA-Z]', '', name.split('(')[0].split('（')[0].split('/')[0].split('#')[0].split('_')[0].split('-')[0].strip().replace(' ', ''))
            return name
        def get_level(data):
            return next((i for i in self.time if i * 60 <= data), 0)
        #########################获取活动告警#######################
        with sql_orm().session_scope() as temp:
            sql, Base = temp
            res = sql.execute(text("select * from data_alarm_now where alarm_name in ('一级低压脱离告警')"))
            colnames = [column[0] for column in res.cursor.description]
            rows = res.fetchall()

        if not rows:
            return "无活动告警"
        df = pd.DataFrame([dict(zip(colnames, row)) for row in rows])

        df['broken_time'] = df['broken_time'].astype('int32')
        df['site_name'] = df.apply(process_site_name, axis=1)
        df['level'] = df['broken_time'].apply(get_level)
        df['发送超长'] = ''
        df['发送超频'] = 0
        #########################初始数据库操作#######################
        with sql_orm().session_scope() as temp:
            sql, Base = temp
            overtime_reduce = Base.classes.msg_freq_overtime_station_out_reduce
            overtime = Base.classes.msg_freq_overtime_station_out_status
            # 更新超时记录
            now_month = datetime.datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            happen_times = pd.to_datetime(df['happen_time'], errors='coerce')
            if happen_times.lt(now_month).any():
                return "活动告警停留在上月，影响超频判断"
            res = sql.query(overtime).filter(overtime.broken_times != 0, ~overtime.begin_time.like(f"{now_month.strftime('%Y-%m')}%")).all()
            for item in res:
                item.broken_times = 0
            # 没有告警的还原
            res = sql.query(overtime).all()
            for item in res:
                if not (df['site_name'] == item.site_name).any():
                    item.level = 0
            # 时间减免站址
            for index, row in df.iterrows():
                res = sql.query(overtime_reduce).filter(overtime_reduce.sitecode == row['site_code']).first()
                if res:
                    df.loc[index, 'broken_time'] = self.get_time_reduce(row['broken_time'])

            #########################判定超长超频+更新数据库#######################
            alarm_history = pd.read_csv(self.down_path, dtype=str).rename(columns={'站址': 'site_name'})
            for index, row in df.iterrows():
                res = sql.query(overtime).filter(overtime.site_name == row['site_name']).first()
                if res is None:
                    row['level']=0 # 逐级升级，避免跳级
                    df.loc[index, '发送超长'] = row['level']
                    temp = overtime(site_name=row['site_name'], level=row['level'], broken_times=1, begin_time=row['happen_time'])
                    sql.merge(temp)
                else:
                    if res.level < row['level']:
                        if res.level not in self.time:
                            # 处理非法 level，比如跳过或重置
                            continue
                        current_index = self.time.index(res.level)
                        if current_index > 0:
                            row['level'] = self.time[current_index - 1] # 逐级升级，避免跳级
                        df.loc[index, '发送超长'] = row['level']
                        res.level = row['level']
                    res.begin_time = row['happen_time']

                    df_chaopin = alarm_history[alarm_history['site_name'] == row['site_name']]
                    if not df_chaopin.empty:
                        df_chaopin['告警发生时间']=df_chaopin['首次告警时间']
                        df_chaopin['告警发生时间'] = pd.to_datetime(df_chaopin['告警发生时间'])
                        df_chaopin['告警发生日期'] = df_chaopin['告警发生时间'].dt.strftime('%Y-%m-%d')
                        df_chaopin = df_chaopin.groupby('告警发生日期').size().reset_index(name='次数')
                        current_date = datetime.datetime.now().strftime('%Y-%m-%d')
                        if current_date in df_chaopin['告警发生日期'].values:
                            df_chaopin.loc[df_chaopin['告警发生日期'] == current_date, '次数'] += 1
                        else:
                            df_chaopin.loc[len(df_chaopin)] = [current_date, 1]
                        broken_times = df_chaopin['次数'].sum()

                        if res.broken_times < broken_times:
                            broken_times=res.broken_times+1# 在原基础上+1，避免跳过级别
                            broken_days=(len(df_chaopin) if len(df_chaopin)<=5 else 5)
                            df.loc[index, '涉及天数'] = broken_days
                            df.loc[index, '发送超频'] = broken_times
                            for i, row_chaopin in df_chaopin.iterrows():
                                temp = datetime.datetime.strptime(row_chaopin['告警发生日期'], '%Y-%m-%d')
                                df.loc[index, f'month{str(i + 1)}'] = temp.month
                                df.loc[index, f'day{str(i + 1)}'] = temp.day
                                df.loc[index, f'num{str(i + 1)}'] = row_chaopin['次数']
                            res.broken_times = broken_times
            sql.commit()
        #########################发送信息#######################
        df.fillna('', inplace=True)
        df.apply(self.send_msg_chaochang, axis=1)
        df.apply(self.send_msg_chaopin, axis=1)

    def send_msg_chaopin(self, row):
        broken_times = row['发送超频']
        if broken_times < 2 or broken_times > 5:
            return

        level_origin = str(broken_times) + 'd'
        code = self.msg_to_code_chaopin[self.level_to_msg[level_origin]][row['涉及天数']]
        data = {'site_name': row['site_name'], 'broken_times': int(broken_times),'month': int(row['month1']),'day':int(row['day1']),'num':int(row['num1'])}
        for i in range(2, 6):
            try:
                data[f'month{i}'] = int(row[f'month{i}'])
                data[f'day{i}'] = int(row[f'day{i}'])
                data[f'num{i}'] = int(row[f'num{i}'])
            except:
                pass
        manager_df_list=[]
        for level in self.level_to_manager[level_origin]:
            manager_df_list.append(AddressBookManagement().get_address_book(city=row['city'], area=row['area'], level=level,businessCategory="一体",tasks="超频退服"))
        manager_df=pd.concat(manager_df_list)
        AddressBookManagement().send_msg(manager_df, data, code)

    def send_msg_chaochang(self, row):
        level_origin = row['发送超长']
        if level_origin == '' or level_origin == 0:
            return

        broken_times = row['发送超频']
        code = self.msg_to_code[self.level_to_msg[level_origin]]
        data = {
            'sitename': row['site_name'],
            'time': str(round(row['broken_time'] / 60, 1)) + '小时',
            'broken_times': int(broken_times)
        }
        manager_df_list=[]
        for level in self.level_to_manager[level_origin]:
            manager_df_list.append(AddressBookManagement().get_address_book(city=row['city'], area=row['area'], level=level,businessCategory="一体",tasks="超长退服"))
        manager_df=pd.concat(manager_df_list)
        AddressBookManagement().send_msg(manager_df, data, code)

def main():
    MsgFreqOvertimeStationOut().process()
if __name__ == '__main__':
    main()