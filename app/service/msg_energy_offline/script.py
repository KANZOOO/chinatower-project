import requests
import json
import datetime
import pandas as pd
from string import Template
from sqlalchemy import update
# 本地模块
from core.sql import sql_orm
from core.msg.msg_text import AddressBookManagement
from core.msg.msg_ding import DingMsg

class MsgEnergyOffline:
    def __init__(self):
        self.level_huandian = {
            0: [],
            2: ['一级督办对象', '二级督办对象'],
            3: ['一级督办对象', '二级督办对象'],
            6: ['一级督办对象', '二级督办对象'],
            12: ['一级督办对象', '二级督办对象', '三级督办对象'],
            24: ['一级督办对象', '二级督办对象', '三级督办对象', '四级督办对象']
        }
        self.level_chongdian = {
            0: [],
            12: ['一级督办对象', '二级督办对象'],
            20: ['一级督办对象', '二级督办对象', '三级督办对象'],
            24: ['一级督办对象', '二级督办对象'],
            48: ['一级督办对象', '二级督办对象', '三级督办对象']
        }
        self.msg_code = 'SMS_501925271'
        self.msg_template = "<能源设备离线督办>${station_name}，有${num}个${buss_type}设备已离线${broken_time}小时，请督办处理，无法及时维修，请在能源专业网管申请退网"
        self.now=datetime.datetime.now()
        self.address_book_management=AddressBookManagement()
    def _get_level(self,sql,pojo,row,valid_levels):
        send_flag=0
        res=sql.query(pojo).filter(pojo.site_code==row["点位编码"]).first()
        if res:
            current_level=res.send_level
            level = current_level
            candidates = [lvl for lvl in valid_levels if lvl > current_level]
            if candidates:
                next_level = min(candidates)  # 下一个级别（紧邻的）
                if row["离线时长"] >= next_level: # 有下一级别且到达下一级别，更新为新的，否则为旧的
                    level = next_level
                    res.send_level = level
                    res.send_time=self.now
                    send_flag = 1

        else:
            level=0
            temp = pojo()
            temp.site_code = row['点位编码']
            temp.send_level = level
            temp.send_time = self.now
            temp.business= row["business"]
            sql.merge(temp)
            sql.commit()

        return level,send_flag
    
    def _prepare_offline_device_data(self, df):
        df = df.fillna('')
        df = df.loc[df['点位编码'] != '']
        df = df.loc[df['在线状态'] == '离线']
        df = df.loc[df['设备状态'] != '退网下线']
        df['市'] = df['市'].str.replace('市', '')
        df['市'] = df['市'].str.replace('分公司', '')
        df['故障数'] = df.groupby('点位编码')['点位编码'].transform('count')
        df['最近一次离线时间'] = pd.to_datetime(df['最近一次离线时间'])
        df = df.sort_values(by='最近一次离线时间').drop_duplicates(subset='点位编码', keep='first').sort_index()
        df['离线时长'] = (self.now - df['最近一次离线时间']).dt.total_seconds() / 3600

        # 还原不离线的
        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.msg_energy_offline_level
            no_broken_list = set(df['点位编码'].unique())
            stmt = (
                update(pojo).
                    where(pojo.site_code.notin_(no_broken_list)).
                    values(send_level=0)
            )
            sql.execute(stmt)
            # 计算下一等级
            huan_mask = df['业务类型'] == '换电'
            chong_mask = df['业务类型'] == '低速充电'
            huandian_levels = sorted([k for k in self.level_huandian.keys() if k > 0])  # [2, 3, 6, 12, 24]
            chongdian_levels = sorted([k for k in self.level_chongdian.keys() if k > 0])  # [12, 24]
            if huan_mask.any():
                df.loc[huan_mask,  ['level', 'send_flag']] = df.loc[huan_mask].apply(
                    lambda row: self._get_level(sql,pojo,row, huandian_levels),
                    axis=1,
                    result_type='expand'
                ).values
            if chong_mask.any():
                df.loc[chong_mask,  ['level', 'send_flag']] = df.loc[chong_mask].apply(
                    lambda row: self._get_level(sql,pojo,row, chongdian_levels),
                    axis=1,
                    result_type='expand'
                ).values
        return df

    def _send_by_site(self, row):
        if row['点位编码']:
            data = {'station_name': f"{row['市']}{row['区县']}{row['点位名称']}",
                    'num': row['故障数'],
                    'buss_type':row['业务类型'],
                    'broken_time': round(row['离线时长'], 2)}

            # 业务类型到 level 字典的映射
            level_map = {
                "低速充电": self.level_chongdian,
                "换电": self.level_huandian
            }
            if row["业务类型"] in level_map:
                level_list = level_map[row["业务类型"]].get(row["level"], [])
                phone_list=[]
                for level in level_list:
                    address_book = self.address_book_management.get_address_book(
                        city=row["市"],
                        area=row["区县"],
                        businessCategory="能源",
                        specificBusiness=row["业务类型"],
                        level=level,
                        tasks=row["business"]
                    )
                    if not address_book.empty:
                        phone_list.append(address_book)
                if phone_list != []:
                    phone_df = pd.concat(phone_list)
                    phone_df = phone_df.drop_duplicates(subset=['phone'])
                    # 判断是否发送短信
                    send_sms = True
                    send_ding = True
                    if row['业务类型'] == '换电':
                        if row['level'] == 3:
                            send_sms = False
                        if row['level'] == 2:
                            send_ding = False
                    elif row['业务类型'] == '低速充电':
                        if row['level'] in (24, 48):
                            send_sms = False
                        if row['level']==12:
                            send_ding = False

                    # 发送短信
                    if send_sms:
                        self.address_book_management.send_msg(phone_df, data, self.msg_code)

                    # 发送钉钉
                    if send_ding:
                        order_text = Template(self.msg_template).substitute(data)
                        d = DingMsg()
                        d.text_at(d.ENERGY_MAINTAIN, order_text, phone_df['phone'].tolist(), [])

    def _down(self,bussiness='4'):
        """
        下载设备离线数据，status:"2"是离线
        :param Authorization:
        :param bussiness:'4'为低速充电，'2'为换电，需要设置devType="0328"，即只看换电柜，电池的离线不管
        :return:设备离线清单
        """
        Authorization = sql_orm().get_cookies("energy")["cookies_str"]
        cols_to_name = {
            '市': 'cityName',
            '区县': 'countyName',
            '点位名称': 'stationName',
            '点位编码': 'stationPubCode',
            '在线状态': 'onlineStatusName',
            '最近一次离线时间': 'outLineTime',
            '设备状态': 'statusName',
            '业务类型': 'businessTypeName'
        }
        headers = {
            'Content-Type': 'application/json;charset=UTF-8',
            'Authorization': Authorization
        }

        df_final = pd.DataFrame(columns=cols_to_name.keys())
        # 定义要遍历的 status 列表
        statuses = ["2"]
        if bussiness == '2':
            statuses.append("1")  # 如果是业务类型2，额外加上 status=1

        for status in statuses:
            data = {
                "devType": "",
                "accessPointId": "",
                "pageNum": 1,
                "pageSize": 100,
                "businessType": bussiness,
                "status": status,
                "deptIds": [],
                "onlineStatus": "0"
            }
            if bussiness == '2':
                data["devType"] = "0328"  # 只要换电柜

            df = pd.DataFrame(columns=cols_to_name.keys())
            page = 1

            # 获取总数
            res = requests.post(
                'https://energy-iot.chinatowercom.cn/api/device/device/page',
                headers=headers,
                data=json.dumps(data)
            )
            total = json.loads(res.text)['total']

            while page * 100 < total + 100:
                data['pageNum'] = page
                i = 0
                rows = []
                while i < 4:
                    try:
                        i += 1
                        res = requests.post(
                            'https://energy-iot.chinatowercom.cn/api/device/device/page',
                            headers=headers,
                            data=json.dumps(data)
                        )
                        text = res.json()
                        rows = text.get('rows', [])
                        break
                    except Exception as e:
                        print(f"请求失败（第{i}次重试），错误：{e}")
                        continue

                for row in rows:
                    temp = {}
                    for key, value in cols_to_name.items():
                        try:
                            temp[key] = row[value]
                        except KeyError:
                            temp[key] = ''
                    new_row = pd.Series(temp, name=len(df))
                    df.loc[len(df)] = new_row

                page += 1

            df_final = pd.concat([df_final, df], ignore_index=True)

        return df_final

    def run(self):
        df = pd.concat([self._down(bussiness='4'),self._down(bussiness='2')])
        df["business"]=df["业务类型"]+"离线提醒"
        df = self._prepare_offline_device_data(df)
        df=df.loc[df["send_flag"]==1]
        df.apply(self._send_by_site, axis=1)
def main():
    MsgEnergyOffline().run()

if __name__ == "__main__":
    main()