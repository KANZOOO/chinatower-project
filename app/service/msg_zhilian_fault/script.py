# zhilian_fault_msg.py
import requests
import pandas as pd
import datetime
from core.sql import sql_orm
from core.msg.msg_text import AddressBookManagement


class ZhiLianFaultMsg:
    def __init__(self):
        self.msg_to_code = 'SMS_485350644'
        self.now = datetime.datetime.now()
        self.level_str_dict = {
            1: "一级督办对象",
            2: "二级督办对象",
            3: "三级督办对象",
            4: "四级督办对象",
            5: "五级督办对象"
        }

    def _calculate_level(self, minutes: int, status: int) -> int:
        limit = 120 if status == 0 else 1440  # 待接单2小时，待回单24小时
        ratio = max(0, limit - minutes) / limit
        if ratio >= 0.75:
            return 1
        if ratio >= 0.50:
            return 2
        if ratio >= 0.25:
            return 3
        if ratio >= 0.15:
            return 4
        return 5

    def _down_fault_data(self):
        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.cookies
            cookie = sql.query(pojo.cookies).filter(pojo.id == 'energy').scalar()
        url = "https://zlzywg.chinatowercom.cn:8070/api/semp/alarmMix/fault/page"
        headers = {
            "Authorization": cookie,
            "Content-Type": "application/json"
        }
        data = []
        for status in (0, 1):  # 0=待接单, 1=待回单
            page = 1
            while True:
                body = {
                    "admProvinceCode": "450000",
                    "serveStatus": "1,3",
                    "pageNum": page,
                    "pageSize": 100,
                    "faultStatus": status
                }
                try:
                    res = requests.post(url, json=body, headers=headers, verify=False, timeout=10)
                    result = res.json()
                    rows = result["data"]["rows"]
                    total = result["data"]["total"]

                    if not rows:
                        break

                    for r in rows:
                        if r.get("maintainSecondName") != "铁塔维护":
                            continue
                        minutes = r["lastTime"]
                        row = {
                            "站址编码": r["stationCode"],
                            "故障状态编码": status,
                            "发生时间": r["createTime"],
                            "持续时间": minutes,
                            "督办等级": self._calculate_level(minutes, status),
                            "管理区域(市)": r["admCityName"],
                            "站址名称": r["stationName"],
                            "剩余时间": minutes,
                            "业务类型": "智联故障监控"
                        }
                        data.append(row)

                    page += 1
                    # 如果已获取所有数据则退出
                    if page * 100 >= total:
                        break

                except Exception as e:
                    print(f"请求失败: {e}")
                    break
        return pd.DataFrame(data)
    def _process_df(self, df):
        if df.empty:
            return df

        # 转换数据类型
        df['故障状态'] = df['故障状态编码'].map({0: '接单', 1: '回单'})
        df["发生时间"] = pd.to_datetime(df["发生时间"])
        df["持续时间"] = df["持续时间"].astype(int)
        df["管理区域(市)"] = df["管理区域(市)"].str.replace("市", "").str.replace("分公司", "")

        # 从数据库获取区域信息
        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.station
            site_codes = df['站址编码'].tolist()
            results = sql.query(pojo).filter(pojo.site_code.in_(site_codes)).all()
            area_map = {r.site_code: r.area for r in results}
            df['区域'] = df['站址编码'].map(area_map).fillna('')

        # 生成唯一ID: 站址编码_故障状态编码_业务类型_故障状态
        df['id'] = df['站址编码'] + '_' + df['故障状态编码'].astype(str) + '_' + df['业务类型'] + '_' + df['故障状态']

        # 去重：保留持续时间最长的记录
        df = df.sort_values('持续时间', ascending=False).drop_duplicates(['站址编码', '故障状态编码'], keep='first')
        df = df.reset_index(drop=True)

        return df

    def _update_send_status(self, df):
        if df.empty:
            return

        with sql_orm().session_scope() as (sql, Base):
            pojo = Base.classes.msg_zhilian_fault_level
            for index, row in df.iterrows():
                res = sql.query(pojo).filter(pojo.id == row.id).first()
                occur_time = row['发生时间']
                current_level = row['督办等级']

                flag = False
                send_level = 0

                if res is None:
                    # 无记录，添加 level=0 的记录
                    temp = pojo()
                    temp.id = row.id
                    temp.level = 0
                    temp.send_time = self.now
                    sql.merge(temp)
                else:
                    if res.send_time < occur_time:
                        # 有记录但时间 < occur_time，更新时间为现在，level 为 0
                        res.send_time = self.now
                        res.level = 0
                    elif res.send_time > occur_time:
                        # 有记录且时间 > occur_time，判断记录 level 是否小于当前 level
                        if res.level < current_level:
                            res.level += 1
                            send_level = res.level
                            res.send_time = self.now
                            sql.commit()
                            flag = True

                if flag:
                    if send_level in [3, 4, 5]:
                        continue
                    level_str = self.level_str_dict[send_level]
                    self._send_msg(row, level_str)

    def _send_msg(self, row, level):
        data = {
            'city': row['管理区域(市)'],
            'area': row['区域'].replace('区', '') if row['区域'] else '未知',
            'site_name': row['站址名称'],
            'type': row['业务类型'],
            'operate_type': row['故障状态'],
            'time': f"{row['剩余时间']}分钟"
        }

        manager_df = AddressBookManagement().get_address_book(
            city=row['管理区域(市)'],
            area=row['区域'],
            businessCategory="拓展",
            level=level,
            tasks="智联故障监控"
        )

        if len(manager_df) > 0:
            # 去重手机号
            manager_df = manager_df.drop_duplicates(subset=["phone"])

            # 发送短信
            AddressBookManagement().send_msg(manager_df, data, self.msg_to_code)

            # 打印发送信息（保持原有格式）
            for idx, mgr_row in manager_df.iterrows():
                sender_info = f"{mgr_row['city']} {mgr_row['position']} {mgr_row['level']} {mgr_row['name']} {mgr_row['phone']}"
                print(f"[发送人] {sender_info}  ->  {data}")

    def run(self):
        df = self._down_fault_data()
        df = self._process_df(df)
        self._update_send_status(df)
def main():
    ZhiLianFaultMsg().run()
if __name__ == "__main__":
    main()