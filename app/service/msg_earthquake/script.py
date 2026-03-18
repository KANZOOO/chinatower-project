import datetime
import pandas as pd
import asyncio
from playwright.async_api import async_playwright
from sqlalchemy import and_,text
from core.config import settings
from core.sql import sql_orm
from core.msg.msg_ding import DingMsg
from core.msg.msg_text import AddressBookManagement
import os
class EarthQuake():
    # 辅助函数
    def __init__(self):
        self.site_list= {
'GX/A0001':'广西南宁兴宁三塘一般站',
'GX/A0002':'广西南宁青秀刘圩一般站',
'GX/A0003':'广西南宁西乡塘石埠一般站',
'GX/A0004':'广西南宁良庆大塘一般站',
'GX/A0005':'广西南宁隆安城厢一般站',
'GX/A0006':'广西南宁马山古零一般站',
'GX/A0007':'广西南宁上林白圩一般站',
'GX/A0008':'广西南宁横县州镇一般站',
'GX/B0001':'广西柳州鱼峰大龙潭一般站',
'GX/B0002':'广西柳州柳江拉堡一般站',
'GX/B0003':'广西柳州融水融水一般站',
'GX/C0001':'广西桂林全州全州一般站',
'GX/C0002':'广西桂林兴安兴安一般站',
'GX/C0003':'广西桂林龙胜龙胜一般站',
'GX/D0001':'广西梧州长洲大塘一般站',
'GX/D0002':'广西梧州蒙山蒙山一般站',
'GX/D0003':'广西梧州岑溪岑城一般站',
'GX/E0001':'广西北海银海南万一般站',
'GX/E0002':'广西北海合浦常乐一般站',
'GX/E0003':'广西北海合浦乌家一般站',
'GX/P0001':'广西防城港港口渔洲坪一般站',
'GX/P0002':'广西防城港港口光坡一般站',
'GX/P0003':'广西防城港东兴东兴一般站',
'GX/N0001':'广西钦州钦南尖山一般站',
'GX/N0002':'广西钦州钦南大番坡一般站',
'GX/N0003':'广西钦州灵山太平一般站',
'GX/N0004':'广西钦州灵山伯劳一般站',
'GX/N0005':'广西钦州浦北小江一般站',
'GX/N0006':'广西钦州浦北寨圩一般站',
'GX/R0001':'广西贵港平南平南一般站',
'GX/R0002':'广西贵港桂平西山一般站',
'GX/K0001':'广西玉林陆川温泉一般站',
'GX/K0002':'广西玉林兴业石南一般站',
'GX/K0003':'广西玉林博白博白一般站',
'GX/K0004':'广西玉林博白东平一般站',
'GX/K0005':'广西玉林北流北流一般站',
'GX/L0001':'广西百色右江永乐一般站',
'GX/L0002':'广西百色田东平马一般站',
'GX/L0003':'广西百色德保城关一般站',
'GX/L0004':'广西百色那坡城厢一般站',
'GX/L0005':'广西百色凌云泗城一般站',
'GX/L0006':'广西百色乐业同乐一般站',
'GX/L0007':'广西百色西林八达一般站',
'GX/L0008':'广西百色隆林新州一般站',
'GX/L0009':'广西百色靖西新靖一般站',
'GX/J0001':'广西贺州八步信都一般站',
'GX/J0002':'广西贺州昭平黄姚一般站',
'GX/M0001':'广西河池宜州庆远一般站',
'GX/M0002':'广西河池南丹城关一般站',
'GX/M0003':'广西河池巴马巴马一般站',
'GX/M0004':'广西河池都安澄江一般站',
'GX/G0001':'广西来宾兴宾小平阳一般站',
'GX/F0001':'广西崇左江州驮卢一般站',
'GX/F0002':'广西崇左龙州龙州一般站',
'GX/F0003':'广西崇左天等天等一般站'
}
        self.earthquake_url='http://8.134.209.192:5173/'
        self.code_to_city={
            'A':'南宁',
            'B':'柳州',
            'C':'桂林',
            'D':'梧州',
            'E':'北海',
            'P':'防城港',
            'N':'钦州',
            'R':'贵港',
            'K':'玉林',
            'L':'百色',
            'J':'贺州',
            'M':'河池',
            'G':'来宾',
            'F':'崇左',

            }
    def init_temp(self,res,pojo):
        temp = pojo()
        # 默认赋值为res储存内容
        temp.siteid = res.siteid
        temp.status = res.status
        temp.time = res.time
        temp.brokentime = res.brokentime
        temp.sendornot = res.sendornot
        temp.lastsendtime = res.lastsendtime
        temp.order_send = res.order_send
        temp.redtimes = res.redtimes
        temp.redtimesday=res.redtimesday
        return temp
    def time_to_hms(self,time):
        if time//3600!=0:
            hour=time//3600
            time=time-hour*3600
        else :hour=0
        if time//60!=0:
            min=time//60
            time=time-min*60
        else:min=0

        seconds=time
        outputime=[hour,min,seconds]
        return outputime
    def send_msg(self,city_list,sum_text):
        d=DingMsg()
        for city in city_list:
            city = self.code_to_city[city]
            d = DingMsg()
            phone_list=[]
            for level in ["一级督办对象","二级督办对象"]:
                address_book = AddressBookManagement.get_address_book(
                    city=city,
                    businessCategory="拓展",
                    level=level,
                    tasks="地震设备故障提醒"
                )
                if not address_book.empty:
                    phone_list.append(address_book)
        if phone_list != []:
            phone_df = pd.concat(phone_list)
            phone_df = phone_df.drop_duplicates(subset=['phone'])
            d.text_at(d.DZ, sum_text, phone_df['phone'].tolist(), [])
        else:
            d.text_at(d.DZ, sum_text, [], [])

    async def send_now(self):
        sum_text = '----地震局监控平台故障提醒----\n当前故障：\n\n本次恢复：'
        flag_blue = 0
        flag_red = 0
        flag_onehour = 0
        flag_twohour = 0
        flag_maintence = 0
        city_list = []
        order_text = ""  # 👈 初始化，避免未定义错误

        # 确保 user_data_dir 存在
        user_data_dir = settings.resolve_path("playwright_user_data")
        os.makedirs(user_data_dir, exist_ok=True)

        async with async_playwright() as p:
            try:
                context = await p.chromium.launch_persistent_context(
                    user_data_dir=user_data_dir,
                    headless=True,
                    args=[
                        "--disable-gpu",
                        "--no-sandbox"
                    ],
                    # executable_path=...  # 如需指定 Chrome 路径，请确保是标准安装版
                )

                page = context.pages[0] if context.pages else await context.new_page()

                await page.goto(self.earthquake_url, timeout=30000)
                await page.wait_for_selector('.ant-btn.btnRed, .ant-btn.btnBlue', timeout=20000)

                elements_red = await page.query_selector_all('.ant-btn.btnRed')
                elements_blue = await page.query_selector_all('.ant-btn.btnBlue')

                df = pd.DataFrame(columns=['站址编码', '故障站址', '发生时间', '恢复时间', '故障持续时间'])

                # === 处理蓝色按钮（恢复）===
                with sql_orm().session_scope() as temp:
                    session, Base = temp
                    pojo = Base.classes.msg_earthquake

                    for element in elements_blue:
                        span = await element.query_selector('span')
                        if not span:
                            continue
                        raw_text = await span.inner_text()  # ✅ 正确 await
                        station_code = raw_text.strip()

                        if station_code in self.site_list:
                            res = session.query(pojo).filter(pojo.siteid == station_code).first()
                            if not res:
                                continue
                            temp_obj = self.init_temp(res, pojo)
                            temp_obj.status = '蓝'
                            temp_obj.brokentime = 0
                            temp_obj.sendornot = 0

                            if res.status == '红':
                                hms = self.time_to_hms(
                                    (datetime.datetime.now() - datetime.datetime.strptime(res.time,
                                                                                          "%Y/%m/%d %H:%M:%S")).seconds
                                )
                                temp_obj.time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                                df.loc[len(df)] = [
                                    station_code,
                                    self.site_list[station_code],
                                    res.time,
                                    temp_obj.time,
                                    f"{hms[0]}小时{hms[1]}分{hms[2]}秒"
                                ]
                                if res.sendornot == 1:
                                    flag_blue = 1
                                    index = sum_text.find('\n本次恢复：') + len('\n本次恢复：')
                                    insert_text = (
                                        f"\n{station_code}-{self.site_list[station_code]}，故障历时："
                                        f"{hms[0]}小时{hms[1]}分{hms[2]}秒 （{res.time[res.time.index(' ') + 1:]}~{temp_obj.time[temp_obj.time.index(' ') + 1:]}）"
                                    )
                                    sum_text = sum_text[:index] + insert_text + sum_text[index:]
                                    city_list.append(station_code[3])

                            session.merge(temp_obj)

                # === 处理红色按钮（故障）===
                with sql_orm().session_scope() as temp:
                    session, Base = temp
                    pojo = Base.classes.msg_earthquake

                    for element in elements_red:
                        span = await element.query_selector('span')
                        if not span:
                            continue
                        raw_text = await span.inner_text()  # ✅ 正确 await
                        station_code = raw_text.strip()

                        if station_code in self.site_list:
                            res = session.query(pojo).filter(pojo.siteid == station_code).first()
                            if not res:
                                continue
                            temp_obj = self.init_temp(res, pojo)
                            temp_obj.status = '红'

                            if res.status == '蓝':
                                temp_obj.time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                                temp_obj.redtimes += 1
                                temp_obj.redtimesday += 1
                            elif res.status == '红':
                                brokentime_sec = (datetime.datetime.now() - datetime.datetime.strptime(res.time,
                                                                                                       "%Y/%m/%d %H:%M:%S")).seconds
                                temp_obj.brokentime = brokentime_sec
                                hms = self.time_to_hms(brokentime_sec)

                                if brokentime_sec >= 360 and res.sendornot == 0:
                                    flag_red = 1
                                    temp_obj.sendornot = 1
                                    temp_obj.lastsendtime = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                                    index = sum_text.find('当前故障：') + len('当前故障：')
                                    insert_text = (
                                        f"\n(新增){station_code}-{self.site_list[station_code]}（故障发生时间："
                                        f"{res.time[res.time.index(' ') + 1:]} 已持续：{hms[0]}小时{hms[1]}分{hms[2]}秒）"
                                    )
                                    sum_text = sum_text[:index] + insert_text + sum_text[index:]
                                    city_list.append(station_code[3])

                                elif res.sendornot == 1:
                                    if (datetime.datetime.now() - datetime.datetime.strptime(res.lastsendtime,
                                                                                             "%Y/%m/%d %H:%M:%S")).seconds >= 7200:
                                        flag_twohour = 1
                                        temp_obj.lastsendtime = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")

                                    index = sum_text.find('\n\n本次恢复')
                                    insert_text = (
                                        f"\n{station_code}-{self.site_list[station_code]}（故障发生时间："
                                        f"{res.time[res.time.index(' ') + 1:]} 已持续：{hms[0]}小时{hms[1]}分{hms[2]}秒）"
                                    )
                                    sum_text = sum_text[:index] + insert_text + sum_text[index:]
                                    flag_maintence = 1
                                    city_list.append(station_code[3])

                                    if (datetime.datetime.now() - datetime.datetime.strptime(res.lastsendtime,
                                                                                             "%Y/%m/%d %H:%M:%S")).seconds >= 3600 and res.order_send == 0:
                                        temp_obj.order_send = 1
                                        order_text = '以下站点故障时长超过1小时，请派单：' + insert_text
                                        flag_onehour = 1

                            session.merge(temp_obj)

                # === 消息发送逻辑 ===
                if flag_blue == 1 or flag_red == 1 or flag_twohour == 1:
                    if flag_blue == 0:
                        idx = sum_text.find('\n本次恢复：') + len('\n本次恢复：')
                        sum_text = sum_text[:idx] + '无' + sum_text[idx:]
                    if flag_red == 0 and flag_maintence == 0:
                        idx = sum_text.find('\n当前故障：') + len('\n当前故障：')
                        sum_text = sum_text[:idx] + '无' + sum_text[idx:]
                    self.send_msg(city_list, sum_text)

                if flag_onehour == 1:
                    self.send_msg([], order_text)

                # === 记录日志 ===
                with sql_orm().session_scope() as temp:
                    session, Base = temp
                    pojo_log = Base.classes.msg_earthquake_log
                    df_renamed = df.rename(columns={
                        '站址编码': 'sitecode',
                        '故障站址': 'sitename',
                        '发生时间': 'happen',
                        '恢复时间': 'recover',
                        '故障持续时间': 'duration'
                    })
                    for _, row in df_renamed.iterrows():
                        log_entry = pojo_log(**row.to_dict())
                        session.add(log_entry)

                return 0

            except Exception as e:
                raise e
            finally:
                if 'context' in locals():
                    await context.close()
    def send_day(self):
        # 发送每日盘点信息
        # 时钟状态推进，避免9点一小时内重复发
        with sql_orm().session_scope() as temp:
            session, Base = temp
            pojo = Base.classes.msg_earthquake_counttime
            flag=0
            res = session.query(pojo).first()
            if (datetime.datetime.now() - datetime.datetime.strptime(res.counttime,"%Y/%m/%d %H:%M:%S")).seconds >= 3600:
                if datetime.datetime.now().strftime('%H') == '09':
                    flag= 1
                res.counttime=datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        serial = 1
        city_list=[]
        df = pd.DataFrame(columns=['类型','时间段','故障次数','站址编码'])
        if flag==1:
            insert_text = '----地震局监控平台24小时报----\n24小时内出现大于或等于5次闪断的站点，请及时处理\n'
            with sql_orm().session_scope() as temp:
                session, Base = temp
                pojo = Base.classes.msg_earthquake
                for station_code in self.site_list.keys():
                    res = session.query(pojo).filter(pojo.siteid == station_code).first()
                    if res.redtimesday >= 5:
                        insert_text += str(serial) + '、' + res.siteid + '-' + self.site_list[res.siteid] + ',故障次数：' + str(res.redtimesday) + '次\n'
                        serial += 1
                        city_list.append(station_code[3])
                        nowtime = datetime.datetime.now()
                        df.loc[len(df)] = ['24小时报', nowtime.strftime("%Y-%m-%d-%H:%M") + ' ' + (nowtime - datetime.timedelta(hours=4)).strftime("%Y-%m-%d-%H:%M"),res.redtimesday,res.siteid]
                    temp = self.init_temp(res,pojo)
                    temp.redtimesday = 0
                    session.merge(temp)

            if '、' in insert_text:
                self.send_msg(city_list,insert_text)
            else:
                insert_text='----地震局监控平台24小时报----\n无24小时内出现大于或等于5次闪断的站点'
                self.send_msg(city_list, insert_text)
    def send_circle(self):
        # 轮询，一天内每10次告警则通知
        with sql_orm().session_scope() as temp:
            session, Base = temp
            pojo=Base.classes.msg_earthquake_ten_times
            end=datetime.datetime.now()
            begin=end-datetime.timedelta(days=1)
            end=end.strftime('%Y-%m-%d %H:%M:%S')
            begin = begin.strftime('%Y-%m-%d %H:%M:%S')
            for station_code in self.site_list.keys():
                res = session.execute(text(f"select count(*) as brokentimes from msg_earthquake_log where sitecode='{station_code}' and happen between '{begin}' and '{end}'")).first()
                site=session.query(pojo).filter(pojo.sitecode==station_code).first()
                if datetime.datetime.now().hour==8 and datetime.datetime.now().minute<10:
                    site.brokentimes=0
                else:
                    if res.brokentimes-site.brokentimes >= 10:
                        site.brokentimes=res.brokentimes
                        order_text=f'{station_code}-{self.site_list[station_code]}\n该站址24小时内故障次数超过{res.brokentimes}次，请派单'
                        self.send_msg([], order_text)

def main():
    earth=EarthQuake()
    asyncio.run(earth.send_now())
    earth.send_day()
    earth.send_circle()
if __name__ == '__main__':
    main()