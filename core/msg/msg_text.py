import json
import pandas as pd
import datetime
from alibabacloud_dysmsapi20170525.client import Client as Dysmsapi20170525Client
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_dysmsapi20170525 import models as dysmsapi_20170525_models
from alibabacloud_tea_util import models as util_models
import uuid
import re
from core.config import settings
from core.sql import sql_orm
class Sample():
    def __init__(self):
        pass

    @staticmethod
    def create_client(
            access_key_id: str,
            access_key_secret: str,
    ) -> Dysmsapi20170525Client:
        config = open_api_models.Config(
            # 必填，您的 AccessKey ID,
            access_key_id=access_key_id,
            # 必填，您的 AccessKey Secret,
            access_key_secret=access_key_secret
        )
        config.endpoint = f'dysmsapi.aliyuncs.com'
        return Dysmsapi20170525Client(config)

    @staticmethod
    def main(
            phone,json_text,code
    ) -> None:
        # 请确保代码运行环境设置了环境变量 ALIBABA_CLOUD_ACCESS_KEY_ID 和 ALIBABA_CLOUD_ACCESS_KEY_SECRET。
        # 工程代码泄露可能会导致 AccessKey 泄露，并威胁账号下所有资源的安全性。以下代码示例使用环境变量获取 AccessKey 的方式进行调用，仅供参考，建议使用更安全的 STS 方式，更多鉴权访问方式请参见：https://help.aliyun.com/document_detail/378659.html
        client = Sample.create_client(settings.alibaba_cloud_access_key_id,settings.alibaba_cloud_access_key_secret)
        send_sms_request = dysmsapi_20170525_models.SendSmsRequest(
            phone_numbers=phone,
            sign_name='润建股份',
            template_code=code,
            template_param=json_text
        )
        runtime = util_models.RuntimeOptions()
        try:
            # 复制代码运行请自行打印 API 的返回值
            ali_response=client.send_sms_with_options(send_sms_request, runtime)
            body = ali_response.body
            biz_id = body.biz_id
            message = body.message
            return (biz_id,message)

        except Exception as error:
            # 如有需要，请打印 error
            print('发送失败'+str(error)+str(phone))

class AddressBookManagement:
    def __init__(self):
        self.db = settings.db_url
    def get_address_book(self, city, businessCategory, level, tasks, area="全部", specificBusiness="全部") -> pd.DataFrame:
        df = sql_orm(self.db).get_data("address_book")
        df = df[df['level'] == level]
        # 1. businessCategory 过滤
        mask_bc = (df['businessCategory'] == "全部") | (df['businessCategory'] == businessCategory)
        df = df[mask_bc]

        # 2. specificBusiness 过滤（仅对 businessCategory == "能源"）
        energy_mask = df['businessCategory'] == "能源"
        keep_energy = (df['specificBusiness'] == "全部") | (df['specificBusiness'] == specificBusiness)
        final_mask = (~energy_mask) | (energy_mask & keep_energy)
        df = df[final_mask]

        # 3. city 过滤：排除 city != "区公司" 且 city != 参数 的记录
        city_keep = (df['city'] == "区公司") | (df['city'] == city)
        df = df[city_keep]

        # 4. area 过滤：排除 area != "无" 且 area != 参数 的记录
        df['area_short'] = df['area'].astype(str).str[:2]
        area_short = str(area)[:2]
        area_keep = (df['area'] == "全部") | (df['area_short'] == area_short)
        df = df[area_keep]
        df.drop(columns=['area_short'], inplace=True)
        # 部分特殊要求
        if businessCategory != "一体":
            df = df[(df['name'] != "兰天桢")]
        if "工单" in tasks:
            df = df[(df['name'] != "谌亮书")]
        now_time = datetime.datetime.now().time()
        if now_time >= datetime.time(21, 0) or now_time < datetime.time(7, 0):
            df = df[df["level"] != "五级督办对象"]
            df = df[df["level"] != "四级督办对象"]
        # 优先保留指定业务
        df = df.sort_values(
            by='businessCategory',
            key=lambda x: x != businessCategory
        )
        df = df.drop_duplicates(subset=['phone'], keep='first')
        df["tasks"]=tasks
        df.drop(columns=['id'], inplace=True)

        return df
    def send_msg(self, df, data, code):
        send_time=datetime.datetime.now()
        df = df.drop_duplicates(subset=['phone']).reset_index(drop=True)
        for index,row in df.iterrows():
            send_id = str(uuid.uuid4())
            send_response = "发送失败"
            phone=row["phone"]
            if phone=="":
                send_response="未配置人员"
            else:
                cleaned = data
                for k, v in data.items():
                    if isinstance(v, str):
                        cleaned_v = v
                        for char in ['(', '（', '/', '#', '_', '－']:
                            cleaned_v = cleaned_v.replace(char, '')
                        cleaned[k] = cleaned_v

                send_id_tmep, send_response_temp = Sample().main(phone=re.sub(r'\D', '', phone), json_text=json.dumps(cleaned), code=code)
                if send_id_tmep:
                    send_id=send_id_tmep
                if send_response_temp:
                    send_response=send_response_temp
            df.loc[index,"send_id"]=send_id
            df.loc[index,"send_response"]=send_response
            df.loc[index,"send_time"] = send_time
            df.fillna("", inplace=True)
        sql_orm(self.db).add_data(df,"msg_log")

