#from core.utils.retry_wrapper import requests_post
from bs4 import BeautifulSoup
import os
import pandas as pd
from datetime import datetime,timedelta
from core.sql import sql_orm
from xlsx2csv import Xlsx2csv
import time
import re
import requests
import numpy as np
from functools import wraps
from sqlalchemy import text
#from spider.script.： import foura_data
import calendar
from core.config import settings
import shutil

"""
通用函数模块，包含项目中常用的工具函数，如重试装饰器、请求函数、文件操作函数等。
"""
def retry(max_attempts=1, delay=2):
    """
    重试装饰器，用于对指定函数进行重试操作。

    :param max_attempts: 最大重试次数，默认为15次。
    :param delay: 每次重试之间的等待时间（秒），默认为2秒。
    :return: 装饰器函数。
    """

    def decorator_retry(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            nonlocal max_attempts  # 声明为非局部变量以便在内部函数中修改
            attempts = 0
            while attempts < max_attempts:
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    attempts += 1
                    if attempts == max_attempts:
                        raise
                    time.sleep(delay)  # 等待一段时间后重试

        return wrapper

    return decorator_retry
@retry()
def requests_post(url, headers={}, data={}, cookies={}, timeout=600):
    """
    # 用途：提交表单、下载文件
    发送POST请求，带有重试机制。
    :param url: 请求的URL地址。
    :param headers: 请求头，默认为空字典。
    :param data: 请求数据，默认为空字典。
    :param cookies: 请求的cookies，默认为空字典。
    :param timeout: 请求超时时间（秒），默认为300秒。
    :return: 请求响应对象。
    """
    return requests.post(url, headers=headers, data=data, cookies=cookies, timeout=timeout)
@retry()
def requests_get(url, headers={}, params={}, cookies={}):
    """
    发送GET请求，带有重试机制。

    :param url: 请求的URL地址。
    :param headers: 请求头，默认为空字典。
    :param params: 请求参数，默认为空字典。
    :param cookies: 请求的cookies，默认为空字典。
    :return: 请求响应对象。
    """
    return requests.get(url, headers=headers, params=params, cookies=cookies)
def get_foura_cookie():
    """
    从数据库中获取指定ID的foura cookie信息。

    :param ID: 数据库中cookie记录的ID，默认为1（对应foura1）。
    :return: 解析后的cookie字典；若ID不存在/格式错误，返回空字典。
    """
    db = sql_orm()
    cookie_result = db.get_cookies(f"foura")
    return cookie_result["cookies"]
def clear_folder(folder_temp):
    """
    清空指定文件夹下的所有文件。

    :param folder_temp: 要清空的文件夹路径。
    """
    for file in os.listdir(folder_temp):
        file = os.path.join(folder_temp, file)
        os.remove(file)
def xlsx_to_csv(folder):
    """
    将指定文件夹下的所有.xlsx文件转换为.csv文件。

    :param folder: 包含.xlsx文件的文件夹路径。
    """
    for file in os.listdir(folder):
        path = os.path.join(folder, file)
        if file.endswith('.xlsx'):
            csv_path = path.replace(".xlsx", ".csv")  # 对应的 .csv 文件路径
            if not os.path.exists(csv_path):
                Xlsx2csv(path, outputencoding="utf-8").convert(csv_path)
def concat_df(folder, output_path, gen_csv=False):
    """
    合并指定文件夹下的所有.csv和.xls文件，并保存为一个新的Excel文件。

    :param folder: 包含.csv和.xls文件的文件夹路径。
    :param output_path: 合并后文件的输出路径。
    :param gen_csv: 是否生成.csv文件，默认为False。
    :return: 合并后的DataFrame和输出路径。
    """
    xlsx_to_csv(folder)
    df_list = []
    for file in os.listdir(folder):
        path = os.path.join(folder, file)
        if '.csv' in file:
            temp = pd.read_csv(path, dtype=str)
            df_list.append(temp)
        elif file.endswith(('.xls')):
            temp = pd.read_excel(path, dtype=str, engine='xlrd')
            df_list.append(temp)

    merge = pd.concat(df_list)
    output_path_str = str(output_path)
    merge.to_excel(output_path_str, index=False)
    if gen_csv:
        csv_output_path = output_path_str.replace('.xlsx', '.csv')
        merge.to_csv(csv_output_path, index=False, encoding='utf-8-sig')
    return merge, output_path
def log_downtime(fuc_name):
    """
    记录指定操作的下载时间到数据库。

    :param fuc_name: 操作类型名称。
    """
    with sql_orm().session_scope() as temp:
        session, Base = temp
        pojo = Base.classes.update_downhour_log
        res = session.query(pojo).filter(pojo.type == fuc_name).first()
        res.time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
def down_file(url, data, path, conten_len_error=3000, xlsx_juge=False):
    """
    下载文件，支持重试机制和文件验证。

    :param url: 文件下载的URL地址。
    :param data: 请求数据。
    :param path: 文件保存路径。
    :param conten_len_error: 内容长度验证阈值，默认为3000。
    :param xlsx_juge: 是否验证文件类型为xlsx或xls，默认为False。
    """
    retry = 3
    while retry >= 0:
        try:
            headers = {
                'Host': 'omms.chinatowercom.cn:9000',
                'Origin': 'http://omms.chinatowercom.cn:9000',
                'Referer': url,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
            }
            cookies = get_foura_cookie()
            res = requests_post(url, headers=headers, cookies=cookies)
            html = BeautifulSoup(res.text, 'html.parser')
            javax = html.find('input', id='javax.faces.ViewState')['value']
            for key, into_data in data.items():
                into_data['javax.faces.ViewState'] = javax
                res = requests_post(url, headers=headers, data=into_data, cookies=cookies)

                if 'FINAL' in key:
                    if len(res.content) < conten_len_error:
                        raise ValueError("内容小于给定大小")
                    if xlsx_juge:
                        xlsx_signature = b'\x50\x4B\x03\x04'  # xlsx 文件的签名
                        xls_signature = b'\x09\x08\x04\x00\x10\x00\x00\x00'  # xls 文件的签名
                        if not res.content.startswith(xlsx_signature) and not res.content.startswith(xls_signature):
                            raise ValueError("内容不是 xlsx 或 xls")
                    with open(path, "wb") as codes:
                        codes.write(res.content)
            return
        except ValueError as e:
            if "内容小于给定大小" in str(e) and retry:
                retry -= 1
                continue
            raise
        except Exception:
            raise
def down_file_no_save(url, data):
    """
    下载文件，支持重试机制和文件验证。

    :param url: 文件下载的URL地址。
    :param data: 请求数据。
    :param path: 文件保存路径。
    :param conten_len_error: 内容长度验证阈值，默认为3000。
    :param xlsx_juge: 是否验证文件类型为xlsx或xls，默认为False。
    """
    try:
        headers = {
            'Host': 'omms.chinatowercom.cn:9000',
            'Origin': 'http://omms.chinatowercom.cn:9000',
            'Referer': url,
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
        }
        cookies = get_foura_cookie()
        res = requests_post(url, headers=headers, cookies=cookies)
        html = BeautifulSoup(res.text, 'html.parser')
        javax = html.find('input', id='javax.faces.ViewState')['value']
        for key, into_data in data.items():
            into_data['javax.faces.ViewState'] = javax
            res = requests_post(url, headers=headers, data=into_data, cookies=cookies)
        soup = BeautifulSoup(res.text, 'html.parser')
        tbody = soup.find('tbody', {'id': 'listForm:list:tb'})
        headers = []
        thead = soup.find('thead')
        if thead:
            header_cells = thead.find_all('th')
            for th in header_cells:
                text = th.get_text(strip=True)
                headers.append(text)
        rows = []
        for tr in tbody.find_all('tr', class_=['rich-table-row']):
            try:
                cells = tr.find_all('td')
                row = []
                for td in cells:
                    text = td.get_text(strip=True)
                    row.append(text)
                rows.append(row)
            except:
                pass
        df = pd.DataFrame(rows, columns=headers if len(headers) == len(rows[0]) else None)
        return df
    except Exception as e:
        print(e)

"""
爬取站址信息，路径：资源管理-站址管理-取消FSU工程状态：交维，改为全选-导出
"""
class Station():
    def __init__(self):
        self.data = foura_data.station
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/siteMge/listSite.xhtml'
        self.down_name = '站址信息'
        self.down_name_en = 'station'
        self.down_suffix = '.xlsx'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        down_list = ['0099977', '0099978', '0099979', '0099980', '0099981', '0099982', '0099983', '0099984', '0099985',
                     '0099986', '0099987', '0099988', '0099989', '0099990']
        for city in down_list:
            for key in ['1', '2', '4', '5']:
                self.data[key]['queryForm:unitHidden'] = city
            path = os.path.join(self.folder_temp, f"{city}{self.down_suffix}")
            if city in ['0099982', '0099983']:
                down_file(self.URL, self.data, path, conten_len_error=3000000)
            else:
                down_file(self.URL, self.data, path)

    def read_file(self):
        xlsx_to_csv(self.folder_temp)
        df_list = []
        for file in os.listdir(self.folder_temp):
            path = os.path.join(self.folder_temp, file)
            if '.csv' in file:
                temp = pd.read_csv(path, dtype=str)
                df_list.append(temp)
        return df_list

    def save_as_sql_xlsx_csv(self, df_list):
        use_cols = {'所属省': 'province', '所属市': 'city', '区县（行政区划）': 'area', '所属区': 'area2',
                    "乡镇（街道）": "village",
                    '行政村（居委会）': 'countryside', '名称': 'site_name',
                    '站址编码': 'site_code', '运维ID': 'site_maitan_code', '站址状态': 'site_status',
                    'FSU工程状态': 'fsu_status',
                    "站址保障等级": 'level', "所属运营商": "belong"}

        def standardized_station(df):
            df = df.where(pd.notnull(df), None)  # nan改为空
            df = df.dropna(subset=['站址编码'])  # 去除主键空值
            df = df.rename(columns=use_cols)
            df = df[list(use_cols.values())]
            df['city'] = df['city'].str.replace('市', '')
            df['city'] = df['city'].str.replace('分公司', '')
            df = df.drop_duplicates(subset=['site_maitan_code'], keep='last')

            return df

        orm = sql_orm()

        # 更新area信息
        with orm.session_scope() as (session, Base):
            # 更新station信息
            session.execute(text('TRUNCATE station'))
            for df in df_list:
                df = standardized_station(df)
                orm.save_data(df, 'station')
            session.execute(text('TRUNCATE station_with_area'))
            session.commit()
            try:
                session.execute(text('insert into station_with_area select site_maitan_code,area from tower.station'))
            except Exception as e:
                raise
        # 保存为文件xlsx/csv
        merge = pd.concat(df_list)
        merge.to_excel(self.output_path, index=False)
        csv_path = str(self.output_path).replace('.xlsx', '.csv')
        merge.to_csv(csv_path, index=False, encoding='utf-8-sig')
    def main(self):
        self.down()
        if len(os.listdir(self.folder_temp)) >= 14:
            df_list = self.read_file()
            self.save_as_sql_xlsx_csv(df_list)
            log_downtime(self.down_name_en)
        else:
            raise FileNotFoundError("文件下载不全，当前仅下载了 {} 个文件".format(len(os.listdir(self.folder_temp))))

"""
爬取两翼站址信息，路径：资源管理-站址管理-两翼-导出
"""
class StationLiangYi():
    def __init__(self):
        self.data = foura_data.station_liangyi
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/siteMge/listSite.xhtml'
        self.down_name = '站址信息'
        self.down_name_en = 'station_liangyi'
        self.down_suffix = '.xls'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        down_list = ['0099977', '0099978', '0099979', '0099980', '0099981', '0099982', '0099983', '0099984', '0099985',
                     '0099986', '0099987', '0099988', '0099989', '0099990']
        for city in down_list:
            for key in ['1', '2']:
                self.data[key]['queryZlForm:unitZLHidden'] = city
            path = os.path.join(self.folder_temp, f"{city}{self.down_suffix}")
            down_file(self.URL, self.data, path)

    def read_file(self):
        df_list = []
        for file in os.listdir(self.folder_temp):
            path = os.path.join(self.folder_temp, file)
            if '.xls' in file:
                temp = pd.read_excel(path, dtype=str)
                df_list.append(temp)
        return df_list

    def main(self):
        self.down()
        if len(os.listdir(self.folder_temp)) >= 14:
            df_list = self.read_file()
            merge = pd.concat(df_list)
            merge.to_excel(self.output_path, index=False)
            log_downtime(self.down_name_en)
        else:
            raise FileNotFoundError("文件下载不全，当前仅下载了 {} 个文件".format(len(os.listdir(self.folder_temp))))

"""
爬取站址信息，路径：运行监控-基站直流负载总电流-查询-导出
"""
class StationDC():
    def __init__(self):
        self.data = foura_data.station_DC
        self.now = datetime.now()
        if self.now.day == 1:
            self.now = self.now - timedelta(days=1)
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/siteMge/staPerDataList.xhtml'
        self.down_name = '基站负载电流'
        self.down_name_en = 'DC'
        self.down_suffix = '.xls'
        # 改动7：统一使用settings.resolve_path解析路径
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_folder = settings.resolve_path(f"spider/down/{self.down_name_en}")
        self.output_path = settings.resolve_path(
            f"{self.output_folder}/{self.now.strftime('%Y%m')}{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        for index, city in enumerate(['0099977,0099978,0099979', '0099980,0099981,0099982',
                                      '0099983,0099984,0099985', '0099986,0099987,0099988', '0099989,0099990']):
            for key in ['1', '2']:
                self.data[key]['queryForm:unitHidden'] = city
                self.data[key]['queryForm:j_id19'] = self.now.year
                self.data[key]['queryForm:queryMonthid'] = self.now.strftime('%m').lstrip('0')
                self.data[key]['queryForm:queryMonthid_hiddenValue'] = self.now.strftime('%m').lstrip('0')
            # 改动8：用os.path.join拼接路径
            path = os.path.join(self.folder_temp, f"{index}{self.down_suffix}")
            down_file(self.URL, self.data, path)

    def sql_process(self, df):
        df = df.drop(columns=["备注"]).fillna("")
        orm = sql_orm()
        orm.save_data_merge(df, '基站负载电流')

    def main(self):
        self.down()
        merge, _ = concat_df(self.folder_temp, self.output_path)
        self.sql_process(merge)
        log_downtime(self.down_name_en)

    def temp(self):
        folder = self.output_folder
        for path in os.listdir(folder):
            if '.xlsx' in path:
                xlsx_path = os.path.join(folder, path)
                df = pd.read_excel(xlsx_path, dtype=str)
                self.sql_process(df)

"""
爬取fsu查询，路径：运行监控-fsu查询-导出
"""
class FsuChaXun():
    def __init__(self):
        self.data = foura_data.fsu_chaxun
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/pwMge/fsuMge/listQuertFsu.xhtml'
        self.down_name = 'fsu清单'
        self.down_name_en = 'fsu_chaxun_all'
        self.down_suffix = '.xlsx'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        for index, city in enumerate(['0099977,0099978,0099979,0099980,0099981,0099982',
                                      '0099983,0099984,0099985,0099986,0099987,0099988,0099989,0099990,2710377449']):
            for key in ['1', '2']:
                self.data[key]['queryForm:unitHidden'] = city
            path = os.path.join(self.folder_temp, f"{index}{self.down_suffix}")
            down_file(self.URL, self.data, path)

    def main(self):
        self.down()
        concat_df(self.folder_temp, self.output_path, gen_csv=True)
        log_downtime(self.down_name_en)

"""
爬取fsu查询，路径：工单管理-日常修理-隐患库-[归档起止时间:清除，地市:全选]-导出
"""
class YinHuanOrder():
    def __init__(self):
        self.data = foura_data.yinhuan_order
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/hiddenFixMge/monitorList.xhtml'
        self.down_name = '隐患工单'
        self.down_name_en = 'yinhuan_order'
        self.down_suffix = '.xls'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self, guidang):
        clear_folder(self.folder_temp)
        # for month in range(1, self.now.month + 1):
        #     last_day = calendar.monthrange(self.now.year, month)[1]
        #     start = datetime(self.now.year, month, 1)
        #     end = datetime(self.now.year, month, last_day, 23, 59, 59)
        start_year = self.now.year - 1
        end_year = self.now.year
        end_month = self.now.month

        for year in range(start_year, end_year + 1):
            max_month = 12 if year < self.now.year else end_month
            for month in range(1, max_month + 1):  # 月份循环
                last_day = calendar.monthrange(year, month)[1]
                start = datetime(year, month, 1)
                end = datetime(year, month, last_day, 23, 59, 59)

                for key in ['1', '2', '3']:
                    self.data[key]['queryForm:queryType'] = guidang
                    self.data[key]['queryForm:findtimeStartInputDate'] = start.strftime('%Y-%m-%d %H:%M')
                    self.data[key]['queryForm:findtimeEndInputDate'] = end.strftime('%Y-%m-%d %H:%M')
                    self.data[key]['queryForm:findtimeStartInputCurrentDate'] = start.strftime('%m/%Y')
                    self.data[key]['queryForm:findtimeEndInputCurrentDate'] = end.strftime('%m/%Y')

                    self.data[key]['queryForm:finishtimeStartInputCurrentDate'] = start.strftime('%m/%Y')
                    self.data[key]['queryForm:hiddenAuditDateStartInputCurrentDate'] = start.strftime('%m/%Y')
                    self.data[key]['queryForm:hiddenRecordDateStartInputCurrentDate'] = start.strftime('%m/%Y')
                    self.data[key]['queryForm:finishtimeEndInputCurrentDate'] = end.strftime('%m/%Y')
                    self.data[key]['queryForm:hiddenAuditDateEndInputCurrentDate'] = end.strftime('%m/%Y')
                    self.data[key]['queryForm:hiddenRecordDateEndInputCurrentDate'] = end.strftime('%m/%Y')
                    self.data[key]['queryForm:finishtimeStartTimeHours'] = datetime.now().strftime('%H')
                    self.data[key]['queryForm:finishtimeStartTimeMinutes'] = datetime.now().strftime('%M')

                # 改动10：用os.path.join拼接路径
                path = os.path.join(self.folder_temp, f"{guidang}_{year}_{month}{self.down_suffix}")
                down_file(self.URL, self.data, path)

    def main(self):
        for guidang in ['Y', 'N']:
            self.down(guidang)
            if guidang == 'Y':
                output_path = str(self.output_path).replace('.xlsx', '已归档.xlsx')
            else:
                output_path = str(self.output_path).replace('.xlsx', '未归档.xlsx')
            concat_df(self.folder_temp, output_path)
        log_downtime(self.down_name_en)

"""
爬取移动接口工单，路径：我的工作台-运维管理综合查询-[故障来源：移动运营商接口][当天工单][工单状态：待故障确认]-查询-工单列表-导出
"""
class YiDongOrder():
    def __init__(self):
        self.data = foura_data.yidong_order
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/billDeal/monitoring/list/billList.xhtml'
        self.down_name = '移动接口工单'
        self.down_name_en = 'yidong_order'
        self.down_suffix = '.xls'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        for index, city in enumerate(
                ['0099977', '0099978', '0099979', '0099980', '0099981', '0099982', '0099983', '0099984', '0099985',
                 '0099986', '0099987', '0099988', '0099989', '0108648', '0108649', '0108650', '0108651']):
            for key in ['1', '2']:
                self.data[key]['queryForm:dealendtimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                self.data[key]['queryForm:dealstarttimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                self.data[key]['queryForm:endtimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                self.data[key]['queryForm:revertendtimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                self.data[key]['queryForm:revertstarttimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                self.data[key]['queryForm:starttimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                self.data[key]['queryForm:queryUnitId'] = city
            # 改动12：用os.path.join拼接路径
            path = os.path.join(self.folder_temp, f"{index}{self.down_suffix}")
            down_file(self.URL, self.data, path)

    def main(self):
        self.down()
        concat_df(self.folder_temp, self.output_path)
        log_downtime(self.down_name_en)

"""
爬取历史所有运营商接口工单，路径：我的工作台-运维管理综合查询-[故障来源：联通+移动+电信运营商接口][当天前历史工单][回单和受理时间：上月1日到本月1日]-查询-导出
每次爬取上个月整月的，本月的不爬取
"""
class YunYingShangOrderHistory():
    def __init__(self, end=datetime.now()):
        self.data = foura_data.yunyingshang_order_history
        self.end = end.replace(day=1, hour=0, minute=0, second=0)
        self.begin = (self.end - timedelta(days=1)).replace(day=1, hour=0, minute=0, second=0)
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/billDeal/monitoring/list/billList.xhtml'
        self.down_name = '运营商接口工单'
        self.down_name_en = 'yunyingshang_order_history'
        self.down_suffix = '.xls'
        # 改动13：统一使用settings.resolve_path解析路径
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(
            f"spider/down/{self.down_name_en}/{self.down_name}{self.begin.strftime('%Y-%m')}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        for index, city in enumerate(
                ['0099977', '0099978', '0099979', '0099980', '0099981', '0099982', '0099983', '0099984', '0099985',
                 '0099986', '0099987', '0099988', '0099989', '0108648', '0108649', '0108650', '0108651']):
            for belong in ['移动运营商接口', '联通运营商接口', '电信运营商接口']:
                for key in ['1', '2']:
                    self.data[key]['queryForm:dealendtimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                    self.data[key]['queryForm:dealstarttimeInputCurrentDate'] = self.now.strftime('%m/%Y')
                    self.data[key]['queryForm:endtimeInputCurrentDate'] = self.end.strftime('%m/%Y')
                    self.data[key]['queryForm:revertendtimeInputCurrentDate'] = self.end.strftime('%m/%Y')
                    self.data[key]['queryForm:revertstarttimeInputCurrentDate'] = self.begin.strftime('%m/%Y')
                    self.data[key]['queryForm:starttimeInputCurrentDate'] = self.begin.strftime('%m/%Y')

                    self.data[key]['queryForm:revertendtimeInputDate'] = self.end.strftime('%Y-%m-%d %H:%M')
                    self.data[key]['queryForm:revertstarttimeInputDate'] = self.begin.strftime('%Y-%m-%d %H:%M')
                    self.data[key]['queryForm:starttimeInputDate'] = self.begin.strftime('%Y-%m-%d %H:%M')
                    self.data[key]['queryForm:endtimeInputDate'] = self.end.strftime('%Y-%m-%d %H:%M')
                    self.data[key]['queryForm:queryUnitId'] = city
                    self.data[key]['queryForm:faultSrc_hiddenValue'] = belong
                    self.data[key]['queryForm:faultSrc'] = belong
                # 改动14：用os.path.join拼接路径
                path = os.path.join(self.folder_temp, f"{city}{belong}{self.down_suffix}")
                down_file(self.URL, self.data, path)

    def main(self):
        self.down()
        concat_df(self.folder_temp, self.output_path)
        log_downtime(self.down_name_en)

"""
爬取录入异常设备清单，路径：资源管理-录入异常设备清单-导出
"""
class LuRuYiChang():
    def __init__(self):
        self.data = foura_data.lururyichang
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/dataMge/listInputEx.xhtml'
        self.down_name = '录入异常清单'
        self.down_name_en = 'luruyichang'
        self.down_suffix = '.xls'
        # 改动15：统一使用settings.resolve_path解析路径
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        for index, city in enumerate(
                ['0099977', '0099978', '0099979', '0099980', '0099981', '0099982', '0099983', '0099984', '0099985',
                 '0099986', '0099987', '0099988', '0099989', '0099990']):
            for key in ['1']:
                self.data[key]['queryForm:unitHidden'] = city
            # 改动16：用os.path.join拼接路径
            path = os.path.join(self.folder_temp, f"{index}{self.down_suffix}")
            down_file(self.URL, self.data, path)

    def main(self):
        self.down()
        concat_df(self.folder_temp, self.output_path)
        log_downtime(self.down_name_en)

"""
爬取运营商站址关系，路径：客户关系-运营商站址关系匹配-导出
"""
class StationAlias():
    def __init__(self):
        self.data = foura_data.station_alias
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/operatorShipMge/listOperator.xhtml'
        self.down_name = '运营商站址关系匹配'
        self.down_name_en = 'stationalias'
        self.down_suffix = '.xls'
        # 改动17：统一使用settings.resolve_path解析路径
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        for index, city in enumerate(
                ['0099977,0099978', '0099979,0099980,0099981', '0099982,0099983,0099984', '0099985,0099986,0099987',
                 '0099988,0099989,0099990']):
            for key in ['1']:
                self.data[key]['queryForm:unitHidden'] = city
            # 改动18：用os.path.join拼接路径
            path = os.path.join(self.folder_temp, f"{index}{self.down_suffix}")
            down_file(self.URL, self.data, path)

    def main(self):
        self.down()
        concat_df(self.folder_temp, self.output_path)
        log_downtime(self.down_name_en)

"""
爬取运营商站址关系，路径：运行监控-FSU监控-[注册状态：离线]-查询-导出
"""
class FsuJianKong():
    def __init__(self):
        self.data=foura_data.fsu_jiankong
        self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/pwMge/fsuMge/listFsu.xhtml'
        self.down_name_en='fsu_hafhour'
        self.down_name_en1='fsu_jiankong_5min'
        self.output_path = settings.resolve_path(f'updatenas/fsu/每半小时fsu离线/{datetime.now().strftime("%Y%m%d_%H%M")}fsu离线.xlsx')
        self.temp_path=settings.resolve_path('spider/down/temp_folder_one_day/fsu5分钟.xlsx')
    def down(self):
        down_file(self.URL, self.data, self.output_path)
        if datetime.now().minute < 30:
            new_path = str(self.output_path).replace('每半小时', '每小时')
            shutil.copy(self.output_path, new_path)
        log_downtime(self.down_name_en)
    def down_5min(self):
        down_file(self.URL, self.data, self.temp_path)
        self.sql_process(self.temp_path)
        log_downtime(self.down_name_en1)
    def sql_process(self,path):
        df = pd.read_excel(path, dtype=str)
        df=df.fillna('')
        df=df.loc[df['离线时间']!='']
        with sql_orm().session_scope() as temp:
            sql, Base = temp
            pojo_brokentime = Base.classes.fsu_brokentime_log
            pojo_brokentimes = Base.classes.fsu_brokentimes_log
            now = datetime.now()
            if now.hour == 7:
                res = sql.query(pojo_brokentimes).all()
                for log in res:
                    log.broken_times = 0
            for index, row in df.iterrows():
                res = sql.query(pojo_brokentimes).filter(pojo_brokentimes.id == row['站址']).first()
                if res == None:
                    # 次数统计
                    log = pojo_brokentimes()
                    log.id = row['站址']
                    log.begin_time = row['离线时间']
                    log.broken_times = 1
                    sql.merge(log)
                    # 离线记录
                    log = pojo_brokentime()
                    log.id = row['站址']
                    log.begin_time = row['离线时间']
                    sql.add(log)
                else:
                    begin_time = datetime.strptime(res.begin_time, '%Y/%m/%d  %H:%M:%S')
                    begin_time_row = datetime.strptime(row['离线时间'], '%Y/%m/%d  %H:%M:%S')
                    if begin_time_row > begin_time:
                        # 次数统计
                        res.begin_time = row['离线时间']
                        res.broken_times += 1
                        # 离线记录
                        log = pojo_brokentime()
                        log.id = row['站址']
                        log.begin_time = row['离线时间']
                        sql.add(log)

"""
爬取历史告警Hbase，路径：运行监控-历史告警Hbase-[告警发生时间区间,告警名称]-导出excel
"""
class AlarmHistoryHbase():
    def __init__(self, year=0, month=0):
        self.data = foura_data.alarm_history_Hbase
        if month != 0:
            self.now = datetime.now().replace(year=year, month=month, day=1)
        else:
            self.now = datetime.now()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/alarmHisHbaseMge/listHisAlarmHbase.xhtml'
        self.down_name = '历史告警'
        self.down_name_en = 'Hbase'
        self.down_suffix = '.xlsx'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def df_sql_process(self):
        for file in os.listdir(self.folder_temp):
            file = os.path.join(self.folder_temp, file)
            try:
                df = pd.read_excel(file, dtype=str)
                df = df.drop_duplicates()
                df = df.fillna('')
                df['告警发生日期'] = df['告警发生时间']
                with sql_orm().session_scope() as temp:
                    sql, Base = temp
                    pojo = Base.classes.hbase
                    for index, row in df.iterrows():
                        temp = pojo(**row.to_dict())
                        sql.merge(temp)
            except Exception as e:
                pass

    def down(self):
        clear_folder(self.folder_temp)
        for alarm_name in ['一级低压脱离告警', '二级低压脱离告警', 'FSU离线', '温度超高', '温度过高',
                           '交流输入停电告警', '总电压过低', '直流输出电压过高告警', '直流输出电压过低告警'
            , '门', '电池供电告警']:
            end = self.now.replace(hour=0, minute=0, second=0)
            if end.day == 1:
                start = (self.now - timedelta(days=1)).replace(day=1, hour=0, minute=0, second=0)
            else:
                start = self.now.replace(day=1, hour=0, minute=0, second=0)
            # 日期循环
            while start < end:
                start_next = (start + timedelta(days=1))
                for key in ['1']:
                    self.data[key]['queryForm:queryalarmName'] = alarm_name
                    self.data[key]['queryForm:firststarttimeInputCurrentDate'] = start.strftime('%m/%Y')
                    self.data[key]['queryForm:firstendtimeInputCurrentDate'] = start_next.strftime('%m/%Y')
                    self.data[key]['queryForm:firststarttimeInputDate'] = start.strftime('%Y-%m-%d %H:%M')
                    self.data[key]['queryForm:firstendtimeInputDate'] = start_next.strftime('%Y-%m-%d %H:%M')

                    # 改动20：用os.path.join拼接路径
                    path = os.path.join(self.folder_temp, f"{alarm_name}{start.strftime('%Y%m%d')}{self.down_suffix}")
                    down_file(self.URL, self.data, path)
                start = start_next

    def main(self):
        self.down()
        self.df_sql_process()
        log_downtime(self.down_name_en)

"""
爬取性能查询，路径：运行监控-性能查询-查询-导出
"""
class Performence():
    # 路径：运行监控-性能查询-查询-导出
    def __init__(self):
        self.data = foura_data.performence
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml'

    def main(self, cities, search_id, folder_temp, out_put, hours=24, content_len=3000, csv=False):
        def down_by_city(city):
            for key in ['1', '2']:
                now = datetime.now()
                start_time = now - timedelta(hours=int(hours))

                # 格式时间化
                end_time_input_date = now.strftime("%Y-%m-%d %H:%M")
                start_time_input_date = start_time.strftime("%Y-%m-%d %H:%M")
                end_time_input_current_date = now.strftime("%m/%Y")
                start_time_input_current_date = start_time.strftime("%m/%Y")

                self.data[key]['queryForm:endtimeInputCurrentDate'] = end_time_input_current_date
                self.data[key]['queryForm:endtimeInputDate'] = end_time_input_date
                self.data[key]['queryForm:starttimeInputCurrentDate'] = start_time_input_current_date
                self.data[key]['queryForm:starttimeInputDate'] = start_time_input_date
                self.data[key]['queryForm:querySpeId'] = search_id
                self.data[key]['queryForm:querySpeIdShow'] = search_id + "..."
                self.data[key]['queryForm:unitHidden2'] = city

            path = os.path.join(folder_temp, f"{city}_{search_id}.xlsx")
            down_file(self.URL, self.data, path, xlsx_juge=True, conten_len_error=int(content_len),
                      )

        for city in cities:
            down_by_city(city)
        # 用于临时下载没成功的文件
        # for city in cities:
        #     flag = 0
        #     for file in os.listdir(folder_temp):
        #         if city in file:
        #             flag=1
        #     if flag==0:
        #         down_by_city(city)
        concat_df(folder_temp, out_put, gen_csv=csv)

"""
爬取性能查询，路径：运行监控-性能查询-查询-导出
"""
class PerformenceBySiteList():
    # 路径：运行监控-性能查询-查询-导出
    # 站址运维ID入参
    def __init__(self):
        self.data = foura_data.performence.copy()
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml'
        del self.data['FINAL']

    def main(self, site_list, search_id, minutes_back=1440):
        df_list = []
        page_size = 200
        for i in range(0, len(site_list), page_size):
            chunk = site_list[i:i + page_size]
            for key in ['1', '2']:
                now = datetime.now()
                start_time = now - timedelta(minutes=minutes_back)

                def format_datetime(dt):
                    return dt.strftime("%Y-%m-%d %H:%M"), dt.strftime("%m/%Y")

                start_time_input_date, start_time_input_current_date = format_datetime(start_time)
                end_time_input_date, end_time_input_current_date = format_datetime(now)
                self.data[key].update({
                    'queryForm:starttimeInputDate': start_time_input_date,
                    'queryForm:starttimeInputCurrentDate': start_time_input_current_date,
                    'queryForm:endtimeInputDate': end_time_input_date,
                    'queryForm:endtimeInputCurrentDate': end_time_input_current_date,
                })
                self.data[key]['queryForm:endtimeTimeHours'] = now.strftime("%H")
                self.data[key]['queryForm:endtimeTimeMinutes'] = now.strftime("%M")
                self.data[key]['queryForm:querySpeId'] = search_id
                self.data[key]['queryForm:querySpeIdShow'] = search_id + "..."
                self.data[key]['queryForm:queryStationId'] = ','.join(chunk)
                self.data[key]['queryForm:queryStationIdShow'] = f'{chunk[0]}...'
                self.data[key]['queryForm:pageSizeText'] = page_size * 4
            df = down_file_no_save(self.URL, self.data)
            df_list.append(df)
        return pd.concat(df_list)

"""
爬取某个站的性能查询，路径：运行监控-性能查询-查询-导出
"""
class SerchPerformence():
    def __init__(self):
        self.data = foura_data.performence
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml'
        self.headers = {
            'Host': 'omms.chinatowercom.cn:9000',
            'Origin': 'http://omms.chinatowercom.cn:9000',
            'Referer': 'http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
        }

    def serch_performence_by_id(self, stationid, serch_id):
        cookies = get_foura_cookie()
        signal = ''
        self.data['2']['queryForm:querySpeId'] = f'{serch_id}'
        self.data['2']['queryForm:querySpeIdShow'] = f'{serch_id}...'
        try:
            res = requests_post(self.URL, headers=self.headers, cookies=cookies)
            html = BeautifulSoup(res.text, 'html.parser')
            javax = html.find('input', id='javax.faces.ViewState')['value']
            into_data = self.data['2']
            into_data['javax.faces.ViewState'] = javax
            now = datetime.now()
            before = now - timedelta(days=1)

            into_data['queryForm:queryStationId'] = f'{stationid}'
            into_data['queryForm:queryStationIdShow'] = f'{stationid}...'

            into_data['queryForm:starttimeInputDate'] = before.strftime('%Y-%m-%d %H:%M')
            into_data['queryForm:starttimeInputCurrentDate'] = before.strftime('%m/%Y')
            into_data['queryForm:endtimeInputDate'] = now.strftime('%Y-%m-%d %H:%M')
            into_data['queryForm:endtimeInputCurrentDate'] = now.strftime('%m/%Y')

            res = requests_post(self.URL, headers=self.headers, data=into_data, cookies=cookies)
            html = BeautifulSoup(res.text, 'html.parser')
            tbody = html.find('tbody', id='listForm:list:tb')
            tr = tbody.find_all('tr')
            for tr_item in tr:
                td = tr_item.find_all('td')
                if td[11].center.text == serch_id:
                    signal = td[14].center.text
                    signal = signal[:6]
            return signal
        except Exception as e:
            return signal

"""
爬取历史性能查询，路径：运行监控-历史性能查询-查询
"""
class HisPerformence():
    def __init__(self):
        self.data = foura_data.his_performence
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml'

    def _get_headers(self):
        return {
            'Host': 'omms.chinatowercom.cn:9000',
            'Origin': 'http://omms.chinatowercom.cn:9000',
            'Referer': self.URL,
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36'
        }

    def _get_viewstate(self, html):
        soup = BeautifulSoup(html, 'html.parser')
        javax = soup.find('input', id='javax.faces.ViewState')
        return javax['value']

    def _parse_devices(self, html, station_name):
        """解析所有设备，返回 {后缀: (did, 完整设备名, 数字)} 字典，相同后缀取编号最小"""
        devices = {}
        pattern = r'id="(\d+)"[^>]*name="rDevice"[^>]*value="([^"]+)"'
        matches = re.findall(pattern, html)

        for did, full_name in matches:
            if not full_name.startswith(station_name + '/'):
                continue

            suffix_part = full_name[len(station_name) + 1:]
            match = re.match(r'(.+?)(\d+)$', suffix_part)
            if match:
                suffix, num_str = match.group(1), match.group(2)
                num = int(num_str)
            else:
                suffix = suffix_part
                num = 0

            if suffix not in devices or num < devices[suffix][2]:
                devices[suffix] = (did, full_name, num)

        return devices

    def _query_data(self, s, headers, cookies, javax, aid, name, device_name, did, mid, start_time, end_time):
        """查询设备的历史数据，返回 [(时间, 实测值), ...] 列表"""
        d3 = self.data['FINAL'].copy()
        d3.update({
            'javax.faces.ViewState': javax,
            'queryFormB:aid': aid,
            'queryFormB:siteNameId': name,
            'queryFormB:deviceName': device_name,
            'queryFormB:did': did,
            'queryFormB:mid': mid,
            'queryFormB:starttimeInputDate': start_time,
            'queryFormB:starttimeInputCurrentDate': start_time[5:7] + '/' + start_time[:4],
            'queryFormB:endtimeInputDate': end_time,
            'queryFormB:endtimeInputCurrentDate': end_time[5:7] + '/' + end_time[:4]
        })
        r3 = s.post(self.URL, headers=headers, data=d3, cookies=cookies)

        results = []
        soup = BeautifulSoup(r3.text, 'html.parser')
        tbody = soup.find('tbody', id='listFormB:list:tb')
        if tbody:
            rows = tbody.find_all('tr')
            for row in rows:
                cells = row.find_all('td')
                if len(cells) >= 15:
                    time_val = cells[13].text.strip()
                    measure_val = cells[14].text.strip()
                    results.append((time_val, measure_val))
        return results

    def main(self, name, device_mid_dict, hours=2):
        """
        入参：
            name: 站址名
            device_mid_dict: {'后缀': ['mid1', 'mid2', ...], ...}
            hours: 查询过去几小时，默认2
            cookie_user: cookie用户标识
        返回：
            {设备完整名: {mid: [(时间, 实测值), ...], ...}, ...}
        """
        s = requests.Session()
        headers = self._get_headers()
        cookies = get_foura_cookie()
        now = datetime.now()
        past = now - timedelta(hours=hours)
        start_time = past.strftime('%Y-%m-%d %H:%M')
        end_time = now.strftime('%Y-%m-%d %H:%M')

        r0 = s.post(self.URL, headers=headers, cookies=cookies)
        javax = self._get_viewstate(r0.text)

        d1 = self.data['1'].copy()
        d1['stationListFormB:nameText'] = name
        d1['javax.faces.ViewState'] = javax
        r1 = s.post(self.URL, headers=headers, data=d1, cookies=cookies)
        aid = re.search(r'id="([^"]+)"[^>]*name="selectFlagB"[^>]*value="%s"' % re.escape(name), r1.text)
        if not aid:
            aid = re.search(r'id="([^"]+)"[^>]*name="selectFlagB"', r1.text)
        aid = aid.group(1)

        d2 = self.data['2'].copy()
        d2['queryForm2B:aid'] = aid
        d2['javax.faces.ViewState'] = javax
        r2 = s.post(self.URL, headers=headers, data=d2, cookies=cookies)

        all_devices = self._parse_devices(r2.text, name)

        results = {}
        for suffix, mid_list in device_mid_dict.items():
            # 模糊匹配设备
            matched = None
            for dev_suffix, (did, full_name, num) in all_devices.items():
                if suffix in dev_suffix:
                    matched = (did, full_name)
                    break

            if matched:
                did, full_name = matched
                device_results = {}
                for mid in mid_list:
                    data_list = self._query_data(s, headers, cookies, javax, aid, name, full_name, did, mid, start_time,
                                                 end_time)
                    device_results[mid] = data_list
                results[full_name] = device_results
        return results

"""
爬取当前活动告警，从数据库拿，字段不完全
"""
class AlarmNow():
    def __init__(self):
        self.db_fields = [
            "ID", "告警类别", "告警对象类型", "告警对象ID", "告警对象名称",
            "告警子对象类型", "告警子对象ID", "告警子对象名称", "原始告警等级", "告警等级",
            "告警原因", "告警摘要", "告警详情", "告警附加信息", "告警发生时间",
            "首次告警的系统时间", "末次告警的系统时间", "告警重复次数", "是否已确认", "确认时间",
            "确认人", "告警类型", "用户备注信息", "是否已恢复", "是否转故障",
            "故障单号", "是否关注", "告警来源", "告警名称", "运维监控站址ID",
            "FSU名称", "告警流水号", "工单状态", "工单编号", "站址名称",
            "告警发生时站址状态", "省", "地市", "站址运维ID", "维修室人员编号",
            "所属主要告警ID", "告警上报时的站址维护状态", "首次告警时间至目前经过时间（分）",
            "离首次告警时间是否超过24小时", "信号量说明",
            "信号量解释", "任务状态", "站址来源", "确认原因", "信号量"
        ]
        self.path_csv = settings.resolve_path('message/battery_life/xls/活动告警.csv')

    def main(self):
        res = requests.get(r'http://clound.gxtower.cn:3980/tt/get_alarm')
        with open(self.path_csv, "wb") as codes:
            codes.write(res.content)

        # 读取时处理空值
        df = pd.read_csv(self.path_csv, dtype=str, keep_default_na=False).rename(
            columns={"运维ID": "站址运维ID", "告警入库时间": "告警发生时间"})
        df = df.replace('', None)
        df.to_csv(self.path_csv, index=False, encoding='utf-8-sig')

        df.columns = df.columns.str.strip()

        # 构建数据
        df_db = pd.DataFrame()
        for field in self.db_fields:
            df_db[field] = df[field] if field in df.columns else None

        # 过滤
        df_db = df_db[df_db["ID"].notna() & (df_db["ID"] != "")]

        with sql_orm().session_scope() as temp:
            sql, Base = temp
            sql.query(Base.classes.alarm_now).delete()
            sql.commit()

            pojo = Base.classes.alarm_now

            # 逐行处理，确保没有 NaN
            for _, row in df_db.iterrows():
                # 把 Series 转成 dict，并处理 NaN
                data = {}
                for k, v in row.items():
                    if pd.isna(v):
                        data[k] = None
                    else:
                        data[k] = v
                sql.add(pojo(**data))

            sql.commit()
        # 南分备份需求
        if datetime.now().minute >= 40:
            backup_dir = settings.resolve_path('message/nanfen_overtime/alarm_backup')
            # 改动21：增加备份目录创建
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            current_time = datetime.now().strftime("%Y%m%d%H")
            new_filename = f'{current_time}.xlsx'
            target_path = os.path.join(backup_dir, new_filename)
            df.to_excel(target_path, index=False)
        return

"""
爬取4A活动告警，允许自定义参数，单独调用,路径：4A运维监控系统-运行监控-告警管理-活动告警监控-[告警等级:一级告警、二级告警，FSU状态：无]-查询-导出
"""
class AlarmNow4AByCity():
    def __init__(self):
        self.data = foura_data.alarm_now
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/alarmMge/listAlarm.xhtml'
        self.down_name = '活动告警'
        self.down_name_en = 'alarm_now'
        self.down_suffix = '.xls'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")
    def down1(self,city,path):
        for key in ['1']:
            self.data[key]['queryForm:unitHidden'] = city
        down_file(self.URL, self.data, path)
    def down(self):
        down_list = ['0099977', '0099978', '0099979', '0099980', '0099981', '0099982', '0099983', '0099984', '0099985',
                     '0099986', '0099987', '0099988', '0099989', '0099990']
        for city in down_list:
            for key in ['1','2']:
                self.data[key]['queryForm:unitHidden'] = city
            path = os.path.join(self.folder_temp, f"{city}{self.down_suffix}")
            down_file(self.URL, self.data, path)
    def read_file(self):
        df_list = []
        for file in os.listdir(self.folder_temp):
            path = os.path.join(self.folder_temp, file)
            if '.xls' in file:
                temp = pd.read_excel(path, dtype=str)
                df_list.append(temp)
        return df_list

    def main(self):
        self.down()
        if len(os.listdir(self.folder_temp)) >= 14:
            df_list = self.read_file()
            merge = pd.concat(df_list)
            merge.to_excel(self.output_path, index=False)
            log_downtime(self.down_name_en)
        else:
            raise FileNotFoundError("文件下载不全，当前仅下载了 {} 个文件".format(len(os.listdir(self.folder_temp))))

"""
爬取故障监控，路径：4A运维监控系统-运行监控-故障管理-故障监控-[活动\历史故障监控,故障场景:退服场景,时间范围：当月一号到今天]-查询-导出
"""
class FaultMonitoring():
    def __init__(self):
        self.data = foura_data.fault_monitoring
        self.URL = 'http://omms.chinatowercom.cn:9000/business/resMge/faultAlarmMge/listFaultActive.xhtml'
        self.now = datetime.now()
        self.start_date_str = datetime(self.now.year, self.now.month, 1, 0, 0)
        self.end_date_str = datetime(self.now.year, self.now.month, self.now.day, 0, 0)
        self.down_name = '故障监控'
        self.down_name_en = 'fault_monitoring'
        self.down_suffix = '.xls'
        self.folder_temp = settings.resolve_path(f'spider/down/{self.down_name_en}/temp')
        self.output_path = settings.resolve_path(f"spider/down/{self.down_name_en}/{self.down_name}.xlsx")

    def down(self):
        clear_folder(self.folder_temp)
        down_list = ['0099977', '0099978', '0099979', '0099980', '0099981', '0099982', '0099983', '0099984', '0099985',
                     '0099986', '0099987', '0099988', '0099989', '0099990']
        for city in down_list:
            for key in ['1']:
                self.data[key]['hisQueryForm:unitHidden'] = city
                self.data[key]["hisQueryForm:treeCityId"] = city
                self.data[key]["hisQueryForm:firststarttimeInputDate"] = self.start_date_str.strftime('%Y-%m-%d %H:%M'),
                self.data[key]["hisQueryForm:firststarttimeInputCurrentDate"] = self.now.strftime('%m/%Y'),
                self.data[key]["hisQueryForm:firstendtimeInputDate"] = self.end_date_str.strftime('%Y-%m-%d %H:%M'),
                self.data[key]["hisQueryForm:firstendtimeInputCurrentDate"] = self.now.strftime('%m/%Y'),
                self.data[key]["hisQueryForm:recoverstarttimeInputCurrentDate"] = self.now.strftime('%m/%Y'),
                self.data[key]["hisQueryForm:recoverendtimeInputCurrentDate"] = self.now.strftime('%m/%Y'),
            path = os.path.join(self.folder_temp, f"{city}{self.down_suffix}")
            down_file(self.URL, self.data, path)

    def read_file(self):
        df_list = []
        for file in os.listdir(self.folder_temp):
            path = os.path.join(self.folder_temp, file)
            if '.xls' in file:
                temp = pd.read_excel(path, dtype=str)
                df_list.append(temp)
        return df_list

    def main(self):
        self.down()
        if len(os.listdir(self.folder_temp)) >= 14:
            df_list = self.read_file()
            merge = pd.concat(df_list)
            merge.to_excel(self.output_path, index=False)
            log_downtime(self.down_name_en)
        else:
            raise FileNotFoundError("文件下载不全，当前仅下载了 {} 个文件".format(len(os.listdir(self.folder_temp))))

"""
爬取故障监控，路径：运行监控-历史性能查询-移动/联通/电信电量存量
"""
class BatteryLevel():
    def __init__(self):
        self.db_fields = ['站址', '设备', '编号', '信号量', '实测值', '时间', '备注']
        self.station_list = [
            '南宁市金融大厦O', '南宁市幼师大厦O', '西乡塘区五菱桂花车辆有限公司',
            '西乡塘区相思湖西路监控杆', 'GX南宁越秀龙盘小区灯塔', '南宁市官桥村O',
            '南宁市建华小学O', '南宁市圣诺宾馆O', '邕宁区中和乡方村南',
            'GX南宁邕宁中和那烈', '桂林市锁厂O-DA', '兴安县志玲路国税局O-DB',
            '桂林市顶新酒店O-DA', '恭城三江十八岭村O-DC', '灵川定江郑家村O-DB',
            '临桂新城吾悦华府小区南侧山坡（搬迁）'
        ]
        self.all_data = []
        self.device_mid_dict = {
            "开关电源": {
                "0406136001": "移动租户电量",
                "0406138001": "联通租户电量",
                "0406140001": "电信租户电量",
                "0406161001": "包干型新业务租户电量",
                "0406163001": "计量型新业务租户XX电量",
                "0406167001": "配电单元XX电量",
                "0406170001": "广电租户电量",
                "0406172001": "行业外租户电量",
                "0445109001": "包干型新业务租户电量",
                "0445111001": "计量型新业务租户XX电量",
                "0445113001": "移动租户电量（5G）",
                "0445115001": "联通租户电量（5G）",
                "0445117001": "电信租户电量（5G）"
            },
            "分路计量": {
                "0445102001": "移动租户电量",
                "0445104001": "联通租户电量",
                "0445106001": "电信租户电量",
                "0406161001": "包干型新业务租户电量",
                "0406163001": "计量型新业务租户XX电量",
                "0406167001": "配电单元XX电量",
                "0406170001": "广电租户电量",
                "0406172001": "行业外租户电量",
                "0445109001": "包干型新业务租户电量",
                "0445111001": "计量型新业务租户XX电量",
                "0445113001": "移动租户电量（5G）",
                "0445115001": "联通租户电量（5G）",
                "0445117001": "电信租户电量（5G）"
            }
        }

    def down(self):
        for station in self.station_list:
            result = HisPerformence().main(name=station, device_mid_dict=self._flatten_mid(), hours=2)
            has_data = False
            for device, mid_data in result.items():
                suffix = next((k for k in self.device_mid_dict if k in device), None)
                if not suffix:
                    continue
                for mid, remark in self.device_mid_dict[suffix].items():
                    data_list = mid_data.get(mid, [])
                    if data_list:
                        has_data = True
                        for time, value in data_list:
                            self.all_data.append({
                                '站址': station, '设备': device, '信号量': mid,
                                '实测值': value if value else None, '时间': time, '备注': remark
                            })
                    else:
                        has_data = True
                        self.all_data.append({
                            '站址': station, '设备': device, '信号量': mid,
                            '实测值': None, '时间': datetime.now(), '备注': remark
                        })
            if not has_data:
                self.all_data.append({
                    '站址': station, '设备': None, '信号量': None,
                    '实测值': None, '时间': datetime.now(), '备注': None
                })

    def _flatten_mid(self):
        return {k: list(v.keys()) for k, v in self.device_mid_dict.items()}

    def df_process(self):
        df = pd.DataFrame(self.all_data)
        df = df.replace({np.nan: None})

        if '设备' in df.columns:
            split_df = df['设备'].str.rsplit('/', n=1, expand=True)
            df['设备'] = split_df[0]  # 前一段
            df['编号'] = split_df[1]  # 后一段

        with sql_orm().session_scope() as temp:
            sql, Base = temp
            pojo = Base.classes.电量信号量
            for _, row in df.iterrows():
                sql.add(pojo(**row.to_dict()))
            sql.commit()

    def main(self):
        self.down()
        self.df_process()


if __name__ == '__main__':
    # FsuChaXun().main()
    # FaultMonitoring().main()
    # AlarmNow4AByCity().main()
    # StationLiangYi().main()
    # Station().main()
    # StationDC().main()
    # BatteryLevel().main()
    # AlarmHistoryHbase().main()
    # YiDongOrder().main()
    YinHuanOrder().main()