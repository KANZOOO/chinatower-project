from bs4 import BeautifulSoup
import os
import pandas as pd
from datetime import datetime,timedelta
#from core.sql import sql_orm
from xlsx2csv import Xlsx2csv
import time
import re
import requests
import numpy as np
from functools import wraps
from sqlalchemy import text
#from spider.schema.schema import data
#import calenda
#from core.config import settings
#import shutil
from spider.schema.data_edge_x import data
"""
通用函数模块，包含项目中常用的工具函数，如重试装饰器、请求函数、文件操作函数等。
"""
def get_foura_cookie():
    """
    从数据库中获取指定ID的foura cookie信息。

    :param ID: 数据库中cookie记录的ID，默认为1（对应foura1）。
    :return: 解析后的cookie字典；若ID不存在/格式错误，返回空字典。
    """
    # db = sql_orm()
    # cookie_result = db.get_cookies(f"foura")
    # return cookie_result["cookies"]
    # cookies_str =input("请输入cookie字符串：")
    cookies_str="Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1774852953; route=528ccec9890537400fabeee91bd86ea5; JSESSIONID=CEA3591F068E513C1C69B47A456B1684; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzc0ODY0ODQ1LCJOQU5PU0VDT05EIjoxMzU4NDA3MzYzMDE1OTM2M30.CbQVL7BrSTNSnXVOJoUM6hqId8EZWUkIlZ0pWBppnD8; acctId=101433247; uid=wx-yeping6; moduleUrl=/layout/index.xhtml; nodeInformation=10.195.54.7:all8180; BIGipServerywjk_new_pool1=260119980.10275.0000"
    cookies = {}
    try:
        for cookie in cookies_str.split(';'):
            key, value = cookie.split('=', 1)
            cookies[key] = value
    except:
        pass
    return cookies
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

# down_file(url="http://omms.chinatowercom.cn:9000/business/resMge/devMge/listEdgeGateway.xhtml",
#           data=data,
#           path=settings.resolve_path("spider/script/station/down/边缘网关.xls"))
