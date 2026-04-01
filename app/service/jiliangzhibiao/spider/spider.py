import os
import re
import time
import requests
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from functools import wraps
from xlsx2csv import Xlsx2csv
from sqlalchemy import text
from core.config import settings

"""
合并完成版：统一Cookie、统一会话、所有请求共享一套登录信息
"""

# ===================== 全局统一配置（只改这里就行） =====================
# 全局唯一 Cookie（所有请求共用）
cookies_str = "route=73d99f9520cc61666c3e6505999568f9; JSESSIONID=943A5DFB9881030C46D94B24BD40B4F8; acctId=100852210; uid=dw.rj.fengsw; moduleUrl=/layout/index.xhtml; nodeInformation=172.29.105.18:all8280; BIGipServerywjk_new_pool1=42016172.10275.0000; Hm_lvt_f6097524da69abc1b63c9f8d19f5bd5b=1775023804; Hm_lpvt_f6097524da69abc1b63c9f8d19f5bd5b=1775023804; HMACCOUNT=4928EF137877E6D0; pwdaToken=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzdWIiOiJSRVMiLCJpc3MiOiJXUzRBIiwiZXhwIjoxNzc1MDMyOTc5LCJOQU5PU0VDT05EIjoxMzU2ODg0Njc5MTY0NzQ2fQ.eI0To_mlECnerL7VOKYJSkj_GafxL4W5Zi434zg3SS8"

# 全局统一请求头
GLOBAL_HEADERS = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Connection": "keep-alive",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Cookie": cookies_str,
    "Host": "omms.chinatowercom.cn:9000",
    "Origin": "http://omms.chinatowercom.cn:9000",
    "Referer": "http://omms.chinatowercom.cn:9000/devMge/initShuntMetering.go",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0",
    "X-Requested-With": "XMLHttpRequest"
}

# ===================== 通用工具函数 =====================
def parse_cookie(cookie_str):
    """解析Cookie字符串为字典（全局共用）"""
    cookies = {}
    try:
        for item in cookie_str.split(";"):
            item = item.strip()
            if not item:
                continue
            key, value = item.split("=", 1)
            cookies[key] = value
    except Exception as e:
        print(f"Cookie解析失败: {e}")
    return cookies

GLOBAL_COOKIE_DICT = parse_cookie(cookies_str)

def retry(max_attempts=15, delay=2):
    """重试装饰器"""
    def decorator_retry(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            attempts = 0
            while attempts < max_attempts:
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    attempts += 1
                    if attempts == max_attempts:
                        raise
                    print(f"重试 {attempts}/{max_attempts}: {str(e)[:50]}")
                    time.sleep(delay)
        return wrapper
    return decorator_retry

@retry()
def requests_post(url, headers=GLOBAL_HEADERS, data={}, cookies=GLOBAL_COOKIE_DICT, timeout=600):
    return requests.post(url, headers=headers, data=data, cookies=cookies, timeout=timeout)

@retry()
def requests_get(url, headers=GLOBAL_HEADERS, params={}, cookies=GLOBAL_COOKIE_DICT, timeout=600):
    return requests.get(url, headers=headers, params=params, cookies=cookies, timeout=timeout)

def xlsx_to_csv(folder):
    if not os.path.exists(folder):
        print(f"文件夹不存在: {folder}")
        return
    for file in os.listdir(folder):
        path = os.path.join(folder, file)
        if file.endswith('.xlsx'):
            csv_path = path.replace(".xlsx", ".csv")
            if not os.path.exists(csv_path):
                try:
                    Xlsx2csv(path, outputencoding="utf-8").convert(csv_path)
                    print(f"转换完成: {file}")
                except Exception as e:
                    print(f"转换失败 {file}: {e}")

def concat_df(folder, output_path, gen_csv=False):
    xlsx_to_csv(folder)
    df_list = []
    if not os.path.exists(folder):
        return pd.DataFrame(), output_path
    for file in os.listdir(folder):
        path = os.path.join(folder, file)
        try:
            if '.csv' in file:
                temp = pd.read_csv(path, dtype=str)
                df_list.append(temp)
            elif file.endswith('.xls'):
                temp = pd.read_excel(path, dtype=str, engine='xlrd')
                df_list.append(temp)
        except:
            pass
    merge_df = pd.concat(df_list, ignore_index=True) if df_list else pd.DataFrame()
    merge_df.to_excel(output_path, index=False)
    if gen_csv:
        merge_df.to_csv(str(output_path).replace('.xlsx','.csv'), index=False, encoding='utf-8-sig')
    return merge_df, output_path

def down_file(url, data, path, conten_len_error=3000, xlsx_juge=False):
    retry_count = 3
    while retry_count >= 0:
        try:
            headers = {
                'Host': 'omms.chinatowercom.cn:9000',
                'Origin': 'http://omms.chinatowercom.cn:9000',
                'Referer': url,
                'User-Agent': GLOBAL_HEADERS["User-Agent"],
            }
            res = requests_post(url, headers=headers, cookies=GLOBAL_COOKIE_DICT)
            html = BeautifulSoup(res.text, 'html.parser')
            javax = html.find('input', id='javax.faces.ViewState')['value']

            for key, into_data in data.items():
                into_data['javax.faces.ViewState'] = javax
                res = requests_post(url, headers=headers, data=into_data, cookies=GLOBAL_COOKIE_DICT)

                if 'FINAL' in key:
                    if len(res.content) < conten_len_error:
                        raise ValueError("内容过小")
                    if xlsx_juge:
                        if not res.content.startswith((b'\x50\x4B\x03\x04', b'\x09\x08\x04\x00\x10\x00\x00\x00')):
                            raise ValueError("非Excel文件")
                    os.makedirs(os.path.dirname(path), exist_ok=True)
                    with open(path, "wb") as f:
                        f.write(res.content)
                    print(f"✅ 下载成功: {path}")
            return
        except ValueError as e:
            retry_count -= 1
            if retry_count < 0:
                raise
            print(f"重试中... {e}")

# ===================== 分流计量下载（双请求） =====================
def download_shunt_meter_excel():
    save_path = settings.resolve_path(r"app/service/jiliangzhibiao/data/分流计量数据.xlsx")
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    data1 = {
        "queryCheckTime": datetime.now().strftime("%Y-%m-%d"),
        "queryDeviceType": "",
        "queryShareFlag": "",
        "queryUserFlag": "",
        "queryAccuracyFlag": "",
        "queryOnlineFlag": "",
        "querystaname": "",
        "querystacode": "",
        "queryDeviceID": "",
        "queryUnitId": "",
        "page": "1",
        "rows": "20"
    }

    session = requests.Session()
    session.headers.update(GLOBAL_HEADERS)
    session.cookies.update(GLOBAL_COOKIE_DICT)

    print("正在请求数据列表...")
    url1 = "http://omms.chinatowercom.cn:9000/devMge/getShuntMeteringData.go"
    resp1 = session.post(url1, data=data1, timeout=30)
    resp1.raise_for_status()

    print("正在导出Excel...")
    url2 = "http://omms.chinatowercom.cn:9000/devMge/exportSMDataExcel.go?ids=1"
    resp2 = session.get(url2, stream=True, timeout=60)
    resp2.raise_for_status()

    with open(save_path, "wb") as f:
        for chunk in resp2.iter_content(1024*1024):
            f.write(chunk)

    print(f"\n✅ 分流计量Excel下载完成：{save_path}")
    return save_path

# ===================== 主入口 =====================
if __name__ == "__main__":
    try:
        # 执行分流计量下载（双请求）
        download_shunt_meter_excel()


    except Exception as e:
        print(f"\n❌ 执行失败: {e}")