import copy
import os
import time
from datetime import datetime
from functools import wraps

import requests
from bs4 import BeautifulSoup

from core.config import settings
from core.sql import sql_orm
from app.service.jiliangzhibiao.spider.schema.schema_jilianghzibiao import (
    dianxin_jiliang,
    dianxin_jiliang5g,
    dianxin_kaiguan,
    liantong_jiliang,
    liantong_jiliang5g,
    liantong_kaiguan,
    yidong_jiliang,
    yidong_jiliang5g,
    yidong_kaiguan,
)


DEFAULT_COOKIE_ID = "dw.rj.fengsw"
BASE_URL = "http://omms.chinatowercom.cn:9000"
PERFORMANCE_URL = f"{BASE_URL}/business/resMge/pwMge/performanceMge/perfdata.xhtml"
SHUNT_QUERY_URL = f"{BASE_URL}/devMge/getShuntMeteringData.go"
SHUNT_EXPORT_URL = f"{BASE_URL}/devMge/exportSMDataExcel.go?ids=1"


LESSEE_DOWNLOAD_TASKS = [
    ("分路计量设备-移动租户电流", yidong_jiliang, "分路计量设备-移动租户电流.xls"),
    ("开关电源-移动租户电流", yidong_kaiguan, "开关电源-移动租户电流.xls"),
    ("分路计量设备-移动租户电流（5G）", yidong_jiliang5g, "分路计量设备-移动租户电流（5G）.xls"),
    ("分路计量设备-联通租户电流", liantong_jiliang, "分路计量设备-联通租户电流.xls"),
    ("开关电源-联通租户电流", liantong_kaiguan, "开关电源-联通租户电流.xls"),
    ("分路计量设备-联通租户电流（5G）", liantong_jiliang5g, "分路计量设备-联通租户电流（5G）.xls"),
    ("分路计量设备-电信租户电流", dianxin_jiliang, "分路计量设备-电信租户电流.xls"),
    ("开关电源-电信租户电流", dianxin_kaiguan, "开关电源-电信租户电流.xls"),
    ("分路计量设备-电信租户电流（5G）", dianxin_jiliang5g, "分路计量设备-电信租户电流（5G）.xls"),
]


def _clean_cookie_dict(raw_cookie_dict):
    cleaned = {}
    for key, value in (raw_cookie_dict or {}).items():
        k = str(key).strip()
        v = str(value).strip().replace("\r", "").replace("\n", "")
        if k:
            cleaned[k] = v
    return cleaned


def get_cookie_dict(cookie_id=DEFAULT_COOKIE_ID):
    """从 MySQL 中读取并清洗 Cookie 字典。"""
    db = sql_orm()
    cookie_result = db.get_cookies(cookie_id)
    return _clean_cookie_dict(cookie_result.get("cookies", {}))


def build_default_headers(cookie_dict):
    cookie_str = "; ".join(f"{k}={v}" for k, v in cookie_dict.items())
    return {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "Cookie": cookie_str,
        "Host": "omms.chinatowercom.cn:9000",
        "Origin": BASE_URL,
        "Referer": f"{BASE_URL}/devMge/initShuntMetering.go",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0",
        "X-Requested-With": "XMLHttpRequest",
    }


def create_session(cookie_id=DEFAULT_COOKIE_ID):
    cookie_dict = get_cookie_dict(cookie_id)
    headers = build_default_headers(cookie_dict)

    session = requests.Session()
    session.headers.update(headers)
    session.cookies.update(cookie_dict)
    return session


def retry(max_attempts=15, delay=2):
    def decorator_retry(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            attempts = 0
            while attempts < max_attempts:
                try:
                    return func(*args, **kwargs)
                except Exception as exc:
                    attempts += 1
                    if attempts >= max_attempts:
                        raise
                    print(f"重试 {attempts}/{max_attempts}: {str(exc)[:80]}")
                    time.sleep(delay)

        return wrapper

    return decorator_retry


@retry()
def requests_post(url, session=None, headers=None, data=None, timeout=600):
    req_session = session or create_session()
    return req_session.post(url, headers=headers, data=data or {}, timeout=timeout)


@retry()
def requests_get(url, session=None, headers=None, params=None, timeout=600, stream=False):
    req_session = session or create_session()
    return req_session.get(url, headers=headers, params=params or {}, timeout=timeout, stream=stream)


def _prepare_lessee_payload(payload, access_status="02"):
    """避免污染 schema 全局字典，每次下载前做深拷贝并设置 queryAccessStatus。"""
    payload_copy = copy.deepcopy(payload)
    for step_name in ("INTO_DATA_1", "INTO_DATA_2", "INTO_DATA_FINAL"):
        if step_name in payload_copy:
            payload_copy[step_name]["queryForm:queryAccessStatus"] = access_status
    return payload_copy


def down_file(url, data, path, conten_len_error=3000, xlsx_juge=False, session=None):
    """
    通用文件下载器：用于租户电流文件导出。
    保持原函数名与参数兼容，便于现有调用不改动。
    """
    retry_count = 3
    req_session = session or create_session()
    request_headers = {
        "Host": "omms.chinatowercom.cn:9000",
        "Origin": BASE_URL,
        "Referer": url,
        "User-Agent": req_session.headers.get("User-Agent", ""),
    }

    while retry_count >= 0:
        try:
            first_resp = requests_post(url, session=req_session, headers=request_headers)
            first_resp.raise_for_status()

            html = BeautifulSoup(first_resp.text, "html.parser")
            view_state_node = html.find("input", id="javax.faces.ViewState")
            if not view_state_node or not view_state_node.get("value"):
                raise ValueError("页面中未找到 javax.faces.ViewState")

            javax_value = view_state_node["value"]

            for step_name, step_data in data.items():
                payload = dict(step_data)
                payload["javax.faces.ViewState"] = javax_value
                step_resp = requests_post(url, session=req_session, headers=request_headers, data=payload)
                step_resp.raise_for_status()

                if "FINAL" in step_name:
                    if len(step_resp.content) < conten_len_error:
                        raise ValueError("导出内容过小，疑似下载失败")
                    if xlsx_juge and not step_resp.content.startswith(
                        (b"\x50\x4B\x03\x04", b"\x09\x08\x04\x00\x10\x00\x00\x00")
                    ):
                        raise ValueError("返回内容不是 Excel 文件")

                    os.makedirs(os.path.dirname(path), exist_ok=True)
                    with open(path, "wb") as file_obj:
                        file_obj.write(step_resp.content)
            return
        except Exception as exc:
            retry_count -= 1
            if retry_count < 0:
                raise
            print(f"重试中... {exc}")


def download_lessee_meter_excels(cookie_id=DEFAULT_COOKIE_ID, session=None):
    """下载 9 个分路计量租户电流文件。"""
    session = session or create_session(cookie_id=cookie_id)
    success_count = 0

    for task_name, task_payload, file_name in LESSEE_DOWNLOAD_TASKS:
        save_path = settings.resolve_path(f"app/service/jiliangzhibiao/down/{file_name}")
        payload = _prepare_lessee_payload(task_payload, access_status="02")
        down_file(url=PERFORMANCE_URL, data=payload, path=save_path, session=session)
        success_count += 1

    print(f"租户电流下载完成：{success_count}/9")


def download_shunt_meter_excel(cookie_id=DEFAULT_COOKIE_ID, session=None):
    """
    下载分路计量设备清单（双请求模式）。
    返回保存路径。
    """
    save_path = settings.resolve_path(r"app/service/jiliangzhibiao/down/分路计量设备清单.xls")
    os.makedirs(os.path.dirname(save_path), exist_ok=True)

    req_session = session or create_session(cookie_id=cookie_id)
    query_payload = {
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
        "rows": "20",
    }

    query_resp = requests_post(SHUNT_QUERY_URL, session=req_session, data=query_payload, timeout=30)
    query_resp.raise_for_status()

    export_resp = requests_get(SHUNT_EXPORT_URL, session=req_session, stream=True, timeout=60)
    export_resp.raise_for_status()

    with open(save_path, "wb") as file_obj:
        for chunk in export_resp.iter_content(1024 * 1024):
            file_obj.write(chunk)

    print(f"分路计量设备清单下载完成")
    return save_path


def run_all_downloads(cookie_id=DEFAULT_COOKIE_ID):
    """一次运行同时下载：租户电流(9个) + 分路计量设备清单(1个)。"""
    session = create_session(cookie_id=cookie_id)
    download_lessee_meter_excels(cookie_id=cookie_id, session=session)
    download_shunt_meter_excel(cookie_id=cookie_id, session=session)
    print("\n全部下载任务执行完成")


if __name__ == "__main__":
    try:
        run_all_downloads()
    except Exception as exc:
        print(f"\n❌ 执行失败：{exc}")
