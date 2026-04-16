import os
import pickle
from datetime import datetime, timedelta
from core.config import settings
from app.service.fadianfeiyong.spider import schema


def power_station_report():
    pickle_quxin = settings.resolve_path("spider/down/pickle_quxin.pkl")
    with open(pickle_quxin, "rb") as f:
        session_quxin = pickle.load(f)

    data = {
        "approvalOfDispatchId": "",
        "area.id": schema.GENERATE_AREA_IDS,
        "area.name": "",
        "asOper": "",
        "beginGenerateDate": begin_time.strftime("%Y-%m-%d %H:%M:%S"),
        "city.id": "",
        "city.name": "",
        "collectorCode": "",
        "endGenerateDate": end_time.strftime("%Y-%m-%d %H:%M:%S"),
        "finishConfigId": "",
        "generateOfficeName": "",
        "generatePowerState": "1",
        "isStart": "",
        "number": "",
        "orderBy": "",
        "pageNo": "1",
        "pageSize": "25",
        "stationCode": "",
        "stationName": "",
        "stopBegeinDate": "",
        "stopEndDate": "",
        "typeCode": "",
        "urlPara": "",
    }
    res = session_quxin.post(url=schema.QUXIN_EXPORT_TOWER_REPORT_URL, data=data)
    with open(self.gen_path, "wb") as f:
        f.write(res.content)