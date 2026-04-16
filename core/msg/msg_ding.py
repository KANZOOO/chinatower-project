import requests
import json
from core.config import settings

class DingMsg():
    def __init__(self):
        # 测试群
        self.TEST = '6b21a8bbca1cb432ccfbb3dcf2b4b7537cb69fd894123ec8638d22f30476ed12'
        # 两翼工单群
        self.LYGD = '3837f4dc1cb2ed6917c0633ba141dc3669d04ec955eaed36201583cc74deedc5'
        # 广西铁塔能源维护业务管理群
        self.ENERGY_MAINTAIN='5db362205aab862d22f33d6a492d05ce1b0682f2ac0e76028e8e453f6be31162'
        # 地震通知群
        self.DZ='c3719b1e39b0fad3c93594275aa2c005e1e8146e593470e6ffe944b9ff650dba'
        # 地震工单群
        self.DZGD = '57140bf1e5796751338c7ecf482ad05924304b39d14f8f3a628c78358ee09627'
        # 退服停电查询群
        self.TF='e801be375ec14c92cc0fd1013bd034335377905300a2617b4b02b72271cb4ebb'
        # bug群
        self.BUG = '0285d0ccda073f351ff66fc9fd7595e3c5c41c133f940cedbe0bb770dbdf0079'
        # 天气群
        self.WEATHER='1ace81dfff87afb6ffdf5d0986cd32ae66f22e87410a0975052065704780c615'
        # 大规模停电群
        self.OUTAGE='d053133b97323ead8c6d27b743b5cf807875f85af699ffffb87cc6c3db863a63'
        # 监控值班群
        self.JK='4d7b210bcadcac3aa175f84f444855e643562d7198d95d5f4179452618eacce1'
        # 发电异常群
        self.FDYC='ef9afde807eb4a33c5e5dcec07b419ad8bfd1a391d21ee499951c081b35cb90a'
        # AI助理测试
        self.AI_TEST="5826d1d2c60d1f1067026a6486a080a74149d949609ac306474ebb17bd4f1cd2"
    def text_at(self,webhook ,msg, num=[], men=[]):
        webhook='https://oapi.dingtalk.com/robot/send?access_token='+webhook
        headers = {'Content-Type': 'application/json;charset=utf-8'}
        data = {
            "at": {
                "atMobiles": num,
                "atUserIds": men,
                "isAtAll": False
            },

            "msgtype": "text",
            "text": {
                "content": msg
            }
        }
        r = requests.post(url=webhook, headers=headers, data=json.dumps(data))
        return r.text

    def picture(self,webhook,down_url,text,title):
        webhook='https://oapi.dingtalk.com/robot/send?access_token='+webhook
        headers = {'Content-Type': 'application/json;charset=utf-8'}
        data = {
            "msgtype": "markdown",
            "markdown": {
                "title": title,
                "text": "![](" + down_url + ")\n" + text
            }
        }
        r = requests.post(url=webhook, headers=headers, data=json.dumps(data))
        print(r.text)
        return r.text

if __name__ == '__main__':
    text="【润建股份】<能源设备离线督办>钦州钦州港点位编码054507020000100366，有40个低速充电设备已离线12.05小时，请督办处理，无法及时维修，请在能源专业网管申请退网"
    d=DingMsg()
    d.text_at(d.TEST,text,[15177191882])