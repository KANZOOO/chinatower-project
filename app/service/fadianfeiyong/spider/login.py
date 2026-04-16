from requests import Session
import datetime
import os
import pickle
from core.utils.yzm_orgnize import main as yzm_orgnize
from core.config import settings

class quxin_login():
    def __init__(self,quxin_session=Session()):
        self.pickle_quxin =settings.resolve_path("spider/down/pickle_quxin.pkl")
        self.session_quxin=quxin_session
        self.session_quxin.headers.update({
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Host": "10.8.3.2:80",
            "Pragma": "no-cache",
            "Referer": "http://10.8.3.2:80/tower_manage_bms/a?login",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
        })
        self.picture_path=settings.resolve_path(r"spider/down/yzm.jfif")
    def auto_login(self):
        self.get_yzm()
        yzm=yzm_orgnize(self.picture_path)
        self.get_cookie(yzm)
    def get_yzm(self):
        res=self.session_quxin.get(url='http://10.8.3.2:80/tower_manage_bms/a/login')
        res=self.session_quxin.post(url='http://10.8.3.2:80/tower_manage_bms/servlet/validateCodeServlet')
        with open(self.picture_path, "wb") as codes:
         codes.write(res.content)
        return self.session_quxin
    def get_cookie(self,validateCode):
        try:
            data={'password': 'Na1FhaIp8cw78O8LRNtCahqEcNXHjgW1XmlnmogFYuW/VvRlbC6epLwFq81G2k3K6gYpQ/St9hR+jdwe+6MbBdrlg9yLM5jeAMyxXPR4ultDkksJWPne+UCvyTqWKhlfcqQZF+BGLhqIVy0MJLFH2/ATOdNNtsqD81eQx5Un1ZE=',
             'username': 'nanning_oil',
             'validateCode': validateCode}
            res=self.session_quxin.post(url='http://10.8.3.2:80/tower_manage_bms/a/login',data=data)
            with open(self.pickle_quxin, 'wb') as f:
                pickle.dump(self.session_quxin, f)
        except Exception as e:
            return '登陆取信失败'+str(e)
def main():
    quxin_login().auto_login()
if __name__ == '__main__':
    main()