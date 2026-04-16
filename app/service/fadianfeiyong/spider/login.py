import json
from Crypto.PublicKey import RSA
from Crypto.Cipher import PKCS1_v1_5
from requests import Session
import pickle
from core.utils.yzm_orgnize import main as yzm_orgnize
from core.config import settings
import base64
import os

class quxin_login():
    def __init__(self,quxin_session=Session()):
        self.username="nanning_oil"
        self.password="Guangxiyj@123!@#"
        self.pickle_quxin =settings.resolve_path("app/service/fadianfeiyong/spider/down/pickle_quxin.pkl")
        self.session_quxin=quxin_session
        self.session_quxin.headers.update({
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Host": "clound.gxtower.cn:11080",
            "Pragma": "no-cache",
            "Referer": "http://clound.gxtower.cn:11080/tower_manage_bms/a?login",
            "Upgrade-Insecure-Requests": "1",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0"
        })
        self.picture_path=settings.resolve_path(r"spider/down/yzm.jfif")
        self.captcha_key=""
        # 确保目录存在
        os.makedirs(os.path.dirname(self.pickle_quxin), exist_ok=True)
        os.makedirs(os.path.dirname(self.picture_path), exist_ok=True)
        print(f"📁 初始化完成")
        print(f"   Session保存路径：{self.pickle_quxin}")
        print(f"   验证码保存路径：{self.picture_path}")

    def auto_login(self):
        print("\n" + "="*60)
        print("🚀 开始自动登录流程")
        print("="*60)
        try:
            print("\n1️⃣  获取验证码...")
            self.get_yzm()
            print("✅ 验证码获取成功")
            
            print("\n2️⃣  识别验证码...")
            yzm=yzm_orgnize(self.picture_path)
            print(f"✅ 验证码识别结果：{yzm}")
            
            print("\n3️⃣  登录并保存Session...")
            self.get_cookie(yzm)
            print("✅ 登录成功，Session已保存")
            
            print("\n" + "="*60)
            print("🎉 登录流程完成！")
            print("="*60)
        except Exception as e:
            print(f"\n❌ 登录失败：{e}")
            import traceback
            traceback.print_exc()
            
    def get_yzm(self):
        print(f"   → 请求登录页面...")
        res = self.session_quxin.get(url='http://clound.gxtower.cn:11080/tower_manage_vue/login', timeout=30)
        print(f"   ✅ 登录页面请求完成，状态码：{res.status_code}")
        
        print(f"   → 请求验证码接口...")
        res = self.session_quxin.get(url='http://clound.gxtower.cn:11080/gapi/auth-server/getCaptchaImg', timeout=30)
        print(f"   ✅ 验证码接口请求完成，状态码：{res.status_code}")
        
        res=json.loads(res.text)
        self.captcha_key = res["data"]["captchaKey"]
        print(f"   ✅ Captcha Key: {self.captcha_key}")
        
        # 2. 解析base64图片并保存
        img_base64 = res["data"]["captchaImg"].split(",")[1]  # 去掉前缀 data:image/png;base64,
        img_data = base64.b64decode(img_base64)
        with open(self.picture_path, "wb") as codes:
            codes.write(img_data)
        print(f"   ✅ 验证码图片已保存：{self.picture_path}")
        
    def get_cookie(self,validateCode):
        try:
            PUBLIC_KEY = """-----BEGIN PUBLIC KEY-----
            MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQC0sMdUdnw+Yv0QVh2uxcnhn7MM5L0xjmhC88GvoNnYUbCIIr1u9gl0xYGe3WZvfGjPgLtssoVJU/O1ooZLWzJ1alqUNFytIpwpeQU6uH45dGU6nWsKcA/7z5bRefIZT8l3JC31rVr+V/GiJ4TBBhgkKDOFIgVYrKTevQhE9sTDGwIDAQAB
            -----END PUBLIC KEY-----"""

            def encrypt_password(password: str) -> str:
                key = RSA.importKey(PUBLIC_KEY)
                cipher = PKCS1_v1_5.new(key)
                encrypted = cipher.encrypt(password.encode('utf-8'))
                return base64.b64encode(encrypted).decode('utf-8')

            data={'captchaCode': validateCode,
             'captchaKey': self.captcha_key,
             'mac': '90a9f65bd3324a54480be69e7d307fd7',
             'password': encrypt_password(self.password),
             'username': self.username}
            
            print(f"   → 发送登录请求...")
            res = self.session_quxin.post(
                url='http://clound.gxtower.cn:11080/tower_manage_bms/a/login',
                json=data,
                timeout=30)
            print(f"   ✅ 登录请求完成，状态码：{res.status_code}")
            print(f"   响应内容：{res.text[:200]}...")
            
            with open(self.pickle_quxin, 'wb') as f:
                pickle.dump(self.session_quxin, f)
            print(f"   ✅ Session已保存到：{self.pickle_quxin}")
        except Exception as e:
            print(f"   ❌ 登录过程出错：{e}")
            import traceback
            traceback.print_exc()
            raise
            
def main():
    print("="*60)
    print("🏃 取信登录脚本启动")
    print("="*60)
    quxin_login().auto_login()
    
if __name__ == '__main__':
    main()