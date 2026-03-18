# spider/script/cookies_keep_energy.py
import time
import json
from seleniumwire import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import requests
from core.sql import sql_orm
from core.msg.msg_ding import DingMsg
from core.utils.yzm_queue import yzm_queue
from core.config import settings
class KeepEnergy:
    def __init__(self):
        # 常量定义
        self.LOGIN_URL = 'https://energy-iot.chinatowercom.cn/login'
        self.GET_INFO_URL = 'https://energy-iot.chinatowercom.cn/api/admin/base/getInfo'
        self.WORKORDER_PAGE_URL = 'https://energy-iot.chinatowercom.cn/api/workorder/workOrder/page'
        # self.USERNAME = 'nyqujkzx'
        # self.PASSWORD = 'T!.t2021m23T'
        self.USERNAME = 'gxnyjkdxcx'
        self.PASSWORD = '15177191882aA`'
        self.CHROMEDRIVER_PATH = settings.index/ 'chromedriver.exe'
        self.GET_AIOT_RUN_TIMES_FILE = settings.resolve_path("spider\down\yzm_queue.pkl")


    def _setup_driver(self):
        """配置并返回Chrome浏览器驱动实例"""
        service = Service(executable_path=self.CHROMEDRIVER_PATH)
        option = webdriver.ChromeOptions()
        # option.add_argument('--headless')  # 如果需要无头模式，取消注释
        option.add_argument("--disable-gpu")
        option.add_argument("--no-sandbox")
        option.binary_location = settings.chrome_binary_path
        option.add_argument(f'--user-data-dir={settings.chrome_user_data_dir}')
        return webdriver.Chrome(service=service, options=option)
    def _login_and_get_authorization(self,driver):
        """执行登录操作并获取Authorization信息"""
        driver.get(self.LOGIN_URL)
        try:
            # 输入账号密码
            input_eles = driver.find_elements(By.CLASS_NAME, 'el-input__inner')
            self.USERNAME_input, self.PASSWORD_input = input_eles[0], input_eles[1]
            self.USERNAME_input.send_keys(self.USERNAME)
            self.PASSWORD_input.send_keys(self.PASSWORD)

            # 点击登录
            login_div = driver.find_element(By.CLASS_NAME, 'login-warpper')
            login_button = login_div.find_element(By.XPATH, '//form[1]/div[3]/div[1]/button[1]/span[1]')
            driver.execute_script("arguments[0].click();", login_button)
            time.sleep(2)

            # 点击“获取验证码”
            verify_div = driver.find_elements(By.CLASS_NAME, 'verify-code-input')[1]
            get_code_button = verify_div.find_element(By.XPATH, './/button[1]/span[1]')
            driver.execute_script("arguments[0].click();", get_code_button)
            time.sleep(5)  # 等待验证码到达

            # 获取验证码输入框和确认按钮
            code_input = verify_div.find_element(By.XPATH, './/div[1]/div[1]/input[1]')
            confirm_button = driver.find_elements(By.CLASS_NAME, 'el-form-item.el-form-item--medium')[3] \
                .find_element(By.XPATH, './/div[1]/button[2]')
            time.sleep(60)
            # 从验证码队列中倒序尝试（最新验证码优先）
            yzm_list = yzm_queue().get_queue()
            for yzm in reversed(yzm_list):
                try:
                    code_input.clear()
                    code_input.send_keys(str(yzm).strip())
                    time.sleep(5)
                    driver.execute_script("arguments[0].click();", confirm_button)
                    time.sleep(5)
                    break  # 尝试一次即可，无论成功与否（后续靠 token 检测）
                except Exception as e:
                    print(f"输入验证码 {yzm} 时出错: {e}")
                    continue

        except Exception as e:
            print("登录流程异常:", e)
            pass

        return driver.requests
    def get_cookies(self):
        """执行AIOT登录并获取Authorization信息，保存到数据库"""
        driver = None
        try:
            driver = self._setup_driver()
            res_list = self._login_and_get_authorization(driver)

            with sql_orm().session_scope() as temp:
                session, Base = temp
                cookies = Base.classes.cookies
                for res in res_list:
                    if self.GET_INFO_URL in res.url:
                        for header in res.headers._headers:
                            if 'Authorization' in header:
                                Authorization = header[1]
                                res = session.query(cookies).filter(cookies.id == "energy").first()
                                if res:
                                    res.cookies = Authorization
                                    print('AIOT 能源网管登录成功')
                                else:
                                    raise
                                break
                        break
        except Exception as e:
            print(e)
            raise
        finally:
            if driver:
                driver.quit()
    def keep_cookies(self):
        """检查AIOT登录状态，若失效则重新登录"""
        try:
            Authorization=sql_orm().get_cookies("energy")["cookies_str"]

            headers = {
                'Content-Type': 'application/json;charset=UTF-8',
                'Authorization': Authorization
            }
            data = {
                "searchType": "1", "alarmClearTimes": [], "pageNum": 1,
                "pageSize": 10, "deptIds": [], "workType": "0",
                "timer": [], "businessType": "", "createTimes": [], "receiptTimes": []
            }

            res = requests.post(self.WORKORDER_PAGE_URL, headers=headers, data=json.dumps(data))
            if res.status_code != 200 or '暂未登录或token已经过期' in res.text:
                d = DingMsg()
                d.text_at(d.TEST, 'AIOT-能源网管登录失效')
                print('AIOT-能源网管登录失效')

        except Exception as e:
            raise

def main_keep_cookies():
    KeepEnergy().keep_cookies()
def main_get_cookies():
    KeepEnergy().get_cookies()


