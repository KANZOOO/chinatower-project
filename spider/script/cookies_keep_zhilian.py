# spider/script/keepalive/cookies_keep_aiot.py
import time
import json
import datetime
import threading
import requests
from playwright.sync_api import sync_playwright, Page
from bs4 import BeautifulSoup
from core.sql import sql_orm
from core.utils.retry_wrapper import requests_get, requests_post
from spider.script.keepalive import keepalive_config as CONFIG


class KeepZhiLian:
    def __init__(self):
        # 常量定义 - AIOT 智联参数
        self.LOGIN_URL = 'https://zlzywg.chinatowercom.cn:8070/login'
        self.API_BASE_URL = 'https://zlzywg.chinatowercom.cn:8070'
        self.DEPT_TREE_URL = f'{self.API_BASE_URL}/api/admin/base/system/dept/treeselect'
        self.SEMP_DEV_LIST_URL = f'{self.API_BASE_URL}/api/semp/sempDev/listSempDev'

        self.USERNAME = 'qujkzxcx'
        self.PASSWORD = 'JKzx@12345'

        # 验证码相关配置（如果需要）
        self.YZM_PATH = r'E:\playwright_user_data\yzm-aiot.gif'
        self.YZM_API_URL = 'http://clound.gxtower.cn:3980/tt/get_yzm'  # 如果需要外部验证码服务

    def _setup_browser(self):
        """配置并返回 Playwright 浏览器实例"""
        self.playwright = sync_playwright().start()
        playwright_user_data_dir = r'E:\playwright_user_data\aiot'

        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=playwright_user_data_dir,
            headless=False,  # AIOT 可能需要图形界面处理验证码
            args=['--start-maximized']
        )

        self.page = self.context.new_page()
        self.page.set_default_timeout(30000)
        return self.page

    def _close_browser(self):
        """关闭浏览器"""
        try:
            if self.context:
                self.context.close()
            if self.playwright:
                self.playwright.stop()
        except Exception as e:
            print(f"关闭浏览器失败: {e}")

    def _open_login_page(self, page: Page):
        """打开 AIOT 登录页面"""
        page.goto(self.LOGIN_URL)
        page.wait_for_timeout(3000)

    def _handle_captcha(self, page: Page):
        """处理验证码（如果需要）"""
        try:
            # 等待验证码图片加载
            captcha_img = page.locator('img.el-image__inner').first
            captcha_img.wait_for(timeout=5000)

            # 截图保存验证码
            captcha_img.screenshot(path=self.YZM_PATH)
            print(f"验证码已保存到: {self.YZM_PATH}")

            # 这里可以集成验证码识别服务
            # 暂时返回空，需要手动输入或外部服务
            return None

        except Exception as e:
            print(f"验证码处理跳过: {e}")
            return None

    def _login(self, page: Page):
        """执行登录操作"""
        # 1. 输入用户名
        page.locator('input.el-input__inner').nth(0).fill(self.USERNAME)

        # 2. 输入密码
        page.locator('input.el-input__inner').nth(1).fill(self.PASSWORD)

        # 3. 处理验证码（如果有）
        captcha_code = self._handle_captcha(page)
        if captcha_code:
            page.locator('input.el-input__inner').nth(2).fill(captcha_code)

        # 4. 点击登录按钮
        page.locator('button:has-text("登录")').click()
        page.wait_for_timeout(3000)

        # 5. 处理二次验证（如果需要）
        try:
            # 检查是否有二次验证弹窗
            verify_input = page.locator('input.el-input__inner').nth(3)
            verify_input.wait_for(timeout=5000)

            # 等待手动输入或从队列获取验证码
            print("等待二次验证...")
            time.sleep(60)  # 给60秒时间处理

        except Exception:
            print("无需二次验证或验证已处理")

    def _get_authorization(self, page: Page):
        """从页面请求中提取 Authorization"""
        # 等待 API 请求触发
        page.wait_for_timeout(5000)

        # 监听网络请求获取 Authorization
        # 方案1：通过执行 JS 从 localStorage 或 sessionStorage 获取
        auth = page.evaluate("""() => {
            return localStorage.getItem('Authorization') || 
                   sessionStorage.getItem('Authorization') || 
                   '';
        }""")

        if auth:
            return auth

        # 方案2：尝试从 cookie 中获取
        cookies = self.context.cookies(self.API_BASE_URL)
        for cookie in cookies:
            if cookie['name'] == 'Authorization':
                return cookie['value']

        return None

    def _login_and_get_auth(self, page: Page):
        """完整登录流程并获取认证信息"""
        max_retries = 3
        auth = None

        for i in range(max_retries):
            try:
                self._open_login_page(page)
                self._login(page)
                auth = self._get_authorization(page)

                if auth and len(auth) > 20:  # 简单验证 token 长度
                    print(f"登录成功，获取到 Authorization")
                    break
                else:
                    print(f"第 {i + 1} 次尝试未获取到有效 Authorization，重试...")
                    page.reload()

            except Exception as e:
                print(f"第 {i + 1} 次登录尝试失败: {e}")
                if i < max_retries - 1:
                    page.reload()
                    time.sleep(2)

        return {'authorization': auth} if auth else None

    def get_auth(self):
        """执行 AIOT 登录并获取 Authorization，保存到数据库"""
        db_id = 3  # AIOT 固定使用 id=3

        try:
            page = self._setup_browser()
            self.context.clear_cookies()

            result = self._login_and_get_auth(page)

            if result and result.get('authorization'):
                with sql_orm().session_scope() as temp:
                    sqlsession, Base = temp
                    pojo = Base.classes.aiot

                    record = sqlsession.query(pojo).filter_by(id=db_id).first()
                    if record:
                        record.Authorization = result.get('authorization')
                        record.LastLoginTime = datetime.datetime.now()
                        print(f'AIOT(数据库id={db_id})登录成功，Authorization已更新')
                    else:
                        print(f"数据库中找不到id={db_id}的记录")
                        # 可选：自动创建记录
                        # new_record = pojo(id=db_id, Authorization=result.get('authorization'), LastLoginTime=datetime.datetime.now())
                        # sqlsession.add(new_record)
            else:
                print(f'AIOT登录失败，未获取到有效Authorization')

        except Exception as e:
            print(f"获取 AIOT Authorization 失败: {e}")
            raise
        finally:
            self._close_browser()

    def keep_auth(self):
        """检查 AIOT 登录状态，若失效则重新登录"""
        db_id = 3

        try:
            with sql_orm().session_scope() as temp:
                sqlsession, Base = temp
                pojo = Base.classes.aiot
                record = sqlsession.query(pojo).filter_by(id=db_id).first()

                if not record or not record.Authorization:
                    print(f'AIOT数据库记录不存在或Authorization为空，需要重新登录')
                    self.get_auth()
                    return

                auth = record.Authorization

                # 测试 API 请求验证有效性
                headers = {
                    'Content-Type': 'application/json;charset=UTF-8',
                    'Authorization': auth
                }
                data = {
                    "searchType": "1",
                    "alarmClearTimes": [],
                    "pageNum": 1,
                    "pageSize": 10,
                    "deptIds": [],
                    "workType": "0",
                    "timer": [],
                    "businessType": "",
                    "createTimes": [],
                    "receiptTimes": []
                }

                res = requests_post(
                    self.SEMP_DEV_LIST_URL,
                    headers=headers,
                    json=data,  # 使用 json 参数自动序列化
                    timeout=20
                )

                # 检查响应状态
                if res.status_code == 200:
                    result = res.json()
                    if result.get('code') == 200 or 'data' in result:
                        print(f'AIOT保活成功 {datetime.datetime.now().strftime("%Y%m%d %H%M")}')
                    else:
                        raise Exception(f"API返回异常: {result.get('msg', '未知错误')}")
                else:
                    raise Exception(f"HTTP状态码异常: {res.status_code}")

        except Exception as e:
            print(f'AIOT保活不成功 {datetime.datetime.now().strftime("%Y%m%d %H%M")}，原因: {e}')
            try:
                self.get_auth()
            except Exception as e2:
                print(f"AIOT重新登录失败: {e2}")
                # 发送告警通知
                try:
                    from utils.send_ding_msg import dingmsg
                    d = dingmsg()
                    d.text_at(d.TEST, f'AIOT-智联登录失效且重登失败: {e2}')
                except Exception:
                    pass


def run_thread(func):
    """在线程中运行函数"""
    thread = threading.Thread(target=func)
    thread.start()


def run_task_min():
    """每分钟执行的任务（如果需要）"""
    keep_aiot = KeepZhiLian()
    run_thread(keep_aiot.keep_auth)


def main_keep_auth():
    """主函数：保活"""
    KeepZhiLian().keep_auth()


def main_get_auth():
    """主函数：获取 authorization"""
    KeepZhiLian().get_auth()