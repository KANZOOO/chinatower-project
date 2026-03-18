# spider/script/cookies_keep_foura.py
import time
import json
import datetime
import threading
import requests
from contextlib import contextmanager
from playwright.sync_api import sync_playwright, Page
from bs4 import BeautifulSoup
from core.sql import sql_orm
from core.utils.retry_wrapper import requests_get, requests_post
from spider.script.keepalive import keepalive_config as CONFIG


class KeepFourA:
    def __init__(self):
        # 常量定义 - 保留所有 FourA 参数
        self.LOGIN_URL = 'http://4a.chinatowercom.cn:20000/uac/'
        self.KEEP_FA_URL = CONFIG.KEEP_FA_URL
        self.KEEP_FA_URL2 = CONFIG.KEEP_FA_URL2
        self.USERNAME_LIST = {
            0: {'username': 'wx-xiepc', 'password': 'Xpc@12345'},
            1: {'username': 'wx-yangsa', 'password': 'Shan!836556438'},
            2: {'username': 'wx-zhouyun9', 'password': '#zY817703.'}
        }
        self.YZM_URL = 'http://clound.gxtower.cn:3980/tt/get_yzm'

    def _setup_browser(self):
        """配置并返回 Playwright 浏览器实例"""
        self.playwright = sync_playwright().start()
        playwright_user_data_dir = r'D:\tower\playwright_user_data'

        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=playwright_user_data_dir,
            headless=False,
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

    def _open_4a(self, page: Page, cookie_user: int = 0):
        """打开4A登录页面"""
        page.goto(self.LOGIN_URL)
        page.wait_for_timeout(5000)

    def _login_menhu(self, page: Page, cookie_user: int = 0):
        """登录门户"""
        user_info = self.USERNAME_LIST.get(cookie_user, self.USERNAME_LIST[0])
        username = user_info['username']
        password = user_info['password']

        # 1. 输入用户名密码
        page.locator('#username').fill(username)
        page.locator('#password').fill(password)

        # 2. 等待10秒
        page.wait_for_timeout(10000)

        # 3. 点击登录按钮
        page.locator('.login_btn').click()
        page.wait_for_timeout(2000)

        # 4. 执行 refreshMsg()（触发验证码发送）
        page.evaluate("refreshMsg();")

        try:
            # 5. 等待验证码输入框出现
            yzm_input = page.locator('#msgCode')
            yzm_input.wait_for(timeout=5000)

            # 6. 等待20秒，让短信发送
            page.wait_for_timeout(20000)

            # 7. 从接口获取验证码
            yzm = requests.get(self.YZM_URL).text

            # 8. 输入验证码并登录
            yzm_input.fill(yzm)
            page.wait_for_timeout(2000)
            page.evaluate("doLogin();")
            page.wait_for_timeout(5000)

        except Exception as e:
            print(f"未检测到验证码输入框，跳过验证码步骤: {e}")

    def _login_yunjian(self, page: Page):
        """登录云鉴（在门户登录成功后调用）"""
        page.evaluate("gores('200007','20988','');")
        page.wait_for_timeout(10000)

    def _login_menhu_and_login_yunjian(self, page: Page, cookie_user: int = 0):
        """组合登录门户和云鉴，返回 cookie"""
        self._login_menhu(page, cookie_user)
        page.wait_for_timeout(3000)

        self._login_yunjian(page)
        page.wait_for_timeout(10000)
        target_cookies = self.context.cookies('http://omms.chinatowercom.cn:9000')
        cookie_str = '; '.join([f"{c['name']}={c['value']}" for c in target_cookies])
        return {'cookie': cookie_str}

    def _refresh_fail(self, page: Page):
        """刷新页面"""
        page.reload()

    def _login_and_get_cookies(self, page: Page, cookie_user: int = 0):
        """执行登录流程并获取 cookies"""
        j = 0
        result = None

        while j < 10:
            j += 1
            flag = 0

            current_page = self.context.pages[-1] if self.context.pages else page
            current_url = current_page.url

            # 如果是空白页或不在登录页面，先打开登录页面
            if current_url == 'about:blank' or not current_url.startswith('http'):
                func = self._open_4a
                func_fail = self._refresh_fail
            elif 'http://4a.chinatowercom.cn:20000/uac/' in current_url:
                func = self._login_menhu_and_login_yunjian
                func_fail = self._refresh_fail
            else:
                # 其他已登录页面，直接获取cookie
                try:
                    cookies = self.context.cookies()
                    cookie_str = '; '.join([f"{c['name']}={c['value']}" for c in cookies])
                    # 检查cookie是否有效（不为空）
                    if cookie_str and len(cookie_str) > 50:  # 简单检查长度
                        return {'cookie': cookie_str}
                    else:
                        print("获取到的cookie太短，可能无效，重新登录...")
                        func = self._open_4a
                        func_fail = self._refresh_fail
                except Exception as e:
                    print(f"获取cookie失败: {e}")
                    func = self._open_4a
                    func_fail = self._refresh_fail

            try:
                result = func(current_page, cookie_user)
                flag = 1
                if result is not None and result.get('cookie'):
                    # 验证cookie有效性
                    if len(result['cookie']) > 50:
                        break
                    else:
                        print("获取的cookie太短，继续重试...")
                        flag = 0
            except Exception as e:
                print(f"登录尝试失败: {e}")

            if flag == 0:
                try:
                    func_fail(current_page)
                except:
                    pass

        return result

    def get_cookies(self, cookie_user: int = 0):
        """执行 FourA 登录并获取 Cookie 信息，保存到数据库"""
        db_id = cookie_user + 1

        try:
            page = self._setup_browser()
            self.context.clear_cookies()

            result = self._login_and_get_cookies(page, cookie_user)

            if result:
                with sql_orm().session_scope() as temp:
                    sqlsession, Base = temp
                    pojo = Base.classes.cookies

                    # 直接按 id 查询，更可靠
                    fa = sqlsession.query(pojo).filter_by(id=db_id).first()
                    if fa:
                        fa.cookies = result.get('cookie')
                        fa.LastLoginTime = datetime.datetime.now()
                        print(f'4A账号{cookie_user}(数据库id={db_id})登录成功')
                    else:
                        print(f"数据库中找不到id={db_id}的记录")
            else:
                print(f'4A账号{cookie_user}登录失败，未获取到cookie')

        except Exception as e:
            print(f"获取 FourA Cookie 失败: {e}")
            raise
        finally:
            self._close_browser()

    def get_all_cookies(self):
        """依次获取所有三个账号的 cookies 并保存到数据库"""
        for cookie_user in range(3):  # 0, 1, 2 对应三个账号
            db_id = cookie_user + 1  # 数据库 id 为 1, 2, 3
            print(f"\n开始处理第 {cookie_user + 1} 个账号 (数据库id={db_id})...")

            try:
                self.get_cookies(cookie_user)
                # 每个账号之间间隔5秒，避免过于频繁
                if cookie_user < 2:  # 最后一个账号不需要等待
                    time.sleep(5)
            except Exception as e:
                print(f"处理账号 {cookie_user} 时出错: {e}")
                continue  # 继续处理下一个账号

        print("\n所有账号处理完成")

    def keep_single_cookie(self, cookie_user: int = 0):
        """检查单个 FourA 账号的登录状态，若失效则重新登录"""
        db_id = cookie_user + 1

        try:
            with sql_orm().session_scope() as temp:
                sqlsession, Base = temp
                pojo = Base.classes.cookies

                # 按 id 查询特定账号
                fa = sqlsession.query(pojo).filter_by(id=db_id).first()

                if not fa:
                    print(f'数据库中找不到id={db_id}的记录，尝试重新登录...')
                    self.get_cookies(cookie_user)
                    return

                cookies_str = fa.cookies
                headers = CONFIG.KEEP_FA_HEADERS.copy()
                headers['Cookie'] = cookies_str
                headers['Referer'] = CONFIG.KEEP_FA_URL2

                res = requests_post(CONFIG.KEEP_FA_URL2, headers=headers, timeout=20)
                html = BeautifulSoup(res.text, 'html.parser', from_encoding='utf-8')
                javax = html.find('input', id='javax.faces.ViewState')['value']

                print(
                    f'4A账号{cookie_user}(数据库id={db_id})保活成功 {datetime.datetime.now().strftime("%Y%m%d %H%M")}')

        except Exception as e:
            print(
                f'4A账号{cookie_user}(数据库id={db_id})保活不成功 {datetime.datetime.now().strftime("%Y%m%d %H%M")}，尝试重新登录...')
            try:
                self.get_cookies(cookie_user)
            except Exception as e2:
                print(f"重新登录失败: {e2}")

    def keep_cookies(self):
        """检查所有 FourA 账号的登录状态，若失效则重新登录"""
        print(f"\n开始保活检查 - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

        for cookie_user in range(3):  # 0, 1, 2 对应三个账号
            try:
                self.keep_single_cookie(cookie_user)
                # 账号之间间隔3秒，避免请求过于频繁
                if cookie_user < 2:
                    time.sleep(3)
            except Exception as e:
                print(f"保活账号 {cookie_user} 时出错: {e}")
                continue

        print(f"保活检查完成 - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")


def run_thread(func):
    """在线程中运行函数"""
    thread = threading.Thread(target=func)
    thread.start()


def run_task_min():
    """每分钟执行的任务"""
    keep_foura = KeepFourA()
    run_thread(keep_foura.keep_cookies)


def main_keep_cookies():
    """主函数：保活所有账号"""
    KeepFourA().keep_cookies()


def main_get_cookies():
    """主函数：获取所有账号的 cookie"""
    KeepFourA().get_all_cookies()