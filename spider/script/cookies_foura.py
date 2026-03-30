import datetime
from playwright.sync_api import sync_playwright
from core.sql import sql_orm
import requests
LOGIN_URL = 'http://4a.chinatowercom.cn:20000/uac/'
YZM_URL = 'http://clound.gxtower.cn:3980//tt/get_yzm'
ACCOUNTS = [
    {'username': 'wx-yeping6', 'password': '3Qw@321654', 'cookie_id': 'wx-yeping6'},
]


class KeepFourA:
    def __init__(self, account_index=0):
        if not 0 <= account_index < len(ACCOUNTS):
            raise ValueError(f"account_index 必须是 0 到 {len(ACCOUNTS) - 1}")
        self.account = ACCOUNTS[account_index]
        self.cookie_id = self.account['cookie_id']
        self.target_url = "http://omms.chinatowercom.cn:9000"

    def _setup_browser(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=False,
            args=['--start-maximized'],
        )
        self.context = self.browser.new_context()

    def _close_browser(self):
        if hasattr(self, 'context') and self.context:
            self.context.close()
        if hasattr(self, 'playwright') and self.playwright:
            self.playwright.stop()

    def _login(self):
        page = self.context.new_page()
        page.goto(LOGIN_URL)
        page.locator('#username').fill(self.account['username'])
        page.locator('#password').fill(self.account['password'])
        page.wait_for_timeout(2000)
        page.locator('.login_btn').click()
        page.wait_for_timeout(2000)
        page.evaluate("refreshMsg();")
        page.wait_for_timeout(15000)
        success = False
        for attempt in range(5):
            try:
                yzm = requests.get(YZM_URL).text.strip()
                page.locator('#msgCode').fill(yzm)
                page.evaluate("doLogin();")
                page.wait_for_timeout(5000)
                page.evaluate("gores('200007','20988','');")
                success = True
                break
            except Exception as e:
                print(f"第{attempt + 1}次尝试gores失败: {e}")
                if attempt < 4:
                    page.wait_for_timeout(5000)
        if not success:
            raise RuntimeError("获取验证码并执行gores失败，已达最大重试次数")
        return page

    def _enter_yunjian(self, page):
        page.evaluate("gores('200007','20988','');")
        page.wait_for_timeout(10000)

    def _enter_zi_ding_yi_bao_biao(self):
        page = self.context.pages[-1]
        page.goto(
            'http://omms.chinatowercom.cn:9000/business/zndw/sso/index.jsp?menuPath=modules/selfTask/views/taskListIndex&url=http://omms.chinatowercom.cn:9000/portal')
        page.wait_for_timeout(10000)

    def _save_cookie(self):
        cookies = self.context.cookies(self.target_url)
        cookie_str = '; '.join(f"{cookie['name']}={cookie['value']}" for cookie in cookies)
        now = datetime.datetime.now()
        with sql_orm().session_scope() as (sqlsession, Base):
            pojo = Base.classes.cookies
            current_cookie = sqlsession.query(pojo).filter_by(id=self.cookie_id).first()
            if current_cookie:
                current_cookie.cookies = cookie_str
                current_cookie.LastLoginTime = now

    def get_cookies(self):
        self._setup_browser()
        try:
            self.context.clear_cookies()
            begin_page = self._login()
            for _ in range(3):
                try:
                    begin_page.reload()
                    self._enter_yunjian(begin_page)
                    self._enter_zi_ding_yi_bao_biao()
                    self._save_cookie()
                    break  # 成功就跳出
                except Exception as e:
                    print(f"未获取到cookie，重新进入运监: {e}")
        finally:
            self._close_browser()


def main():
    for index in range(len(ACCOUNTS)):
        keep = KeepFourA(account_index=index)
        keep.get_cookies()


if __name__ == '__main__':
    main()