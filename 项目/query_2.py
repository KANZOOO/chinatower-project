# spider/script/cookies_keep_foura.py
import datetime
import requests
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from core.config import settings
from core.sql import sql_orm
# from core.utils.retry_wrapper import requests_post
# from spider.script.keepalive import keepalive_config as CONFIG

LOGIN_URL = 'http://4a.chinatowercom.cn:20000/uac/'
YZM_URL = 'http://clound.gxtower.cn:3980/tt/get_yzm'
ACCOUNTS = [
    {'username': 'wx-yeping6', 'password': '3Qw@321654'},
    #{'username': 'wx-huangwl14', 'password': 'Cgz275694332@@@'},
    # {'username': 'wx-zhangzj102', 'password': '040213zzJ!'},
]

class KeepFourA:
    # 作用：创建实例时选择用哪个账号
    # 比如：index=0 就用第一个账号 wx-yeping6
    def __init__(self, account_index=0):
        if account_index not in (0, 1):
            raise ValueError("account_index 必须是 0 或 1")
        self.account = ACCOUNTS[account_index]
        self.context = None
        self.playwright = None

    def _setup_browser(self):
        # 作用：打开一个 Chromium 浏览器窗口
        self.playwright = sync_playwright().start()
        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=settings.playwright_user_data_dir,
            headless=False,#   - 不是无头模式（能看到浏览器界面）
            args=['--start-maximized']#   - 窗口最大化
        )
        page = self.context.new_page()
        page.set_default_timeout(30000)#   - 设置超时时间 30 秒
        return page# 返回：一个浏览器页面对象

    def _close_browser(self):
        # 作用：用完浏览器后关闭，释放资源
        # 就像看完电视要关机一样
        if self.context:
            self.context.close()
        if self.playwright:
            self.playwright.stop()

    def _get_current_page(self):
        # 作用：拿到当前正在操作的浏览器标签页
        # 返回最后一个打开的页面
        return self.context.pages[-1] if self.context.pages else None

    def _extract_cookies(self, domain=None):
        # 作用：从浏览器中取出登录后的 Cookie
        cookies = self.context.cookies(domain) if domain else self.context.cookies()
        # 处理：把所有cookie拼成一串字符串
        cookie_str = '; '.join(f"{c['name']}={c['value']}" for c in cookies)
        # 检查：如果长度小于 50，说明 cookie 无效，返回空字符串
        return cookie_str if len(cookie_str) > 50 else ""

    def _open_login_page(self, page):
        # 作用：访问 4A 系统的登录页面
        # 等待：打开后等 5 秒，让页面加载完成
        page.goto(LOGIN_URL)
        page.wait_for_timeout(5000)

    def _login_with_sms(self, page):
        # 作用：填写账号密码 + 验证码，完成登录
        #   1. 填用户名、密码
        page.locator('#username').fill(self.account['username'])
        page.locator('#password').fill(self.account['password'])
        page.wait_for_timeout(10000)
        #   2. 点击登录按钮
        page.locator('.login_btn').click()
        page.wait_for_timeout(2000)
        #   3. 触发短信验证码
        page.evaluate("refreshMsg();")
        page.wait_for_timeout(40000)
        #   4. 从外部服务获取验证码
        yzm = requests.get(YZM_URL).text.strip()
        #   5. 填入验证码并提交
        page.locator('#msgCode').fill(yzm)
        page.evaluate("doLogin();")
        page.wait_for_timeout(5000)

    def _login_yunjian(self, page):
        # 作用：4A 登录后跳转到云剑系统
        # 通过 JavaScript 执行跳转操作
        page.evaluate("gores('200007','20988','');")
        page.wait_for_timeout(10000)

    def _refresh_page(self, page):
        # 作用：重新加载当前页面
        # 备用方案：当其他操作失败时使用
        page.reload()

    def _login_and_get_cookie(self):
        """主循环：根据当前页面状态执行对应操作"""
        # 作用：智能判断当前状态，自动完成登录流程
        # 最多尝试 10 次
        for _ in range(10):
            current = self._get_current_page()
            if not current:
                continue
            url = current.url

            # 状态判断 + 对应动作
            #如果页面空白则刷新页面
            if url == 'about:blank' or not url.startswith('http'):
                func, fallback = self._open_login_page, self._refresh_page
            #如果页面正常则登陆并获取cookie
            elif '4a.chinatowercom.cn:20000/uac/' in url:
                try:
                    #登陆4A并跳转到云剑
                    self._login_with_sms(current)
                    self._login_yunjian(current)
                    cookie = self._extract_cookies('http://omms.chinatowercom.cn:9000')
                    if cookie:
                        return {'cookie': cookie}
                except Exception as e:
                    print(f"登录失败: {e}")
                func, fallback = self._open_login_page, self._refresh_page
            else:
                cookie = self._extract_cookies()
                if cookie:
                    return {'cookie': cookie}
                else:
                    print("Cookie无效，重新登录...")
                    func, fallback = self._open_login_page, self._refresh_page

            # 执行主操作
            try:
                func(current)
            except Exception as e:
                print(f"执行操作失败: {e}")
                try:
                    fallback(current)
                except:
                    pass

        return None

    def get_cookies(self):
        # 作用：完整流程：启动浏览器→登录→保存 cookie 到数据库→关闭浏览器
        self._setup_browser()
        try:
            self.context.clear_cookies()
            result = self._login_and_get_cookie()
            if not result:
                raise Exception("未能获取有效 Cookie")

            with sql_orm().session_scope() as (sqlsession, Base):
                pojo = Base.classes.cookies
                #    保存到数据库的 cookies 表（id='foura'的记录）
                fa = sqlsession.query(pojo).filter_by(id="foura").first()
                fa.cookies = result['cookie']
                #    更新最后登录时间
                fa.LastLoginTime = datetime.datetime.now()
        finally:
            self._close_browser()

    # def keep_cookies(self):
    #     cookies_str=sql_orm().get_cookies("foura")["cookies_str"]
    #     headers = CONFIG.KEEP_FA_HEADERS.copy()
    #     headers.update({'Cookie': cookies_str, 'Referer': CONFIG.KEEP_FA_URL2})
    #     try:
    #         res = requests_post(CONFIG.KEEP_FA_URL2, headers=headers, timeout=20)
    #         soup = BeautifulSoup(res.text, 'html.parser', from_encoding='utf-8')utf-8