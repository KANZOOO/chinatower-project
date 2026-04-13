import datetime
import sys
import os

# 添加项目根目录到Python搜索路径
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../../..')))

from playwright.sync_api import sync_playwright
from core.sql import sql_orm
import requests
LOGIN_URL = 'http://4a.chinatowercom.cn:20000/uac/'
YZM_URL = 'https://ai.runjian.com:38114/web/get_yzm/list'
# YZM_URL = 'http://clound.gxtower.cn:3980//tt/get_yzm'
ACCOUNTS = [
     {'username': 'wx-yeping6', 'password': '3Qw@321654', 'cookie_id': 'wx-yeping6'},
     # {'username': 'dw.rj.fengsw', 'password': '1qaz@WSX', 'cookie_id': 'dw.rj.fengsw'},
    #  {'username': 'dw.rj.chengj', 'password': 'Tower1234#', 'cookie_id': 'dw.rj.chengj'},
    # {'username': 'wx-liangyc11', 'password': 'Tower@1234。','cookie_id': 'wx-liangyc11'},
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
        # 新建页面并访问登录地址
        page = self.context.new_page()
        page.goto(LOGIN_URL)

        # 输入账号密码
        page.locator('#username').fill(self.account['username'])
        page.locator('#password').fill(self.account['password'])
        page.wait_for_timeout(2000)  # 等待2秒确保输入完成
        page.locator('.login_btn').click()  # 点击登录按钮
        page.wait_for_timeout(2000)  # 等待登录请求处理
        page.evaluate("refreshMsg();")  # 执行页面刷新验证码的JS函数

        # 关键：等待15秒，预留验证码生成和同步时间（防止验证码被覆盖）
        page.wait_for_timeout(15000)

        # 标记是否登录成功
        login_success = False
        # 标记最大重试次数（防止无限循环）
        max_verify_attempts = 5

        # 第一步：获取验证码列表
        try:
            # 请求验证码接口并解析返回数据
            verify_response = requests.get(YZM_URL, timeout=10)
            verify_response.raise_for_status()  # 检查请求是否成功
            verify_data = verify_response.json()  # 解析JSON数据
            codes_list = verify_data.get("codes", [])  # 提取验证码列表
            print(f"获取到验证码列表：{codes_list}")

            # 检查验证码列表是否为空
            if not codes_list:
                raise RuntimeError("验证码接口返回的codes列表为空")
        except Exception as e:
            raise RuntimeError(f"获取验证码列表失败: {e}")

        # 第二步：遍历验证码列表尝试登录（从索引0开始）
        for idx, yzm in enumerate(codes_list):
            try:
                print(f"正在尝试第{idx + 1}个验证码：{yzm}")

                # 清空验证码输入框并填入当前验证码
                page.locator('#msgCode').fill("")  # 先清空防止残留
                page.locator('#msgCode').fill(yzm)

                # 执行登录JS函数
                page.evaluate("doLogin();")
                page.wait_for_timeout(5000)  # 等待登录请求处理

                # 执行页面跳转函数（核心业务跳转）
                page.evaluate("gores('200007','20988','');")
                page.wait_for_timeout(3000)  # 等待跳转完成

                # 验证是否登录成功（通过判断是否打开新页面/目标页面是否加载）
                # 方式1：检查页面数量是否增加（新页面打开）
                if len(self.context.pages) > 1:
                    login_success = True
                    break
                # 方式2：检查页面URL是否跳转到目标域（备选验证方式）
                current_url = page.url
                if "omms.chinatowercom.cn" in current_url:
                    login_success = True
                    break

                # 如果未成功，提示并继续下一个验证码
                page.wait_for_timeout(2000)

            except Exception as e:
                print(f"尝试验证码{yzm}时出错: {e}")
                continue

        # 遍历完所有验证码仍未成功
        if not login_success:
            raise RuntimeError(f"遍历所有验证码（共{len(codes_list)}个）均登录失败")

        print("登录流程执行完成，返回登录页面对象")
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