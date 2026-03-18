import asyncio
import os
import re
import pickle
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# 假设这些模块已存在
from core.config import settings  # 用于获取项目根目录等
from core.msg.msg_ding import DingMsg

class WeatherSpider:
    def __init__(self):
        self.base_url = "http://gx.weather.com.cn"
        self.weather_url = f"{self.base_url}/weather.shtml"
        self.user_data_dir = settings.resolve_path("playwright_user_data")
        os.makedirs(self.user_data_dir, exist_ok=True)
        self.data_text = ""
        self.pkl_path=settings.resolve_path("app\service\msg_weather_report\weather_data.pkl")
    async def parse_land_forecast(self, html_content: str):
        if not html_content:
            print("没有获取到页面内容")
            return None

        soup = BeautifulSoup(html_content, 'html.parser')
        dl_tag = soup.find('dl', id='mainContent')
        h3_tag = soup.find('h3')
        if h3_tag:
            date_text = h3_tag.get_text(strip=True)
            date_text = date_text.replace('【字体：大中小】', '').strip()
            date_text = date_text.split('来源：')[0].strip()
            self.data_text = date_text

        if dl_tag:
            p_tags = dl_tag.find_all('p')
            guangxi_cities = ["南宁", "柳州", "桂林", "梧州", "北海", "防城港", "钦州", "贵港", "玉林", "百色", "贺州", "河池", "来宾", "崇左", "全区"]
            for p in p_tags:
                text = p.get_text(strip=True).replace('\n', '').replace('\r', '')
                if text.startswith("今天") and any(city in text for city in guangxi_cities):
                    return text
            print("未找到目标段落")
        else:
            print("未找到广西陆地天气预报信息")
        return None

    async def parse_weather_alerts(self, page, alert_url: str):
        full_url = alert_url if alert_url.startswith('http') else f"{self.base_url}/{alert_url}"
        try:
            await page.goto(full_url, timeout=30000)
            await page.wait_for_selector(".commontop .left", timeout=10000)
            await page.wait_for_selector(".alarml", timeout=10000)

            datetime_element = await page.query_selector(".commontop .left")
            datetime_text = (await datetime_element.inner_text()).strip() if datetime_element else "日期时间未知"

            alert_list = await page.query_selector(".alarml")
            if not alert_list:
                return []

            li_elements = await alert_list.query_selector_all("li")
            alerts_data = []

            for li in li_elements:
                alert_text = (await li.inner_text()).strip()
                if not alert_text:
                    continue

                # 提取时间
                datetime_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}', alert_text)
                datetime_str = datetime_match.group() if datetime_match else "未知时间"

                # 提取地区
                location_match = re.search(r'广西壮族自治区[^发布]+', alert_text)
                location_str = location_match.group() if location_match else "未知地区"

                # 提取预警类型
                alert_type_match = re.search(r'发布[^预警]+预警', alert_text)
                alert_type_str = alert_type_match.group() if alert_type_match else "未知预警"

                final_alert = f"{datetime_str}，{location_str}{alert_type_str}"
                alerts_data.append(final_alert)

            return alerts_data

        except PlaywrightTimeoutError:
            print("预警页面加载超时")
            return []
        except Exception as e:
            print(f"解析预警信息出错: {e}")
            return []

    async def run(self):
        async with async_playwright() as p:
            context = await p.chromium.launch_persistent_context(
                user_data_dir=self.user_data_dir,
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-gpu",
                    "--disable-dev-shm-usage",
                ],
            )

            try:
                # 主页：获取天气预报链接
                page = await context.new_page()
                await page.goto(self.weather_url, timeout=30000)
                await page.wait_for_timeout(3000)  # 确保 JS 渲染完成

                content = await page.content()
                soup = BeautifulSoup(content, 'html.parser')

                # === 1. 获取陆地天气预报 ===
                target_div1 = soup.find('div', class_='indextopnews')
                link1 = target_div1.find('a') if target_div1 else None
                land_forecast = None

                if link1 and (link1_href := link1.get('href')):
                    details_url = f"{self.base_url}/{link1_href.lstrip('/')}"
                    detail_page = await context.new_page()
                    await detail_page.goto(details_url, timeout=30000)
                    await detail_page.wait_for_timeout(2000)
                    detail_content = await detail_page.content()
                    land_forecast = await self.parse_land_forecast(detail_content)
                    await detail_page.close()

                # === 2. 获取预警信息 ===
                target_div2 = soup.find('div', class_='bottom indexbottom')
                link2 = target_div2.find('a') if target_div2 else None
                alerts = []

                if link2 and (link2_href := link2.get('href')):
                    alert_page = await context.new_page()
                    alerts = await self.parse_weather_alerts(alert_page, link2_href)
                    await alert_page.close()

                # === 3. 构建数据结构 ===
                data = {
                    "广西": (self.data_text + '，' + land_forecast).replace('。', '') if land_forecast else self.data_text,
                    "南宁": "", "柳州": "", "桂林": "", "梧州": "", "北海": "",
                    "防城港": "", "钦州": "", "贵港": "", "玉林": "", "百色": "",
                    "贺州": "", "河池": "", "来宾": "", "崇左": ""
                }

                for alert in alerts:
                    for city in data.keys():
                        if city != '广西' and city in alert:
                            data[city] += alert + '<br>'
                            break

                # === 4. 比较并发送钉钉通知 ===
                old_data = {}
                if os.path.exists(self.pkl_path):
                    try:
                        with open(self.pkl_path, "rb") as file:
                            old_data = pickle.load(file)
                    except Exception as e:
                        print(f"读取旧数据失败: {e}")

                if old_data.get('广西') != data['广西']:
                    text = f"广西天气预报：\n{land_forecast}\n(预警信息来源： 中国天气网)"
                    print(text)
                    d = DingMsg()
                    await asyncio.to_thread(d.text_at, d.WEATHER, text)  # 如果 text_at 是同步函数

                # 保存新数据
                with open(self.pkl_path, 'wb') as f:
                    pickle.dump(data, f)

            except Exception as e:
                print(f"爬虫运行出错: {e}")
                raise
            finally:
                await context.close()


# 异步入口（可单独调用）
async def main():
    spider = WeatherSpider()
    await spider.run()

# 如果直接运行此脚本
if __name__ == "__main__":
    asyncio.run(main())