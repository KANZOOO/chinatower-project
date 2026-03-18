import asyncio
from playwright.async_api import async_playwright
import os
from core.config import settings
from core.msg.msg_ding import DingMsg
async def fetch_weather_image_url():
    index_url = 'http://products.weather.com.cn/yb.html?page=YB&type=YB_TQQS_3D'
    user_data_dir = settings.resolve_path("playwright_user_data")
    os.makedirs(user_data_dir, exist_ok=True)

    async with async_playwright() as p:
        # 启动持久化上下文（带用户数据目录）
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=True,  # 无头模式
            args=[
                "--no-sandbox",
                "--disable-gpu",
                "--remote-debugging-port=9222",
                "--disable-dev-shm-usage",  # 避免内存不足（可选）
            ],
        )

        page = context.pages[0] if context.pages else await context.new_page()

        try:
            # 清除 cookies（Playwright 中清除 storage 更彻底）
            await context.clear_cookies()
            await page.goto(index_url, timeout=30000)
            await page.wait_for_timeout(5000)  # 等待5秒

            # 刷新页面（模拟你的 refresh）
            await page.reload(timeout=30000)
            await page.wait_for_timeout(5000)

            # 提取文本（可选，你主要需要图片 URL）
            text_father = await page.query_selector("//b[contains(text(),'未来三天')]/../../following-sibling::p[1]")
            if text_father:
                spans = await text_father.query_selector_all('span')
                text = ''.join([await span.inner_text() for span in spans])

            # 获取图片元素的 src
            img_element = await page.query_selector("//b[contains(text(),'未来三天')]/../../following-sibling::p[2]/child::span/child::img")
            if img_element:
                down_url = await img_element.get_attribute('src')
                d = DingMsg()
                await asyncio.to_thread(d.picture, d.WEATHER,down_url,text,"天气预报")  # 如果 text_at 是同步函数


        except Exception as e:
            print("Error:", e)
        finally:
            await context.close()

def main():
    asyncio.run(fetch_weather_image_url())
# 运行协程
if __name__ == "__main__":
    main()