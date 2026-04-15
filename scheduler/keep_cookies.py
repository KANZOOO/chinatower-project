# scheduler_keep_foura.py
import time
import threading
import schedule
from core.task_logger import log_task_execution
from spider.script.cookies_foura import main as keep_foura


def run_task_in_thread(task_func, task_name):
    """在独立线程中运行任务并记录日志"""
    def wrapper():
        try:
            log_task_execution(task_name, task_func)
        except Exception as e:
            print(f"{task_name}: {e}")

    thread = threading.Thread(target=wrapper, daemon=True)
    thread.start()
    return thread


def refresh_all_cookies_with_retry():
    """带重试机制的cookie刷新"""
    max_retries = 2
    retry_delay = 60
    
    for attempt in range(max_retries + 1):
        try:
            print(f"强制刷新所有账号cookie... (第{attempt + 1}次尝试)")
            keep_foura()
            print("Cookie刷新成功！")
            return True
        except Exception as e:
            print(f"Cookie刷新失败（第{attempt + 1}次）：{e}")
            if attempt < max_retries:
                print(f"等待{retry_delay}秒后重试...")
                time.sleep(retry_delay)
            else:
                print("已达到最大重试次数，放弃本次刷新")
                return False


def refresh_all_cookies():
    """定时任务调用的入口函数"""
    refresh_all_cookies_with_retry()


def schedule_loop():
    """主调度循环"""
    print("调度循环开始")
    
    # 每6小时强制重新登录所有账号，刷新cookie
    schedule.every(6).hours.do(
        run_task_in_thread,
        refresh_all_cookies,
        "更新-运监cookies"
    )
    
    # 每天早上7点30分固定运行
    schedule.every().day.at("07:30").do(
        run_task_in_thread,
        refresh_all_cookies,
        "更新-运监cookies（定时）"
    )

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    refresh_all_cookies_with_retry()
    schedule_loop()