# scheduler_keep_cookies.py
import time
import threading
import schedule
from core.task_logger import log_task_execution
from spider.script.cookies_keep_energy import main_keep_cookies,main_get_cookies
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

def schedule_loop():
    """主调度循环"""
    print("循环开始")
    schedule.every(5).minutes.do(run_task_in_thread, main_keep_cookies, "保活-能源cookies")
    schedule.every(120).minutes.do(run_task_in_thread, main_get_cookies, "更新-能源cookies")

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    schedule_loop()

