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


def refresh_all_cookies():
    print("强制刷新所有账号cookie...")
    keep_foura().get_cookies()
def schedule_loop():
    """主调度循环"""
    print("循环开始")
    # 每6小时强制重新登录所有账号，刷新cookie
    schedule.every(6).hours.do(
        run_task_in_thread,
        refresh_all_cookies,
        "更新-运监cookies"
    )

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    refresh_all_cookies()
    schedule_loop()