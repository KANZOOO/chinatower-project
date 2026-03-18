# scheduler_keep_cookies.py
import time
import threading
import schedule
from scheduler.task_logger import log_task_execution
from spider.script.keepalive.cookies_keep_foura import KeepFourA


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

    # 每5分钟保活一次（检查所有3个账号，失效的自动重新登录）
    schedule.every(5).minutes.do(
        run_task_in_thread,
        lambda: KeepFourA().keep_cookies(),
        "保活-运监cookies"
    )

    # 每6小时强制重新登录所有账号，刷新cookie
    schedule.every(360).minutes.do(
        run_task_in_thread,
        lambda: KeepFourA().get_all_cookies(),
        "更新-运监cookies"
    )

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    # 启动时先初始化所有3个账号
    print("启动时初始化所有账号...")
    KeepFourA().get_all_cookies()

    print("初始化完成，进入定时循环")
    schedule_loop()