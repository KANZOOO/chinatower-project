# scheduler_keep_zhilian.py
import time
import threading
import schedule
from scheduler.task_logger import log_task_execution
from spider.script.keepalive.cookies_keep_zhilian import KeepZhiLian


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
    print("AIOT 循环开始")

    # 每5分钟保活一次（检查 Authorization 是否有效，失效则重新登录）
    schedule.every(5).minutes.do(
        run_task_in_thread,
        lambda: KeepZhiLian().keep_auth(),
        "保活-AIOT智联"
    )

    # 每6小时强制重新登录，刷新 Authorization
    schedule.every(360).minutes.do(
        run_task_in_thread,
        lambda: KeepZhiLian().get_auth(),
        "更新-AIOT智联"
    )

    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    # 启动时先初始化
    print("启动时初始化 AIOT...")
    KeepZhiLian().get_auth()

    print("AIOT 初始化完成，进入定时循环")
    schedule_loop()