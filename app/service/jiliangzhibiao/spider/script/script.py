import sys
import os
import time
import threading
import schedule

# 添加项目根目录到 Python 路径
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..', '..', '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from app.service.jiliangzhibiao.spider.spider import run_all_downloads
from app.service.jiliangzhibiao.spider.process import run_device_process, run_lessee_process

class Jiliangzhibiao:
    def run_down(self):
        # 统一调用 spider.py 中的总下载入口：
        # 同时下载「分路计量租户数据(9个) + 分路计量设备清单(1个)」
        run_all_downloads()
    def process(self):
        # 只跑设备流程：保留下一行，注释 run_lessee_process()
        # 只跑租户流程：注释 run_device_process()，保留下一行
        # 两个都跑：两行都保留
        ok_device = True
        ok_lessee = True
        ok_device = run_device_process()
        ok_lessee = run_lessee_process()
        if not (ok_device and ok_lessee):
            raise RuntimeError("jiliangzhibiao process 执行失败")

    def main_task(self):
        print("📥 run_down 开始")
        self.run_down()
        print("✅ run_down 完成")
        print("⚙️ process 开始")
        self.process()
        print("✅ process 完成")


def run_task_in_thread(task_func, task_name):
    """在线程中执行任务，避免阻塞 schedule 循环。"""
    def _runner():
        try:
            print(f"▶ 开始执行任务：{task_name}")
            task_func()
            print(f"✅ 任务完成：{task_name}")
        except Exception as e:
            print(f"❌ 任务失败：{task_name}，错误：{str(e)}")

    thread = threading.Thread(target=_runner, daemon=True)
    thread.start()


def schedule_loop():
    """主调度循环"""
    print("定时任务调度器已启动")
    print("每天 08:00 自动执行数据下载和处理")
    zhi_lian_tongbao = Jiliangzhibiao()

    schedule.every().day.at("08:00").do(
        run_task_in_thread,
        zhi_lian_tongbao.main_task,
        "每天 8 点运行脚本"
    )
    while True:
        schedule.run_pending()
        time.sleep(1)

def main():
    ji_liang_zhibiao = Jiliangzhibiao()
    try:
        ji_liang_zhibiao.main_task()
    except Exception as e:
        print(f"❌ 脚本执行失败：{str(e)}")
        raise
if __name__ == '__main__':
    main()
    schedule_loop()
