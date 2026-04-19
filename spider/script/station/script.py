from core.config import settings
from spider.script.model import down_file
from spider.schema.schema import zhineng_wangguan, zhineng_shexiangtou
import schedule
import time
import subprocess
import sys
from threading import Thread

class ZhiLianTongBao:
    def __init__(self):
        self.path_wangguan=settings.resolve_path("spider/script/station/down/边缘网关.xls")
        self.path_shexiangtou=settings.resolve_path("spider/script/station/down/边缘摄像头.xls")
    
    def _run_step_subprocess(self, step_name, code):
        """
        将每个处理步骤放到独立Python子进程中执行，隔离Excel COM状态，
        避免pywin32与xlwings在同一进程串行运行时造成工作簿数据异常。
        """
        print(f"开始执行：{step_name}")
        result = subprocess.run(
            [sys.executable, "-c", code],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.stdout:
            print(result.stdout.rstrip())
        if result.returncode != 0:
            if result.stderr:
                print(result.stderr.rstrip())
            raise RuntimeError(f"{step_name}执行失败，退出码={result.returncode}")
        print(f"执行完成：{step_name}")

    def down(self, url, data, path):
        for key, into_data in data.items():
            if key in ['INTO_DATA_1', 'INTO_DATA_2', 'INTO_DATA_FINAL']:
                data[key]['queryForm:queryAccessStatus'] = "02"
        down_file(url=url, data=data, path=path)

    def run_down(self):
        wangguan_url = "http://omms.chinatowercom.cn:9000/business/resMge/devMge/listEdgeGateway.xhtml"
        self.down(url=wangguan_url, data=zhineng_wangguan, path=self.path_wangguan)
        print(f" 网关数据下载成功")
        
        shexiangtou_url = "http://omms.chinatowercom.cn:9000/business/resMge/devMge/listEdgeGatewayDev.xhtml"
        self.down(url=shexiangtou_url, data=zhineng_shexiangtou, path=self.path_shexiangtou)
        print(f" 摄像头数据下载成功")
    
    def process(self):
        # 注意：网关(pywin32)与摄像头(xlwings)处理放到不同子进程，避免Excel实例冲突导致数据错乱。
        self._run_step_subprocess(
            "网关全流程处理",
            "from spider.script.gateway_process import process_gateway_full_list; process_gateway_full_list()",
        )
        self._run_step_subprocess(
            "摄像头清单处理",
            "from spider.script.monitor_process import process_camera_main_list; process_camera_main_list()",
        )
        print("数据处理完成")
    
    def main_task(self):
        """完整的下载 + 处理流程"""
        self.run_down()
        self.process()

def main():
    zhi_lian_tongbao = ZhiLianTongBao()
    zhi_lian_tongbao.run_down()
    zhi_lian_tongbao.process()


def run_task_in_thread(task_func, task_name):
    """在线程中运行任务，避免阻塞调度器"""

    def wrapper():
        try:
            print(f"\n🕒 [{time.strftime('%Y-%m-%d %H:%M:%S')}] 开始执行：{task_name}")
            task_func()
            print(f"✅ [{time.strftime('%Y-%m-%d %H:%M:%S')}] 执行完成：{task_name}\n")
        except Exception as e:
            print(f"❌ 任务执行失败 [{task_name}]: {e}")
            import traceback
            traceback.print_exc()

    thread = Thread(target=wrapper)
    thread.start()


def schedule_loop():
    """主调度循环"""
    print("定时任务调度器已启动")
    print("每天 08:00 自动执行数据下载和处理")
    zhi_lian_tongbao = ZhiLianTongBao()

    schedule.every().day.at("08:00").do(
        run_task_in_thread,
        zhi_lian_tongbao.main_task,
        "每天 8 点运行脚本"
        # "数据下载"
    )
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    main()
    schedule_loop()

