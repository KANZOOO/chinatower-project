# scheduler/scheduler_main.py
import schedule
import time
import threading
from core.task_logger import log_task_execution
from app.service.msg_energy_offline.script import main as msg_energy_offline
from app.service.msg_energy_order.script import main as msg_energy_order
from app.service.msg_zhilian_order.script import main as msg_zhilian_order
from app.service.msg_zhilian_fault.script import main as msg_zhilian_fault
from app.service.msg_zhilian_transit.script import main as msg_zhilian_transit
from app.service.msg_freq_overtime_station_out.script import main as msg_freq_overtime_station_out
from app.service.msg_earthquake.script import main as msg_earthquake
from app.service.msg_weather_report.script import main as msg_weather
from app.service.msg_weather_report.script2 import main as msg_weather_hour

from spider.script.oracle_data import main as oracle_data

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
    schedule.every(5).minutes.do(run_task_in_thread,msg_energy_offline,"能源设备离线短信")
    schedule.every(5).minutes.do(run_task_in_thread,msg_energy_order,"能源工单短信")
    schedule.every(5).minutes.do(run_task_in_thread,msg_zhilian_order,"智联工单短信")
    # schedule.every(5).minutes.do(run_task_in_thread,msg_zhilian_fault,"智联故障监控短信")
    # schedule.every(5).minutes.do(run_task_in_thread,msg_zhilian_transit,"智联在途工单短信")
    schedule.every(5).minutes.do(run_task_in_thread, msg_freq_overtime_station_out, "超长超频退服提醒")
    schedule.every(5).minutes.do(run_task_in_thread, oracle_data, "全量活动告警-退服历史告警")
    schedule.every(5).minutes.do(run_task_in_thread, msg_earthquake, "地震设备故障提醒")
    schedule.every(60).minutes.do(run_task_in_thread, msg_weather_hour, "每小时天气预报")
    schedule.every().day.at("08:30").do(run_task_in_thread, msg_weather, "每日天气预报")
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    # oracle_data()
    schedule_loop()

