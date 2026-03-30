# utils/task_logger.py
import time
from datetime import datetime
import pandas as pd
from typing import Callable, Any
from core.sql import sql_orm
from core.config import settings

# 初始化数据库 ORM（假设你已有）
orm = sql_orm(settings.db_url)

def log_task_execution(task_name: str, func: Callable, *args, **kwargs) -> Any:
    """
    通用任务执行 + 日志记录装饰器/调用器
    - 自动记录开始时间、结束时间、耗时、状态、错误信息
    - 写入数据库表 task_log
    """
    start_time = time.time()
    db_start = datetime.now()
    error_msg = None
    status = "SUCCESS"
    result = None

    try:
        result = func(*args, **kwargs)
        return result
    except Exception as exc:
        status = "FAILURE"
        error_msg = str(exc)
        # 可选：打印到 stdout/stderr 便于调试
        print(f"[TASK ERROR] {task_name} failed: {error_msg}")
        raise  # 保留异常传播
    finally:
        duration = round(time.time() - start_time, 3)
        log_data = {
            "task_name": task_name,
            "status": status,
            "start_time": db_start,
            "end_time": datetime.now(),
            "duration_seconds": duration,
            "error_message": error_msg or "",
        }

        # 写入数据库
        try:
            df = pd.DataFrame([log_data])
            orm.add_data(df, "task_log")
        except Exception as e:
            print(f"[LOG WRITE ERROR] Failed to log task '{task_name}': {e}")