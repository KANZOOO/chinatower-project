# utils/task_logger_service.py
import time
from datetime import datetime
import pandas as pd
from typing import Callable, Any, Optional
from functools import wraps
from core.sql import sql_orm
from core.config import settings
import asyncio
orm = sql_orm(settings.db_url)

def log_task_execution(task_name: str):
    """
    装饰器工厂：返回一个装饰器，用于包裹异步或同步函数，并记录执行日志。
    """
    def decorator(func: Callable) -> Callable:
        if asyncio.iscoroutinefunction(func):
            @wraps(func)
            async def async_wrapper(*args, **kwargs):
                start_time = time.time()
                db_start = datetime.now()
                error_msg = None
                status = "SUCCESS"
                result = None

                try:
                    result = await func(*args, **kwargs)
                    return result
                except Exception as exc:
                    status = "FAILURE"
                    error_msg = str(exc)
                    print(f"[TASK ERROR] {task_name} failed: {error_msg}")
                    raise
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
                    try:
                        df = pd.DataFrame([log_data])
                        orm.add_data(df, "task_log")
                    except Exception as e:
                        print(f"[LOG WRITE ERROR] Failed to log task '{task_name}': {e}")
            return async_wrapper
        else:
            @wraps(func)
            def sync_wrapper(*args, **kwargs):
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
                    print(f"[TASK ERROR] {task_name} failed: {error_msg}")
                    raise
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
                    try:
                        df = pd.DataFrame([log_data])
                        orm.add_data(df, "task_log")
                    except Exception as e:
                        print(f"[LOG WRITE ERROR] Failed to log task '{task_name}': {e}")
            return sync_wrapper
    return decorator