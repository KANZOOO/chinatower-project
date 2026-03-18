# app/api/web_api.py
from fastapi import APIRouter, File, UploadFile
from pydantic import BaseModel
import datetime
from core.task_logger_service import log_task_execution
from core.config import settings


# 👇 定义请求/响应模型
class TaskResponse(BaseModel):
    status: str

