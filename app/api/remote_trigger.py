# app/api/remote_trigger.py
from fastapi import APIRouter, File, UploadFile
from pydantic import BaseModel
import datetime
from core.task_logger_service import log_task_execution
from app.service.msg_zhilian_online.script import main as msg_zhilian_online
from app.service.msg_freq_overtime_fsu_offline.script import main as msg_freq_overtime_fsu_offline

from core.config import settings


# 👇 定义请求/响应模型
class TaskResponse(BaseModel):
    status: str

# 👇 创建独立的 router 实例
router = APIRouter(
    prefix="/tt/tower",  # 统一前缀，避免每个路由写 /tt
    tags=["任务触发"]  # 用于 Swagger 分组
)
# 👇 规范化的 POST 路由
@router.post(
    "/zhilian_online",
    response_model=TaskResponse,
    summary="触发智联设备在线提醒",
    description="远程传入文件，保存，执行服务器上的智联设备在线状态检测任务"
)
@log_task_execution("服务器-智联设备在线提醒")
async def run_msg_zhilian_online(file: UploadFile = File(...)) -> TaskResponse:
    contents = await file.read()
    file_path=settings.resolve_path(r'app/service/msg_zhilian_online/data/智联在线情况.xlsx')
    with open(file_path, "wb") as buffer:
        buffer.write(contents)
    msg_zhilian_online()
    return TaskResponse(status="任务已执行并记录日志")
# 👇 规范化的 POST 路由
@router.post(
    "/msg_freq_overtime_fsu_offline",
    response_model=TaskResponse,
    summary="触发超长超频fsu离线短信",
    description="远程传入文件，保存，执行"
)
@log_task_execution("服务器-超长超频fsu离线")
async def run_msg_freq_overtime_fsu_offline(file: UploadFile = File(...)) -> TaskResponse:
    contents = await file.read()
    now = datetime.datetime.now().strftime("%Y%m%d")
    path = settings.resolve_path(f"app/service/msg_freq_overtime_fsu_offline/data/fsu离线情况{now}.xlsx")
    with open(path, "wb") as buffer:
        buffer.write(contents)
    msg_freq_overtime_fsu_offline()
    return TaskResponse(status="任务已执行并记录日志")