import logging
from app.service.web_tower_manager.script import address_book_service
from fastapi import HTTPException,  status,APIRouter, Request, Header, Query
from typing import Optional
logger = logging.getLogger(__name__)

router = APIRouter(
    prefix="/tt/tower/api",
    tags=["通讯录管理"]
)

@router.get("/address_book")
async def address_book_list(
        request: Request,
        city: str = Query("", description="城市"),
        area: str = Query("", description="区域"),
        name: str = Query("", description="姓名"),
        phone: str = Query("", description="电话"),
        belong: str = Query("", description="归属"),
        position: str = Query("", description="岗位"),
        businessCategory: str = Query("", description="业务大类"),
        specificBusiness: str = Query("", description="具体业务"),
        level: str = Query("", description="督办等级"),
        x_role: Optional[str] = Header(None, alias="X-Role"),
        role: str = Query("tt", description="角色")):
    actual_role = x_role or role or "tt"
    # 如果角色不是 "tt"，则严格限制 IP
    # 获取真实客户端 IP（支持 X-Forwarded-For）
    client_ip = request.headers.get("x-forwarded-for")
    if client_ip:
        client_ip = client_ip.split(",")[0].strip()
    else:
        client_ip = request.client.host if request.client else "127.0.0.1"
    if actual_role != "tt":
        allowed_ip = "113.16.135.199"
        if client_ip != allowed_ip:
            logger.warning(f"IP 访问拒绝: role={actual_role}, ip={client_ip}, path=/address_book")
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Access denied: only allowed from specific IP for non-tt roles."
            )

    logger.info(f"查询参数: role={actual_role}, city={city}, area={area}, name={name}")
    params = {
        'city': city,
        'area': area,
        'name': name,
        'phone': phone,
        'belong': belong,
        'position': position,
        'businessCategory': businessCategory,
        'specificBusiness': specificBusiness,
        'level': level
    }
    result = await address_book_service.get_list(params, actual_role)
    return result

@router.post("/address_book/update")
async def update_address_book(request: Request):
    try:
        data = await request.json()
        row = data.get('data', {})
        op = data.get('operation', 'add')
        role = request.headers.get('X-Role', 'tt')
        logger.info(f"更新操作: {op}, 数据: {row}, 角色: {role}")
        success, msg, result_data = await address_book_service.update(row, op, role)
        if success:
            return {"code": 0, "msg": msg}
        else:
            return {"code": 1, "msg": msg}
    except Exception as e:
        logger.error(f"更新接口异常: {str(e)}")
        return {"code": 1, "msg": f"操作失败：{str(e)}"}

@router.post("/address_book/delete")
async def delete_address_book(request: Request):
    try:
        data = await request.json()
        row_id = data.get('id')
        role = request.headers.get('X-Role', 'tt')
        logger.info(f"删除操作, ID: {row_id}, 角色: {role}")
        success, msg = await address_book_service.delete(row_id, role)
        if success:
            return {"code": 0, "msg": msg}
        else:
            return {"code": 1, "msg": msg}
    except Exception as e:
        logger.error(f"删除接口异常: {str(e)}")
        return {"code": 1, "msg": f"删除失败：{str(e)}"}

@router.get("/address_book/statistics")
async def get_address_book_statistics(
        city: str = Query(""),
        area: str = Query(""),
        businessCategory: str = Query(""),
        x_role: str = Header(None, alias="X-Role"),
        role: str = Query("tt")):
    try:
        actual_role = x_role or role or "tt"
        params = {
            'city': city,
            'area': area,
            'businessCategory': businessCategory
        }
        data = await address_book_service.get_statistics(params, actual_role)
        return data
    except Exception as e:
        logger.error(f"统计接口异常: {str(e)}")
        return []