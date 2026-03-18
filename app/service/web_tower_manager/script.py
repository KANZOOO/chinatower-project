# app/service/web_tower_manager/script.py
from sqlalchemy import text
import pandas as pd
from core.sql import sql_orm
import logging
from typing import Dict, List, Any, Tuple, Optional

logger = logging.getLogger(__name__)
db = sql_orm()
def get_db_engine():
    return db.get_engine()


class AddressBookService:
    @staticmethod
    async def get_list(params: Dict[str, str], role: str) -> List[Dict[str, Any]]:
        """
        获取通讯录列表
        params: 查询参数字典，包含 city, area, name, phone, belong, position, businessCategory, specificBusiness, level
        role: 用户角色 (tt/all)
        返回: 数据列表
        """
        try:
            # 构建 SQL
            sql = "SELECT * FROM core.address_book WHERE 1=1"
            query_params = {}

            # 权限过滤
            if role == 'tt':
                sql += " AND belong = :belong"
                query_params['belong'] = '代维人员'

            # 筛选条件
            city = params.get('city', '')
            area = params.get('area', '')
            name = params.get('name', '')
            phone = params.get('phone', '')
            belong = params.get('belong', '')
            position = params.get('position', '')
            businessCategory = params.get('businessCategory', '')
            specificBusiness = params.get('specificBusiness', '')
            level = params.get('level', '')

            if city:
                sql += " AND city = :city"
                query_params['city'] = city
            if area and area != '全部':
                sql += " AND area = :area"
                query_params['area'] = area
            if name:
                sql += " AND name LIKE :name"
                query_params['name'] = f'%{name}%'
            if phone:
                sql += " AND phone LIKE :phone"
                query_params['phone'] = f'%{phone}%'
            if belong:
                sql += " AND belong = :belong"
                query_params['belong'] = belong
            if position:
                sql += " AND position = :position"
                query_params['position'] = position
            if businessCategory:
                sql += " AND businessCategory = :business_category"
                query_params['business_category'] = businessCategory
            if specificBusiness:
                sql += " AND specificBusiness = :specificBusiness"
                query_params['specificBusiness'] = specificBusiness
            if level:
                sql += " AND level = :level"
                query_params['level'] = level

            sql += " ORDER BY id DESC"

            # 执行查询
            engine = get_db_engine()
            df = pd.read_sql(text(sql), engine, params=query_params)

            return df.to_dict(orient='records')

        except Exception as e:
            logger.error(f"查询失败: {str(e)}")
            return []

    @staticmethod
    async def update(data: Dict[str, Any], operation: str, role: str) -> Tuple[bool, str, Optional[Dict]]:
        try:
            row = data

            # 权限检查
            if role == 'tt' and row.get('belong') != '代维人员':
                return False, "您只能操作代维人员数据", None

            # 数据验证
            if not row.get('name') or not row.get('phone'):
                return False, "姓名/电话不能为空", None

            # 准备字段
            cols = {
                'city': row.get('city', ''),
                'area': row.get('area', ''),
                'belong': row.get('belong', ''),
                'position': row.get('position', ''),
                'level': row.get('level', ''),
                'name': row.get('name', ''),
                'phone': row.get('phone', ''),
                'businessCategory': row.get('businessCategory', ''),
                'specificBusiness': row.get('specificBusiness', '')
            }

            engine = get_db_engine()

            if operation == 'add':
                # 新增
                sql = """
                INSERT INTO core.address_book
                (city, area, belong, position, level, name, phone, businessCategory, specificBusiness)
                VALUES
                (:city, :area, :belong, :position, :level, :name, :phone, :businessCategory, :specificBusiness)
                """
                with engine.connect() as conn:
                    result = conn.execute(text(sql), cols)
                    conn.commit()
                    new_id = result.lastrowid
                    logger.info(f"新增成功, ID: {new_id}")
                return True, "操作成功", {"id": new_id}
            else:
                # 编辑
                if not row.get('id'):
                    return False, "更新失败：缺少有效数据ID", None

                cols['id'] = row.get('id')
                sql = """
                UPDATE core.address_book
                SET city=:city, area=:area, belong=:belong, position=:position,
                    level=:level, name=:name, phone=:phone, 
                    businessCategory=:businessCategory, specificBusiness=:specificBusiness
                WHERE id=:id
                """
                with engine.connect() as conn:
                    conn.execute(text(sql), cols)
                    conn.commit()
                    logger.info(f"更新成功, ID: {cols['id']}")
                return True, "操作成功", None

        except Exception as e:
            logger.error(f"更新失败: {str(e)}")
            return False, f"操作失败：{str(e)}", None

    @staticmethod
    async def delete(row_id: int, role: str) -> Tuple[bool, str]:
        try:
            if not row_id:
                return False, "id 不能为空"

            engine = get_db_engine()

            # 权限检查（tt角色只能删除代维人员）
            if role == 'tt':
                with engine.connect() as conn:
                    result = conn.execute(
                        text("SELECT belong FROM core.address_book WHERE id=:id"),
                        {'id': row_id}
                    )
                    belong = result.scalar()
                    if belong != '代维人员':
                        return False, "您只能删除代维人员数据"

            # 执行删除
            with engine.connect() as conn:
                conn.execute(
                    text("DELETE FROM core.address_book WHERE id = :id"),
                    {'id': row_id}
                )
                conn.commit()
                logger.info(f"删除成功, ID: {row_id}")

            return True, "删除成功"

        except Exception as e:
            logger.error(f"删除失败: {str(e)}")
            return False, f"删除失败：{str(e)}"

    @staticmethod
    async def get_statistics(params: Dict[str, str], role: str) -> List[Dict[str, Any]]:
        try:
            city = params.get('city', '')
            area = params.get('area', '')
            businessCategory = params.get('businessCategory', '')

            # 构建 SQL
            sql = """
            SELECT 
              city,
              area,
              SUM(CASE WHEN businessCategory = '一体' THEN 1 ELSE 0 END) AS business_one,
              SUM(CASE WHEN businessCategory = '能源' THEN 1 ELSE 0 END) AS business_energy,
              SUM(CASE WHEN businessCategory = '拓展' THEN 1 ELSE 0 END) AS business_expand,
              SUM(CASE WHEN level = '一级督办对象' THEN 1 ELSE 0 END) AS level_1,
              SUM(CASE WHEN level = '二级督办对象' THEN 1 ELSE 0 END) AS level_2,
              SUM(CASE WHEN level = '三级督办对象' THEN 1 ELSE 0 END) AS level_3,
              SUM(CASE WHEN level = '四级督办对象' THEN 1 ELSE 0 END) AS level_4,
              SUM(CASE WHEN level = '五级督办对象' THEN 1 ELSE 0 END) AS level_5
            FROM core.address_book
            WHERE 1=1
            """
            query_params = {}

            if role == 'tt':
                sql += " AND belong = :belong"
                query_params['belong'] = '代维人员'

            if city:
                sql += " AND city = :city"
                query_params['city'] = city
            if area:
                sql += " AND area = :area"
                query_params['area'] = area
            if businessCategory:
                sql += " AND businessCategory = :business_category"
                query_params['business_category'] = businessCategory

            sql += " GROUP BY city, area"

            engine = get_db_engine()
            with engine.connect() as conn:
                result = conn.execute(text(sql), query_params)
                data = [dict(zip(result.keys(), row)) for row in result]

            return data

        except Exception as e:
            logger.error(f"统计接口异常：{str(e)}")
            return []

address_book_service = AddressBookService()