import cx_Oracle
from sqlalchemy import text
import pandas as pd
from datetime import datetime
from core.sql import sql_orm
from core.config import settings

def _convert_enum(field, val, mapping):
    ORACLE_NULL = getattr(cx_Oracle, 'NULL', None)
    if pd.isna(val) or val in (None, ORACLE_NULL):
        return None
    cfg = mapping.get(field)
    if cfg and "enum" in cfg:
        return cfg["enum"].get(str(val), val)
    return val


def _apply_mapping(df, mapping):
    # 仅保留 mapping 中存在的列
    cols_to_keep = [col for col in df.columns if col in mapping]
    df = df[cols_to_keep].copy()

    # 重命名
    rename_map = {k: v["cn_name"] for k, v in mapping.items() if k in cols_to_keep}
    df.rename(columns=rename_map, inplace=True)

    # 枚举转换
    for orig_col, config in mapping.items():
        if orig_col not in cols_to_keep:
            continue
        cn_name = config["cn_name"]
        if "enum" in config:
            df[cn_name] = df[cn_name].apply(lambda x: _convert_enum(orig_col, x, mapping))

    return df


def _get_oracle_connection():
    # Oracle 连接配置（建议后续从环境变量读取）
    ORACLE_CONFIG = {
        "user": "ITOWER_VIF_gxfgs",
        "password": "itowergxfgs_cu8de",
        "host": "101.227.251.222",
        "port": 1521,
        "service_name": "prod"
    }
    dsn = cx_Oracle.makedsn(
        host=ORACLE_CONFIG["host"],
        port=ORACLE_CONFIG["port"],
        service_name=ORACLE_CONFIG["service_name"]
    )
    return cx_Oracle.connect(
        user=ORACLE_CONFIG["user"],
        password=ORACLE_CONFIG["password"],
        dsn=dsn
    )


def fetch_alarmat():
    """
    获取 alarmat 全量数据（仅 STATIONSTATUS = '2'）
    返回清洗并映射后的 DataFrame
    """
    # 字段映射（仅用于重命名和枚举转换）
    ALARM_FIELD_MAPPING = {
        # 基础标识字段
        "ID": {
            "cn_name": "ID",
        },
        # 告警分类相关字段（系统固定值）
        "CATEGORY": {
            "cn_name": "告警类别",
            "enum": {"1": "动环告警", "2": "其它告警"}
        },
        "OBJTYPE": {
            "cn_name": "告警对象类型",
            "enum": {"1": "FSU", "2": "其它类型"}
        },
        "SUBOBJTYPE": {
            "cn_name": "告警子对象类型",
            "enum": {"1": "动环设备", "2": "其它类型"}
        },
        "ALARMTYPE": {
            "cn_name": "告警类型",
            "enum": {"1": "设备告警", "2": "其它告警"}
        },
        # 对象标识字段
        "OBJID": {
            "cn_name": "告警对象ID",
        },
        "OBJNAME": {
            "cn_name": "告警对象名称",
        },
        "SUBOBJID": {
            "cn_name": "告警子对象ID",
        },
        "SUBOBJNAME": {
            "cn_name": "告警子对象名称",
        },
        # 告警等级相关
        "ORIGLEVEL": {
            "cn_name": "原始告警等级",
            "enum": {"1": "一级告警", "2": "二级告警", "3": "三级告警", "4": "四级告警"}
        },
        "ALARMLEVEL": {
            "cn_name": "告警等级",
            "enum": {"1": "一级告警", "2": "二级告警", "3": "三级告警", "4": "四级告警"}
        },
        # 告警描述信息
        "CAUSE": {
            "cn_name": "告警原因",
        },
        "CODE": {
            "cn_name": "告警摘要",
        },
        "DETAILINFO": {
            "cn_name": "告警详情",
        },
        "EXTRAINFO": {
            "cn_name": "告警附加信息",
        },
        "ALARMNAME": {
            "cn_name": "告警名称",
        },
        "ALARMEXPLAIN": {
            "cn_name": "信号量说明",
        },
        "FAULTEXPLAIN": {
            "cn_name": "信号量解释",
        },
        # 时间相关字段
        "INSERTTIME": {
            "cn_name": "告警入库时间",
        },
        "FIRSTSYSTEMTIME": {
            "cn_name": "首次告警的系统时间",
        },
        "LASTSYSTEMTIME": {
            "cn_name": "末次告警的系统时间",
        },
        "CONFIRMTIME": {
            "cn_name": "确认时间",
        },
        # 状态标识字段
        "TIMES": {
            "cn_name": "告警重复次数",
        },
        "IFCONFIRM": {
            "cn_name": "是否已确认",
            "enum": {"Y": "已确认", "N": "未确认"},
        },
        "IFRECOVER": {
            "cn_name": "是否已恢复",
            "enum": {"Y": "已恢复", "N": "未恢复"},
        },
        "IFCONVERTFAULT": {
            "cn_name": "是否转故障",
            "enum": {"Y": "已转故障", "N": "未转故障"},
        },
        "IFATTENTION": {
            "cn_name": "是否关注",
            "enum": {"Y": "已关注", "N": "未关注"},
        },
        "IFOVERTIME": {
            "cn_name": "离首次告警时间是否超过24小时",
            "enum": {"Y": "是", "N": "否"},
        },
        # 关联关系字段
        "FAULTID": {
            "cn_name": "故障单号",
        },
        "STATIONID": {
            "cn_name": "运维监控站址ID",
        },
        "PARENT_ALARMID": {
            "cn_name": "所属主要告警ID",
        },

        "PROVINCEID": {
            "cn_name": "省",
            "enum": {"0098364": "广西"},

        },
        "CITYID": {
            "cn_name": "地市",
            "enum": {"0099977": "南宁",
                     "0099978": "柳州",
                     "0099979": "桂林",
                     "0099980": "玉林",
                     "0099981": "贵港",
                     "0099982": "百色",
                     "0099983": "河池",
                     "0099984": "钦州",
                     "0099985": "梧州",
                     "0099986": "北海",
                     "0099987": "防城港",
                     "0099988": "崇左",
                     "0099989": "来宾",
                     "0099990": "贺州"
                     },
        },
        # 站址相关信息
        "STATIONNAME": {
            "cn_name": "站址名称",
        },
        "STATIONCODE": {
            "cn_name": "运维ID",
        },
        "STATIONSTATUS": {
            "cn_name": "告警发生时站址状态",
            "enum": {"1": "工程", "2": "交维", "3": "退网下线", "6": "测试", "7": "存量监控", "8": "检修"},
        },
        "STATIONSERVICESTATUS": {
            "cn_name": "告警上报时的站址维护状态",
        },
        # 人工操作相关字段
        "CONFIRMPERSON": {
            "cn_name": "确认人",
        },
        "CONFIRMINFO": {
            "cn_name": "确认原因",
        },
        "USER_REMARK": {
            "cn_name": "用户备注信息",
        },
        # 其他标识字段
        "SOURCE": {
            "cn_name": "告警来源",
            "enum": {"1": "动环监控设备", "2": "系统", "8": "设备管理中台"},
        },
        "SERIALNO": {
            "cn_name": "告警流水号",
        },
        "ALARMLASTED": {
            "cn_name": "首次告警时间至目前经过时间（分）",
        },
        "MID": {
            "cn_name": "信号量",
        },
        "SITESOURCE": {
            "cn_name": "站址来源",
            "enum": {"7": "自有站", "8": "社会站"},
        },
        "TASKSTATUS": {
            "cn_name": "任务状态",
        },
        "MAINTENANCEROOMPERSONID": {
            "cn_name": "维修室人员编号",
        },
        "BILLID": {
            "cn_name": "工单编号",
        },
        "BILL_STATUS": {
            "cn_name": "工单状态",
        },
        "FSUNAME": {
            "cn_name": "FSU名称",
        }
    }

    sql = """
          SELECT *
          FROM ITOWER.V_IF_PROV_ALARMAT_GXFGS
          WHERE STATIONSTATUS = :status \
          """

    with _get_oracle_connection() as conn:
        df = pd.read_sql(sql, conn, params={"status": "2"})

    # 清理 Oracle NULL
    ORACLE_NULL = getattr(cx_Oracle, 'NULL', None)
    df = df.where(pd.notna(df) & (df != ORACLE_NULL), None)

    # 应用字段映射
    df = _apply_mapping(df, ALARM_FIELD_MAPPING)

    return df


def fetch_his_alarm(
        alarm_name="一级低压脱离告警",
        start_time=None,
        end_time=None
):
    HIS_ALARM_MAPPING = {
        # 基础标识字段
        "ID": {
            "cn_name": "ID",
        },
        "CATEGORY": {
            "cn_name": "告警类别",
            "enum": {"1": "动环告警", "2": "其它告警"}
        },
        "OBJTYPE": {
            "cn_name": "告警对象类型",
            "enum": {"1": "FSU", "2": "其它类型"}
        },
        "SUBOBJTYPE": {
            "cn_name": "告警子对象类型",
            "enum": {"1": "动环设备", "2": "其它类型"}
        },
        "ALARMTYPE": {
            "cn_name": "告警类型",
            "enum": {"1": "设备告警", "2": "其它告警"}
        },
        # 对象标识字段
        "OBJID": {
            "cn_name": "告警对象ID",
        },
        "OBJNAME": {
            "cn_name": "站址",
        },
        "SUBOBJID": {
            "cn_name": "告警子对象ID",
        },
        "SUBOBJNAME": {
            "cn_name": "告警子对象名称",
        },
        # 告警等级相关
        "ORIGLEVEL": {
            "cn_name": "原始告警等级",
            "enum": {"1": "一级告警", "2": "二级告警", "3": "三级告警", "4": "四级告警"}
        },
        "ALARMLEVEL": {
            "cn_name": "告警等级",
            "enum": {"1": "一级告警", "2": "二级告警", "3": "三级告警", "4": "四级告警"}
        },
        # 告警描述信息
        "CAUSE": {
            "cn_name": "告警原因",
        },
        "CODE": {
            "cn_name": "告警摘要",
        },
        "DETAILINFO": {
            "cn_name": "告警详情",
        },
        "EXTRAINFO": {
            "cn_name": "告警附加信息",
        },
        "ALARMNAME": {
            "cn_name": "告警名称",
        },
        "ALARMEXPLAIN": {
            "cn_name": "信号量说明",
        },
        "FAULTEXPLAIN": {
            "cn_name": "信号量解释",
        },
        # 时间相关字段
        "INSERTTIME": {
            "cn_name": "告警入库时间",
        },
        "FIRSTSYSTEMTIME": {
            "cn_name": "首次告警的系统时间",
        },
        "LASTSYSTEMTIME": {
            "cn_name": "末次告警的系统时间",
        },
        "CONFIRMTIME": {
            "cn_name": "确认时间",
        },
        "FIRSTALARMTIME": {
            "cn_name": "首次告警时间",
        },
        "LASTALARMTIME": {
            "cn_name": "末次告警时间",
        },
        "CLEARTIME": {
            "cn_name": "告警清除时间",
        },
        "CLEARSYSTEMTIME": {
            "cn_name": "告警清除系统时间",
        },
        # 状态标识字段
        "TIMES": {
            "cn_name": "告警重复次数",
        },
        "IFCONFIRM": {
            "cn_name": "是否已确认",
            "enum": {"Y": "已确认", "N": "未确认"},
        },
        "IFRECOVER": {
            "cn_name": "是否已恢复",
            "enum": {"Y": "已恢复", "N": "未恢复"},
        },
        "IFCONVERTFAULT": {
            "cn_name": "是否转故障",
            "enum": {"Y": "已转故障", "N": "未转故障"},
        },
        "IFATTENTION": {
            "cn_name": "是否关注",
            "enum": {"Y": "已关注", "N": "未关注"},
        },
        # 关联关系字段
        "FAULTID": {
            "cn_name": "故障单ID",
        },
        "STATIONID": {
            "cn_name": "运维监控站址ID",
        },
        "PARENT_ALARMID": {
            "cn_name": "所属主要告警ID",
        },
        "PROVINCEID": {
            "cn_name": "省",
            "enum": {"0098364": "广西"},

        },
        "CITYID": {
            "cn_name": "地市",
            "enum": {"0099977": "南宁",
                     "0099978": "柳州",
                     "0099979": "桂林",
                     "0099980": "玉林",
                     "0099981": "贵港",
                     "0099982": "百色",
                     "0099983": "河池",
                     "0099984": "钦州",
                     "0099985": "梧州",
                     "0099986": "北海",
                     "0099987": "防城港",
                     "0099988": "崇左",
                     "0099989": "来宾",
                     "0099990": "贺州"
                     },
        },
        # 站址相关信息
        "STATIONCODE": {
            "cn_name": "运维监控站址编码",
        },
        "STATIONSTATUS": {
            "cn_name": "告警上报时FUS状态",
            "enum": {"1": "工程", "2": "交维", "3": "退网下线", "6": "测试", "7": "存量监控", "8": "检修"},
        },
        "STATIONSERVICESTATUS": {
            "cn_name": "告警上报时的站址维护状态",
        },
        # 人工操作相关字段
        "CONFIRMPERSON": {
            "cn_name": "确认人",
        },
        "CONFIRMINFO": {
            "cn_name": "确认原因",
        },
        "USER_REMARK": {
            "cn_name": "用户备注信息",
        },
        "CLEARPERSON": {
            "cn_name": "告警清除人",
        },
        "CLEARCAUSE": {
            "cn_name": "告警清除原因",
        },
        # 其他标识字段
        "SOURCE": {
            "cn_name": "告警来源",
        },
        "SERIALNO": {
            "cn_name": "告警流水号",
        },
        "CLEARTYPE": {
            "cn_name": "告警清除方式",
            "enum": {"1": "自动清除", "2": "人工清除", "": "人工清除"},
        },
        "MID": {
            "cn_name": "信号量",
        }
    }
    """
    获取历史告警数据

    参数:
        alarm_name (str): 告警名称关键词，默认"一级低压脱离告警"
        start_time (str or datetime): 首次告警时间起始，格式如 "2025-01-01 00:00:00"
        end_time (str or datetime): 末次告警时间截止

    返回:
        pandas.DataFrame: 清洗并映射后的数据
    """
    base_sql = """
               SELECT *
               FROM ITOWER.V_IF_PROV_ALARM_GXFGS
               WHERE STATIONSTATUS = :status \
               """
    params = {"status": "2"}

    if alarm_name:
        base_sql += " AND ALARMNAME LIKE '%' || :alarm_name || '%'"
        params["alarm_name"] = alarm_name

    if start_time:
        if isinstance(start_time, datetime):
            start_time = start_time.strftime("%Y-%m-%d %H:%M:%S")
        base_sql += " AND FIRSTALARMTIME >= TO_DATE(:start_time, 'YYYY-MM-DD HH24:MI:SS')"
        params["start_time"] = start_time

    if end_time:
        if isinstance(end_time, datetime):
            end_time = end_time.strftime("%Y-%m-%d %H:%M:%S")
        base_sql += " AND LASTALARMTIME <= TO_DATE(:end_time, 'YYYY-MM-DD HH24:MI:SS')"
        params["end_time"] = end_time

    with _get_oracle_connection() as conn:
        df = pd.read_sql(base_sql, conn, params=params)

    # 清理 NULL
    ORACLE_NULL = getattr(cx_Oracle, 'NULL', None)
    df = df.where(pd.notna(df) & (df != ORACLE_NULL), None)

    # 应用字段映射
    df = _apply_mapping(df, HIS_ALARM_MAPPING)

    return df

def main():
    df1 = fetch_alarmat()
    column_mapping = {
        '地市': 'city',
        '站址名称': 'site_name',
        '告警名称': 'alarm_name',
        '首次告警的系统时间': 'happen_time',
        '首次告警时间至目前经过时间（分）': 'broken_time',
        '告警对象ID':'site_maitan_code'
    }
    df1 = df1[list(column_mapping.keys())].rename(columns=column_mapping)
    query = text("SELECT area, village, countryside, site_code, belong,site_maitan_code FROM station")
    station = pd.read_sql_query(query, con=sql_orm().engine)
    df1 = df1.merge(station,on='site_maitan_code', how='left')
    df1 = df1.where(pd.notnull(df1), None)

    sql_orm().truncate_add_data(df1,"data_alarm_now")

    # 获取本月一级低压脱离告警
    now = datetime.now()
    start_of_month = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    df2 = fetch_his_alarm(alarm_name="一级低压脱离告警", start_time=start_of_month, end_time=now)
    df2.to_csv(settings.resolve_path("spider/down/oracle_data/data_alarm_his_out.csv"))
