# app/deps.py 或直接在 main.py 中
from fastapi import Request, HTTPException
from starlette.status import HTTP_403_FORBIDDEN

# ===== 全局默认白名单（可被路由覆盖）=====
GLOBAL_ALLOWED_IPS = {
    '116.8.36.226', '61.139.230.157', '61.139.230.181', '171.110.76.107',
    '222.84.213.17', '171.107.219.82', '219.159.146.53', '117.140.123.79',
    '116.10.6.134', '106.127.124.209', '117.140.123.206', '116.10.1.166',
    '113.16.187.43', '10.8.3.32', '116.1.248.102', '116.10.5.2',
    '180.139.165.224', '117.140.61.84', '111.59.241.99', '117.183.114.127',
    '222.84.227.242', '113.17.102.211', '116.253.163.53', '116.9.189.183',
    '222.83.251.174', '113.15.110.114', '222.83.251.171', '117.141.80.223',
    '111.59.123.253', '111.12.120.207', '221.7.183.44', '222.83.251.170',
    '117.141.73.61', '171.109.118.158', '111.58.194.137', '116.10.6.123',
    '222.84.119.210', '222.83.251.173', '113.15.231.142', '182.90.255.99',
    '111.58.97.254', '58.59.135.226', '100.106.58.145', '202.103.217.181',
    '171.109.242.25', '117.141.250.11', '171.109.118.155', '180.141.32.202',
    '171.108.137.122', '219.159.250.25', '219.159.146.80', '113.17.90.142',
    '172.16.30.69', '117.140.214.145', '113.17.98.172', '113.16.135.199',
    '171.111.245.54', '221.7.154.190', '172.16.30.114', '117.140.61.192',
    '222.83.251.172', '222.216.152.201', '171.107.144.114', '171.107.101.209',
    '111.59.192.55'
}

def get_client_ip(request: Request) -> str:
    """获取真实客户端 IP"""
    xff = request.headers.get("x-forwarded-for")
    if xff:
        return xff.split(",")[0].strip()
    return request.client.host if request.client else "127.0.0.1"

def require_ip_whitelist(allowed_ips=None):
    """
    返回一个依赖函数，用于限制 IP 访问
    :param allowed_ips: 允许的 IP 集合。若为 None，则使用全局白名单；若为 empty set/list，则公开。
    """
    def dependency(request: Request):
        if allowed_ips == [] or allowed_ips == set():  # 显式公开
            return

        client_ip = get_client_ip(request)

        ips_to_check = allowed_ips if allowed_ips is not None else GLOBAL_ALLOWED_IPS

        if client_ip not in ips_to_check:
            print(f"❌ 拦截非法 IP 访问: {client_ip} -> {request.url.path}")
            raise HTTPException(status_code=HTTP_403_FORBIDDEN, detail="IP not allowed")

    return dependency