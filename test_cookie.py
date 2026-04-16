import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.sql import sql_orm
from core.config import settings

print("=" * 60)
print("📊 数据库连接测试")
print("=" * 60)

print(f"\n📁 项目根目录：{settings.index}")
print(f"📄 .env 文件路径：{settings.model_config['env_file']}")

print(f"\n🔌 数据库配置：")
print(f"   Host: {settings.db_host}")
print(f"   Port: {settings.db_port}")
print(f"   User: {settings.db_user}")
print(f"   DB Name: {settings.db_name}")

try:
    print("\n🔄 正在连接数据库...")
    db = sql_orm()
    print("✅ 数据库连接成功！")
    
    print("\n🔍 测试获取Cookie (ID=dw.rj.fengsw)...")
    result = db.get_cookies("dw.rj.fengsw")
    
    print(f"\n✅ 获取成功！")
    print(f"   返回类型：{type(result)}")
    print(f"   包含的键：{list(result.keys())}")
    
    if "cookies" in result:
        print(f"\n🍪 Cookie 字典：")
        print(f"   类型：{type(result['cookies'])}")
        print(f"   键值对数量：{len(result['cookies'])}")
        for k, v in list(result['cookies'].items())[:3]:
            print(f"   {k} = {v[:50]}...")
    
    if "cookies_str" in result:
        print(f"\n📝 Cookie 字符串：")
        print(f"   长度：{len(result['cookies_str'])}")
        print(f"   前100字符：{result['cookies_str'][:100]}...")
    
except Exception as e:
    print(f"\n❌ 错误：{e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 60)