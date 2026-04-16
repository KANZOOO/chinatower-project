import os
import pickle
from datetime import datetime, timedelta
from requests import Session
from core.config import settings


class PowerStationSpider:
    def __init__(self):
        """初始化发电费用爬虫"""
        self.pickle_quxin = settings.resolve_path("spider/down/pickle_quxin.pkl")
        self.gen_path = settings.resolve_path("app/service/fadianfeiyong/down/发电站报表.xlsx")
        self.session_quxin = None
        # 确保输出目录存在
        os.makedirs(os.path.dirname(self.gen_path), exist_ok=True)
        print(f"📁 发电费用爬虫初始化")
        print(f"   Session文件路径：{self.pickle_quxin}")
        print(f"   输出文件路径：{self.gen_path}")
        
    def load_session(self):
        """从 pickle 文件加载 session"""
        if not os.path.exists(self.pickle_quxin):
            raise FileNotFoundError(f"Session文件不存在：{self.pickle_quxin}，请先运行登录脚本")
        
        print(f"\n📂 正在加载Session...")
        with open(self.pickle_quxin, "rb") as f:
            self.session_quxin = pickle.load(f)
        print(f"✅ Session加载成功")
        return self.session_quxin
    
    def power_station_report(self, days=1):
        """
        获取发电站报表
        
        Args:
            days: 查询天数，默认查询最近1天
        """
        if self.session_quxin is None:
            self.load_session()
        
        # 计算时间范围
        end_time = datetime.now()
        begin_time = end_time - timedelta(days=days)
        
        print(f"\n📅 查询时间范围：{begin_time.strftime('%Y-%m-%d')} ~ {end_time.strftime('%Y-%m-%d')}")
        
        # 请求数据
        data = {
            "approvalOfDispatchId": "",
            "area.id": "",
            "area.name": "",
            "asOper": "",
            "beginGenerateDate": begin_time.strftime("%Y-%m-%d %H:%M:%S"),
            "city.id": "",
            "city.name": "",
            "collectorCode": "",
            "endGenerateDate": end_time.strftime("%Y-%m-%d %H:%M:%S"),
            "finishConfigId": "",
            "generateOfficeName": "",
            "generatePowerState": "1",
            "isStart": "",
            "number": "",
            "orderBy": "",
            "pageNo": "1",
            "pageSize": "25",
            "stationCode": "",
            "stationName": "",
            "stopBegeinDate": "",
            "stopEndDate": "",
            "typeCode": "",
            "urlPara": "",
        }
        
        # 请求URL（使用和登录相同的域名）
        url = "http://clound.gxtower.cn:11080/tower_manage_bms/a/export"
        
        print(f"🔄 正在请求数据...")
        res = self.session_quxin.post(url=url, data=data, timeout=60)
        print(f"✅ 请求完成，状态码：{res.status_code}")
        
        # 保存文件
        with open(self.gen_path, "wb") as f:
            f.write(res.content)
        
        print(f"✅ 发电站报表下载完成：{self.gen_path}")
        return self.gen_path


def main():
    """主函数"""
    print("="*60)
    print("🏃 发电站报表下载脚本启动")
    print("="*60)
    try:
        spider = PowerStationSpider()
        spider.power_station_report(days=1)
        print("\n" + "="*60)
        print("🎉 任务完成！")
        print("="*60)
    except Exception as e:
        print(f"\n❌ 执行失败：{e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()