import os
import pickle
from datetime import datetime, timedelta
from core.config import settings


class PowerStationSpider:
    def __init__(self):
        """初始化发电费用爬虫"""
        self.pickle_quxin = settings.resolve_path("app/service/fadianfeiyong/spider/down/pickle_quxin.pkl")
        self.output_dir = settings.resolve_path("app/service/fadianfeiyong/down")
        self.session_quxin = None
        # 确保输出目录存在
        os.makedirs(self.output_dir, exist_ok=True)
        print(f"📁 发电费用爬虫初始化")
        print(f"   Session文件路径：{self.pickle_quxin}")
        print(f"   输出目录路径：{self.output_dir}")
        
    def load_session(self):
        """从 pickle 文件加载 session"""
        if not os.path.exists(self.pickle_quxin):
            raise FileNotFoundError(f"Session文件不存在：{self.pickle_quxin}，请先运行登录脚本")
        
        print(f"\n📂 正在加载Session...")
        with open(self.pickle_quxin, "rb") as f:
            self.session_quxin = pickle.load(f)
        print(f"✅ Session加载成功")
        return self.session_quxin
    
    def export_report(self, begin_time: datetime, end_time: datetime, file_name: str):
        """
        按指定时间范围导出发电站报表
        """
        if self.session_quxin is None:
            self.load_session()

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

        output_path = os.path.join(self.output_dir, file_name)
        # 保存文件
        with open(output_path, "wb") as f:
            f.write(res.content)

        print(f"✅ 发电站报表下载完成：{output_path}")
        return output_path

    def export_two_reports(self):
        """
        导出两个报表：
        1. 前一天 00:00:00 ~ 23:59:59
        2. 当月1号 00:00:00 ~ 前一天 23:59:59
        """
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)

        # 报表1：前一天整天
        yesterday_begin = datetime.combine(yesterday, datetime.min.time())
        yesterday_end = datetime.combine(yesterday, datetime.max.time()).replace(microsecond=0)
        yesterday_file_name = f"发电站报表_前一天_{yesterday.strftime('%Y%m%d')}.xlsx"
        path_yesterday = self.export_report(yesterday_begin, yesterday_end, yesterday_file_name)

        # 报表2：当月1号到前一天
        month_begin_date = today.replace(day=1)
        month_begin = datetime.combine(month_begin_date, datetime.min.time())
        month_end = yesterday_end
        month_file_name = (
            f"发电站报表_当月累计_{month_begin_date.strftime('%Y%m%d')}"
            f"_{yesterday.strftime('%Y%m%d')}.xlsx"
        )
        path_month = self.export_report(month_begin, month_end, month_file_name)

        return [path_yesterday, path_month]


def main():
    """主函数"""
    print("="*60)
    print("🏃 发电站报表下载脚本启动")
    print("="*60)
    try:
        spider = PowerStationSpider()
        result_paths = spider.export_two_reports()
        print("📄 生成文件：")
        for path in result_paths:
            print(f"   - {path}")
        print("\n" + "="*60)
        print("🎉 任务完成！")
        print("="*60)
    except Exception as e:
        print(f"\n❌ 执行失败：{e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
