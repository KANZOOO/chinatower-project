import os
from core.config import settings
# from spider.config.data_edge_x import data
from spider.script.model import down_file
from spider.script.gateway_process import process_gateway_main_list
from spider.script.monitor_process import process_camera_main_list
from spider.schema.schema import zhineng_wangguan, zhineng_shexiangtou


class ZhiLianTongBao:
    def __init__(self):
        self.path_wangguan=settings.resolve_path("spider/script/station/down/边缘网关.xls")
        self.path_shexiangtou=settings.resolve_path("spider/script/station/down/边缘摄像头.xls")
    def down(self,url,data,path):
        for key, into_data in data.items():
            if key in ['INTO_DATA_1', 'INTO_DATA_2', 'INTO_DATA_FINAL']:
                data[key]['queryForm:queryAccessStatus'] = "02"
        down_file(url=url,
                  data=data,
                  path=path)

    def run_down(self):
        # 修正：需要从 schema 中获取正确的 URL 字符串
        # 网关数据下载 URL
        wangguan_url = "http://omms.chinatowercom.cn:9000/business/resMge/devMge/listEdgeGateway.xhtml"
        self.down(url=wangguan_url, data=zhineng_wangguan, path=self.path_wangguan)
        print(f"✅ 网关数据下载成功")
        # 摄像头数据下载 URL
        shexiangtou_url = "http://omms.chinatowercom.cn:9000/business/resMge/devMge/listEdgeGatewayDev.xhtml"
        self.down(url=shexiangtou_url, data=zhineng_shexiangtou, path=self.path_shexiangtou)
        print(f"✅ 摄像头数据下载成功")
    def process(self):
        """整合网关+摄像头数据处理流程（仅保留调试用的核心逻辑）"""
        # 1. 执行网关数据处理
        gateway_result = process_gateway_main_list()
        if gateway_result:
            print("✅ 网关数据处理完成")
        camera_result = process_camera_main_list()
        if camera_result:
            print("✅ 摄像头数据处理完成")

def main():
    zhi_lian_tongbao = ZhiLianTongBao()
    zhi_lian_tongbao.run_down()
    zhi_lian_tongbao.process()

if __name__ == '__main__':
    main()