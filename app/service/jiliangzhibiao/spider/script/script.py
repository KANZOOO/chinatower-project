import sys
import os

# 添加项目根目录到 Python 路径
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from app.service.jiliangzhibiao.spider.spider import download_shunt_meter_excel, down_file
from app.service.jiliangzhibiao.spider.schema.schema_jilianghzibiao import yidong_jiliang,yidong_kaiguan,yidong_jiliang5g,liantong_jiliang,liantong_kaiguan,liantong_jiliang5g,dianxin_jiliang,dianxin_kaiguan,dianxin_jiliang5g
from core.config import settings

class Jiliangzhibiao:
    def __init__(self):
        self.path_yidong_jiliang=settings.resolve_path("app/service/jiliangzhibiao/down/分路计量设备-移动租户电流.xls")
        self.path_yidong_kaiguan=settings.resolve_path("app/service/jiliangzhibiao/down/开关电源-移动租户电流.xls")
        self.path_yidong_jiliang5g=settings.resolve_path("app/service/jiliangzhibiao/down/分路计量设备-移动租户电流（5G）.xls")
        self.path_liantong_jiliang=settings.resolve_path("app/service/jiliangzhibiao/down/分路计量设备-联通租户电流.xls")
        self.path_liantong_kaiguan=settings.resolve_path("app/service/jiliangzhibiao/down/开关电源-联通租户电流.xls")
        self.path_liantong_jiliang5g=settings.resolve_path("app/service/jiliangzhibiao/down/分路计量设备-联通租户电流（5G）.xls")
        self.path_dianxin_jiliang=settings.resolve_path("app/service/jiliangzhibiao/down/分路计量设备-电信租户电流.xls")
        self.path_dianxin_kaiguan=settings.resolve_path("app/service/jiliangzhibiao/down/开关电源-电信租户电流.xls")
        self.path_dianxin_jiliang5g=settings.resolve_path("app/service/jiliangzhibiao/down/分路计量设备-电信租户电流（5G）.xls")
    def down(self,url,data,path):
        for key, into_data in data.items():
            if key in ['INTO_DATA_1', 'INTO_DATA_2', 'INTO_DATA_FINAL']:
                data[key]['queryForm:queryAccessStatus'] = "02"
        down_file(url=url,
                  data=data,
                  path=path)

    def run_down(self):
        # 分路租户异常
        # 分路计量设备-移动租户电流数据 URL
        url1 = "http://omms.chinatowercom.cn:9000/business/resMge/pwMge/performanceMge/perfdata.xhtml"
        self.down(url=url1, data=yidong_jiliang, path=self.path_yidong_jiliang)
        print(f"✅ 分路计量设备-移动租户电流数据下载成功")
        # 开关电源-移动租户电流数据 URL
        self.down(url=url1, data=yidong_kaiguan, path=self.path_yidong_kaiguan)
        print(f"✅ 开关电源-移动租户电流数据下载成功")
        # 分路计量设备-移动租户电流（5G）数据 URL
        self.down(url=url1, data=yidong_jiliang5g, path=self.path_yidong_jiliang5g)
        print(f"✅ 分路计量设备-移动租户电流（5G）数据下载成功")
        # 分路计量设备-联通租户电流数据 URL
        self.down(url=url1, data=liantong_jiliang, path=self.path_liantong_jiliang)
        print(f"✅ 分路计量设备-联通租户电流数据下载成功")
        # 开关电源-联通租户电流数据 URL
        self.down(url=url1, data=liantong_kaiguan, path=self.path_liantong_kaiguan)
        print(f"✅ 开关电源-联通租户电流数据下载成功")
        # 分路计量设备-联通租户电流（5G）数据 URL
        self.down(url=url1, data=liantong_jiliang5g, path=self.path_liantong_jiliang5g)
        print(f"✅ 分路计量设备-联通租户电流（5G）数据下载成功")
        # 分路计量设备-电信租户电流数据 URL
        self.down(url=url1, data=dianxin_jiliang, path=self.path_dianxin_jiliang)
        print(f"✅ 分路计量设备-电信租户电流数据下载成功")
        # 开关电源-电信租户电流数据 URL
        self.down(url=url1, data=dianxin_kaiguan, path=self.path_dianxin_kaiguan)
        print(f"✅ 开关电源-电信租户电流数据下载成功")
        # 分路计量设备-电信租户电流（5G）数据 URL
        self.down(url=url1, data=dianxin_jiliang5g, path=self.path_dianxin_jiliang5g)
        print(f"✅ 分路计量设备-电信租户电流（5G）数据下载成功")

    def process(self):
        pass

def main():
    ji_liang_zhibiao= Jiliangzhibiao()
    ji_liang_zhibiao.run_down()
    ji_liang_zhibiao.process()
    download_shunt_meter_excel()

if __name__ == '__main__':
    main()