from urllib.parse import parse_qs, unquote
from pprint import pprint

def parse_url_to_schema(query_string_list, page_name):
    for idx, query_string in enumerate(query_string_list, 1):
        print(f"\n--- {page_name} 第{idx}条URL解析结果 ---")
        decoded_query_string = unquote(query_string)
        parsed_dict = parse_qs(decoded_query_string, keep_blank_values=True)
        # 转换为单值字典（原逻辑）
        into_data = {k: v[0] if v else '' for k, v in parsed_dict.items()}
        pprint(into_data)

# 第一个网页的URL查询字符串列表、替换为你实际的URL即可
data = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%A7%BB%E5%8A%A8%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0445101001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id7&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%A7%BB%E5%8A%A8%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0445101001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id7&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id7',
]

# 第二个网页的URL查询字符串列表、替换为你实际的URL即可
data2 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%A7%BB%E5%8A%A8%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0406135001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%A7%BB%E5%8A%A8%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0406135001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id24',
]

# 第三个网页的URL查询字符串列表、替换为你实际的URL即可
data3 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%A7%BB%E5%8A%A8%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81%EF%BC%885G%EF%BC%89&queryForm%3Amid=0445112001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%A7%BB%E5%8A%A8%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81%EF%BC%885G%EF%BC%89&queryForm%3Amid=0445112001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id24',
]

data4 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E8%81%94%E9%80%9A%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0445103001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E8%81%94%E9%80%9A%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0445103001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id24',
]
data5 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E8%81%94%E9%80%9A%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0406137001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E8%81%94%E9%80%9A%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0406137001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id24&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id24',
]
data6 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E8%81%94%E9%80%9A%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81%EF%BC%885G%EF%BC%89&queryForm%3Amid=0445114001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E8%81%94%E9%80%9A%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81%EF%BC%885G%EF%BC%89&queryForm%3Amid=0445114001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id30',
]
data7 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%94%B5%E4%BF%A1%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0445105001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%94%B5%E4%BF%A1%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0445105001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id30',
]
data8 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%94%B5%E4%BF%A1%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0406139001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%94%B5%E4%BF%A1%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81&queryForm%3Amid=0406139001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id30',
]
data9 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%94%B5%E4%BF%A1%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81%EF%BC%885G%EF%BC%89&queryForm%3Amid=0445116001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id52=queryForm%3Aj_id52&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3Aaid=&queryForm%3Amongositecode=&queryForm%3AsiteProvinceId=0098364&queryForm%3AqueryFlag=&queryForm%3AunitHidden1=&queryForm%3AunitHidden2=0099977&queryForm%3AunitHidden3=0098364&queryForm%3AunitTypeHidden=undefined&queryForm%3AsiteNameId=&queryForm%3AdeviceName=&queryForm%3Adid=&queryForm%3AmidName=%E7%94%B5%E4%BF%A1%E7%A7%9F%E6%88%B7%E7%94%B5%E6%B5%81%EF%BC%885G%EF%BC%89&queryForm%3Amid=0445116001&queryForm%3AqueryStationId=&queryForm%3AqueryStationIdShow=&queryForm%3AqueryFsuId=&queryForm%3AmidType=%E9%81%A5%E6%B5%8B&queryForm%3AquerySpeId=&queryForm%3AquerySpeIdShow=&queryForm%3AstarttimeInputDate=&queryForm%3AstarttimeInputCurrentDate=03%2F2026&queryForm%3AendtimeInputDate=&queryForm%3AendtimeInputCurrentDate=04%2F2026&queryForm%3AquerySiteSourceCode=&queryForm%3AifRestrict=true&queryForm%3AcurrPageObjId=1&queryForm%3ApageSizeText=35&queryForm%3ApanelOpenedState=&javax.faces.ViewState=j_id30&queryForm%3Aj_id53=queryForm%3Aj_id53&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id421=j_id421&j_id421%3Aj_id423=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id30',
]
data10 = [
    #1
    '',
    #2
    '',
    #3
    '',
]





parse_url_to_schema(data, "第一个网页")
parse_url_to_schema(data2, "第二个网页")
parse_url_to_schema(data3, "第三个网页")
parse_url_to_schema(data4, "第四个网页")
parse_url_to_schema(data5, "第五个网页")
parse_url_to_schema(data6, "第六个网页")
parse_url_to_schema(data7, "第七个网页")
parse_url_to_schema(data8, "第八个网页")
parse_url_to_schema(data9, "第九个网页")
parse_url_to_schema(data9, "第十个网页")