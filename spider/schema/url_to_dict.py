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
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3AunitHidden=&queryForm%3AqueryRoomIdHidden=&queryForm%3AqueryStationIdHidden=&queryForm%3Aj_id11=&queryForm%3Aj_id15=&queryForm%3Aj_id19=&queryForm%3Aj_id23=&queryForm%3Aj_id27=&queryForm%3Aj_id31=&queryForm%3AqueryAccessStatus_hiddenValue=02%2C03&queryForm%3AqueryAccessStatus=02&queryForm%3AqueryAccessStatus=03&queryForm%3Aj_id38=&queryForm%3AqueryEdgeNode_hiddenValue=&queryForm%3Aj_id45=&queryForm%3Aj_id49=&queryForm%3Aj_id53=&queryForm%3Aj_id57=&queryForm%3Aj_id61=&queryForm%3Aj_id65=&queryForm%3Aj_id69=&queryForm%3Aj_id73=&queryForm%3AqueryProjectCode=&queryForm%3AqueryProjectName=&queryForm%3Aj_id80=&queryForm%3Aj_id84=&queryForm%3AqueryCrewAreaId=&queryForm%3AqueryCrewAreaName=&queryForm%3Aj_id91=&queryForm%3Aj_id95=&queryForm%3Aj_id99=&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&javax.faces.ViewState=j_id6&queryForm%3Aj_id103=queryForm%3Aj_id103&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3AunitHidden=&queryForm%3AqueryRoomIdHidden=&queryForm%3AqueryStationIdHidden=&queryForm%3Aj_id11=&queryForm%3Aj_id15=&queryForm%3Aj_id19=&queryForm%3Aj_id23=&queryForm%3Aj_id27=&queryForm%3Aj_id31=&queryForm%3AqueryAccessStatus_hiddenValue=02%2C03&queryForm%3AqueryAccessStatus=02&queryForm%3AqueryAccessStatus=03&queryForm%3Aj_id38=&queryForm%3AqueryEdgeNode_hiddenValue=&queryForm%3Aj_id45=&queryForm%3Aj_id49=&queryForm%3Aj_id53=&queryForm%3Aj_id57=&queryForm%3Aj_id61=&queryForm%3Aj_id65=&queryForm%3Aj_id69=&queryForm%3Aj_id73=&queryForm%3AqueryProjectCode=&queryForm%3AqueryProjectName=&queryForm%3Aj_id80=&queryForm%3Aj_id84=&queryForm%3AqueryCrewAreaId=&queryForm%3AqueryCrewAreaName=&queryForm%3Aj_id91=&queryForm%3Aj_id95=&queryForm%3Aj_id99=&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&javax.faces.ViewState=j_id6&queryForm%3Aj_id104=queryForm%3Aj_id104&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id652=j_id652&j_id652%3AexportBtn=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id6',
]

# 第二个网页的URL查询字符串列表、替换为你实际的URL即可
data2 = [
    #1
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3AunitHidden=&queryForm%3AqueryRoomIdHidden=&queryForm%3AqueryStationIdHidden=&queryForm%3Aj_id11=&queryForm%3Aj_id15=&queryForm%3Aj_id19=&queryForm%3Aj_id23=&queryForm%3Aj_id27=&queryForm%3Aj_id31=&queryForm%3AqueryAccessStatus_hiddenValue=02%2C03&queryForm%3AqueryAccessStatus=02&queryForm%3AqueryAccessStatus=03&queryForm%3Aj_id38=&queryForm%3AqueryEdgeNode_hiddenValue=&queryForm%3Aj_id45=&queryForm%3Aj_id49=&queryForm%3Aj_id53=&queryForm%3Aj_id57=&queryForm%3Aj_id61=&queryForm%3Aj_id65=&queryForm%3Aj_id69=&queryForm%3Aj_id73=&queryForm%3AqueryProjectCode=&queryForm%3AqueryProjectName=&queryForm%3Aj_id80=&queryForm%3Aj_id84=&queryForm%3AqueryCrewAreaId=&queryForm%3AqueryCrewAreaName=&queryForm%3Aj_id91=&queryForm%3Aj_id95=&queryForm%3Aj_id99=&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&javax.faces.ViewState=j_id6&queryForm%3Aj_id103=queryForm%3Aj_id103&',
    #2
    'AJAXREQUEST=_viewRoot&queryForm=queryForm&queryForm%3AunitHidden=&queryForm%3AqueryRoomIdHidden=&queryForm%3AqueryStationIdHidden=&queryForm%3Aj_id11=&queryForm%3Aj_id15=&queryForm%3Aj_id19=&queryForm%3Aj_id23=&queryForm%3Aj_id27=&queryForm%3Aj_id31=&queryForm%3Aj_id35=&queryForm%3Aj_id39=&queryForm%3Aj_id43=&queryForm%3Aj_id47=&queryForm%3Aj_id51=&queryForm%3Aj_id55=&queryForm%3Aj_id59=&queryForm%3Aj_id63=&queryForm%3Aj_id67=&queryForm%3AcurrPageObjId=0&queryForm%3ApageSizeText=35&javax.faces.ViewState=j_id7&queryForm%3Aj_id72=queryForm%3Aj_id72&AJAX%3AEVENTS_COUNT=1&',
    #3
    'j_id362=j_id362&j_id362%3AexportBtn=%E5%85%A8%E9%83%A8&javax.faces.ViewState=j_id7'
]


parse_url_to_schema(data, "第一个网页")

parse_url_to_schema(data2, "第二个网页")