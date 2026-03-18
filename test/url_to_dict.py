from urllib.parse import parse_qs, unquote
from pprint import pprint

data =['AJAXREQUEST=_viewRoot&j_id1499=j_id1499&autoScroll=&javax.faces.ViewState=j_id25&j_id1499%3Aj_id1537=j_id1499%3Aj_id1537&hiddenGroup=6&hiddenClass=undefined&AJAX%3AEVENTS_COUNT=1&']
# 解析查询字符串为字典
for query_string in data:
    decoded_query_string = unquote(query_string)
    parsed_dict = parse_qs(decoded_query_string, keep_blank_values=True)
    INTO_DATA = {k: v[0] if v else '' for k, v in parsed_dict.items()}
    pprint(INTO_DATA)
