import pandas as pd
import json
import requests
from concurrent.futures import ThreadPoolExecutor

map={}
def hitapi(order_id):
    url = "https://api.crofarm.com/spa/consumer/order/invoice/v1/?order_id={}".format(order_id)
    headers = {
        "Content-type": "application/json",
        'access-token':'ca2c76dda6bbfebce43416658d736b83ac2456c9b9d07368ede730358d444f55'
    }
    preview_response = requests.get(url, headers=headers)
    preview_response = json.loads(preview_response.text)

    if('link' in preview_response):
            print("{}  ->  {}".format(order_id,preview_response['link']))
            map[order_id]=preview_response['link']

    else:
        print(preview_response)
        map[order_id]=''

def make_clickable(val):
     return '<a target="_blank" href="{}">{}</a>'.format(val,val)
    
out=pd.read_excel("file_name.xlsx")
# print(out.columns)
# loc = out.columns.get_loc("order_id")

order_id=out.loc[:,'Order ID']

with ThreadPoolExecutor(max_workers=len(order_id)) as exe:
    exe.map(hitapi,order_id)
    

invoice=[map[i] for i in order_id]
out['invoice']=invoice
print(out)
# out.to_excel("file_name.xlsx",index=False)
