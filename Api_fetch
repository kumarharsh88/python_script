# import pandas lib as pd
import pandas as pd
import requests

# read by default 1st sheet of an excel file
df = pd.read_excel('file_name.xlsx')
print(df)
order_id_details=df.loc[:,'Order ID']

invoice=[]
for id in order_id_details:
    headers={'access-token':"ca2c76dda6bbfebce43416658d736b83ac2456c9b9d07368ede730358d444f55"}
    response = requests.get("https://api.crofarm.com/spa/consumer/order/invoice/v1/?order_id={}".format(id),headers=headers)
    invoice_detail=response.json()
    invoice.append(invoice_detail.get('link'))
    print(id)
 

df['invoice']=invoice
df.to_excel("file_name.xlsx",index=False)
