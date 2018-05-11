import json 
import requests

URL = "https://graph.microsoft.com/v1.0/me/drive/items/01XIYPASAOAUC35YQNN5DISDSY3ILCP6OC/workbook/worksheets('Sheet1')/range(address='A1:D2')"

BASE = "https://graph.microsoft.com"
VERSION = "v1.0"

headers  = {"Content-Type": "application/json"}


#Start Session
itemID = "01XIYPASAOAUC35YQNN5DISDSY3ILCP6OC"
postURL = BASE+VERSION+"/me/drive/items"+itemID+"/workbook/createSession"
payload = {'persistChanges' : 'true'}

resp = request.post(postURL,headers = headers, json=payload)
if resp.status_code != 200:
    print('error: ' + str(resp.status_code))
else:
    print('Success')

#Patch data in
SheetName = "'Sheet1'"
patchEdit = "range(address='A1:C2')"
patchURL = "{base}/{version}/me/drive/items/{item_id}/workbook/worksheets({choice})/{edit}".format(base=BASE,
                                                                                        version=VERSION,
                                                                                        item_id=itemID,
                                                                                        choise=SheetName,
                                                                                        edit=patchEdit)
patchPayload = {
  "values": [
    [0,0,0],
    ["hot","cross","buns"]
    ]
  }

resp = request.patch(patchURL,headers=headers,json=patchPayload)

if resp.status_code != 200:
    print('error: ' + str(resp.status_code))
else:
    print('Success')