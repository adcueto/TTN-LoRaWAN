import time
import ttn
import json
import openpyxl
import base64
from datetime import datetime
global iRow

app_id = "beacon-encontrack"
access_key = "ttn-account-v2.czFauryQ1RGYZTkuneK-GtFj_C3o38EsHvbW062SC1o"
dest_filename = "/Users/turulo/Developer/Python/Hello/LoRa.xlsx"

doc = openpyxl.load_workbook(filename=dest_filename)
general = doc.worksheets[0]  #Sheet General

iRow =2

def uplink_callback(msg, client):
    global iRow;

    print("Received uplink from ", msg.dev_id)
    general.cell(row=iRow, column=1).value = msg.dev_id
    general.cell(row=iRow, column=2).value = msg.metadata.gateways[0].gtw_id
    general.cell(row=iRow, column=3).value = msg.payload_raw
    doc.save(filename = dest_filename)
    iRow+=1
    print(iRow)

handler = ttn.HandlerClient(app_id, access_key)
# using mqtt client
mqtt_client = handler.data()
mqtt_client.set_uplink_callback(uplink_callback)
mqtt_client.connect()
time.sleep(50000)
mqtt_client.close() #close mqtt client
