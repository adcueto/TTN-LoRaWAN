import time
import ttn
import json
import openpyxl
import base64
from datetime import datetime
global iRow, iColDev, iColMeta, iColGw, nGw;

app_id = "beacon-encontrack"
access_key = "ttn-account-v2.czFauryQ1RGYZTkuneK-GtFj_C3o38EsHvbW062SC1o"
dest_filename = "LoRa.xlsx"  #The name of File Excel

doc = openpyxl.load_workbook(filename=dest_filename) #Open File Excel
gSheet = doc.worksheets[0]  #"general" Sheet
iRow=2 #begin of row second

#================= Function Callback ===========================================
def uplink_callback(msg, client):
    global iRow, iColDev, iColMeta, iColGw, nGw; #variable definition

    print("Received uplink from ", msg.dev_id)
    nGw = len(msg[6][6]) # gateways number
    iColDev=1; iColMet=7; iColGw=14

    for iVal in range(6):
        gSheet.cell(row=iRow, column=iColDev).value = msg[iVal]
        gSheet.cell(row=iRow, column=iColMet).value = msg.metadata[iVal]
        iColDev+=1; iColMet+=1 #column increment
    for iGw in range(nGw):
        for gVal in range(6):
            gSheet.cell(row=iRow, column=iColGw).value = msg.metadata.gateways[iGw][gVal]
            iColGw+=1;

    gSheet.cell(row=iRow, column=13).value = nGw
    doc.save(filename = dest_filename) #save file
    print("File updated")
    iRow+=1 #row increment

#================= end Function=================================================

handler = ttn.HandlerClient(app_id, access_key)
# using mqtt client
mqtt_client = handler.data()
mqtt_client.set_uplink_callback(uplink_callback)
mqtt_client.connect()
time.sleep(50000) # duration time
mqtt_client.close() #close mqtt client
