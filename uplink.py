import time
import ttn
import json
import openpyxl
import base64
from base64 import b64decode
global iRow, iColDev, iColMeta, iColGw, nGw;

app_id = "beacon-encontrack"
access_key = "ttn-account-v2.czFauryQ1RGYZTkuneK-GtFj_C3o38EsHvbW062SC1o"
dest_filename = "LoRa.xlsx"  #The name of File Excel

doc = openpyxl.load_workbook(filename=dest_filename) #Open File Excel
gSheet = doc.worksheets[0]  #"general" Sheet
iRow=3 #begin of row third

#================= Function Callback ===========================================
def uplink_callback(msg, client):
    global iRow, iColMeta, iColGw, nGw; #variable definition

    print("Received uplink from ", msg.dev_id)

    nGw = len(msg[6][6]) # gateways number
    iColDev=1; iColMet=7; iColGw=14
    #===== General processing =================================================
    gSheet.cell(row=iRow, column=1).value = msg.dev_id
    gSheet.cell(row=iRow, column=2).value = msg.hardware_serial
    gSheet.cell(row=iRow, column=3).value = msg.port
    gSheet.cell(row=iRow, column=4).value = msg.counter
    gSheet.cell(row=iRow, column=5).value = b64decode(msg.payload_raw).hex() #convert payload from base64 to hexadecimal
    #====== Metadata processing ===============================================
    for iVal in range(6):
        gSheet.cell(row=iRow, column=iColMet).value = msg.metadata[iVal]
       iColMet+=1 #column increment
    #====== Gateways processing ================================================
    for iGw in range(nGw):
        for gVal in range(6):
            gSheet.cell(row=iRow, column=iColGw).value = msg.metadata.gateways[iGw][gVal]
            iColGw+=1;#column increment

    gSheet.cell(row=iRow, column=13).value = nGw # print gateway number
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
