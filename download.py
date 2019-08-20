
import time
import ttn

#
app_id = "beacon-encontrack"
access_key = "ttn-account-v2.czFauryQ1RGYZTkuneK-GtFj_C3o38EsHvbW062SC1o"

handler = ttn.HandlerClient(app_id, access_key)
mqtt_client = handler.data()
mqtt_client.connect()
mqtt_client.send(dev_id="fipy",  pay="AQID", port=1, conf=True, sched="replace")
mqtt_client.close()
