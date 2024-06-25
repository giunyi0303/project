from multiprocessing.connection import Client
import time
import os
from ppadb.client import Client as AdbClient
from com.dtmilano.android.viewclient import ViewClient

adb_conf = dict(host='127.0.0.1',port=5037)
cur_client = AdbClient(**adb_conf)
cur_devices = cur_client.devices()
len(cur_devices)


# 여기에 쉘 명령어를 추가합니다.
os.system('adb shell am start -a android.net.wifi.PICK_WIFI_NETWORK')
vc = ViewClient(*ViewClient.connectToDeviceOrExit())
wifi_switch = vc.findViewWithText('사용 안함')
if wifi_switch:
    wifi_switch.touch()  # 스위치를 토글합니다.
    vc = ViewClient(*ViewClient.connectToDeviceOrExit())
# 특정 텍스트를 가진 뷰를 찾습니다.
wework_device_view = vc.findViewWithText('')

# 뷰가 존재하면, 그 뷰를 선택합니다.
if wework_device_view:
    wework_device_view.touch()
    
os.system(f'adb shell input text {pw}')
os.system('adb shell input keyevent 66')
os.system('adb shell input keyevent 3')



# https://dtmilano.github.io/AndroidViewClient/
