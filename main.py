import pyautogui
import os
import subprocess
import time
import cv2
import configparser

from typing import final

CONFIG_PATH:         final = r".\MetasysOfficeConfig.ini"

URL_ERROR:           final = "설치 프로그램 경로가 잘못 되었습니다\n관리자에게 문의하세요"
INIT_MSG:            final = "확인을 누르면 자동으로 설치를 시작 합니다"
COMPLETE_MSG:        final = "설치가 완료 되었습니다\n바탕화면에 office_key.txt 시리얼 코드를 사용하세요"
EXIT_MSG:            final = "이곳을 클릭하고 Ctrl + C를 눌러 강제 종료합니다..."

img_list = list()
steps    = dict()
config   = configparser.ConfigParser()
    
def next(target: pyautogui.locateOnScreen):
    pyautogui.moveTo(target)
    pyautogui.click()
    pyautogui.press('enter')

def step1(target: pyautogui.locateOnScreen):
    pyautogui.moveTo(target)
    pyautogui.click()
    for i in range(5):
        pyautogui.press('tab')
    
    pyautogui.hotkey('alt', 'a')
    
    for i in range(2):
        pyautogui.press('tab')
    
    pyautogui.press('enter')

def step2(target: pyautogui.locateOnScreen):
    installer_key = [ "DV2M3", "N4UTE", "ECWE8", "R3BU7" ]
    pyautogui.moveTo(target)

    for i in installer_key:
        for j in i:
            pyautogui.press(j)

    for i in range(4):
        pyautogui.press('tab')

    pyautogui.press('enter')

    reg = subprocess.Popen(os.path.join(config["Path"]["Office"], "x86_key.reg"), shell=True, stdout = subprocess.PIPE)

    time.sleep(1)

    for i in range(2):
        pyautogui.press('tab')
    pyautogui.press('enter')

    time.sleep(1)

    pyautogui.press('enter')
    reg.kill()
    
def next_step() -> bool:
    os.system("cls")
    print(EXIT_MSG)

    while True:
        for img in img_list:
            try:
                url = os.path.join(config["Path"]["Img"], img)
                target = pyautogui.locateOnScreen(url, confidence=0.95)
                if target is not None:
                    if img == "PreInstalled.png" or img == "PreInstalled_2.png" or img == "Setup_6.png":
                        return True
                    else:
                        steps[img](target)
                        return False
            except:
                continue
            
        time.sleep(1)

def force_quit(subproc: subprocess.Popen):
    serial_key_path = os.path.expanduser(r"~\Desktop\office_key.txt")
    
    if os.path.isfile(serial_key_path) == False:
        serial_key_file = open(serial_key_path, "w")
        serial_key_file.write("DV2M3-N4UTE-ECWE8-R3BU7")
        serial_key_file.close()

    pyautogui.alert(COMPLETE_MSG)
    cv2.destroyAllWindows()
    subproc.terminate()
    exit()

def main():
    if  os.path.isfile(CONFIG_PATH) == False:
        config["Path"]             = {}
        config["Path"]["Office"]   = r"."
        config["Path"]["Img"]   = r".\img"

        with open(CONFIG_PATH, 'w') as outfile:
            config.write(outfile)
    else:
        config.read(CONFIG_PATH)

    for subpath in config["Path"]:
        if os.path.isdir(config["Path"][subpath]) == False:
            pyautogui.alert(URL_ERROR)
            exit()
    
    alert_img = cv2.imread(os.path.join(config["Path"]["Img"], "Alert.png"))
    cv2.imshow("Alert", alert_img)
    cv2.moveWindow("Alert", int(pyautogui.size()[0] / 2) - int(alert_img.shape[1] / 2), 0)

    pyautogui.alert(INIT_MSG)

    installer = subprocess.Popen(
        [ os.path.join(config["Path"]["Office"], "Setup.exe") ], 
        shell=True, 
        stdout = subprocess.PIPE, 
        stderr=subprocess.STDOUT)
    
    global img_list
    img_list = os.listdir(config["Path"]["Img"])
    img_list.remove("Alert.png")
    
    steps["PreInstalled.png"]   = lambda: ()
    steps["PreInstalled_2.png"] = lambda: ()
    steps["Setup_0.png"]        = next
    steps["Setup_1.png"]        = next
    steps["Setup_2.png"]        = step1
    steps["Setup_3.png"]        = next
    steps["Setup_4.png"]        = next
    steps["Setup_5.png"]        = step2
    steps["Setup_6.png"]        = lambda: ()

    while True:
        end_flag = next_step()

        if end_flag == True:
            force_quit(installer)
        else:
            continue

if __name__ == "__main__":
    main()