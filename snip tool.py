import os
from pathlib import Path
import time
import pyautogui
from PIL import Image

absolutePath1 = Path('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\Electric New hoshin.xlsm').resolve()

os.system(f'start excel.exe "{absolutePath1}"')

time.sleep(5)

pyautogui.leftClick(x=528, y=1003)
time.sleep(3)
pyautogui.leftClick(x=528, y=1003)
time.sleep(3)

myScreenshot = pyautogui.screenshot()   
myScreenshot.save(r'G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\E4.PNG')

im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\E4.PNG')
im_crop = im.crop((57,126,626,897))
im_crop.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\E4.PNG', quality=110)

im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\E4.PNG')
im_resize = im.resize((1974,2889))
im_resize.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\E4.PNG', quality=110)


pyautogui.leftClick(x=601, y=1001)
time.sleep(3)
pyautogui.leftClick(x=601, y=1001)

myScreenshot = pyautogui.screenshot()   
myScreenshot.save(r'G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\EWP.PNG')

im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\EWP.PNG')
im_crop = im.crop((57,126,626,897))
im_crop.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\EWP.PNG', quality=110)

im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\EWP.PNG')
im_resize = im.resize((1974,2889))
im_resize.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\EWP.PNG', quality=110)

pyautogui.leftClick(x=31, y=15)
pyautogui.leftClick(x=960, y=595)
pyautogui.leftClick(x=1886, y=11)

time.sleep(5)


def second():
    time.sleep(5)
    
    absolutePath2 = Path('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\ATM-IND-CHASS-TC New hoshin.xlsm').resolve()
    
    os.system(f'start excel.exe "{absolutePath2}"')

    time.sleep(10)

    pyautogui.leftClick(x=864, y=1000)
    time.sleep(3)
    pyautogui.leftClick(x=864, y=1000)
    time.sleep(3)

    myScreenshot = pyautogui.screenshot()   
    myScreenshot.save(r'G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Industrial.PNG')

    im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Industrial.PNG')
    im_crop = im.crop((57,126,626,897))
    im_crop.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Industrial.PNG', quality=110)

    im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Industrial.PNG')
    im_resize = im.resize((1974,2889))
    im_resize.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Industrial.PNG', quality=110)

    pyautogui.leftClick(x=975, y=1003)
    time.sleep(3)
    pyautogui.leftClick(x=975, y=1003)

    myScreenshot = pyautogui.screenshot()   
    myScreenshot.save(r'G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Chassis.PNG')

    im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Chassis.PNG')
    im_crop = im.crop((57,126,626,897))
    im_crop.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Chassis.PNG', quality=110)

    im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Chassis.PNG')
    im_resize = im.resize((1974,2889))
    im_resize.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\Chassis.PNG', quality=110)

    pyautogui.leftClick(x=1104, y=1001)
    time.sleep(3)
    pyautogui.leftClick(x=1104, y=1001)

    myScreenshot = pyautogui.screenshot()   
    myScreenshot.save(r'G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\ATM.PNG')

    im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\ATM.PNG')
    im_crop = im.crop((57,126,626,897))
    im_crop.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\ATM.PNG', quality=110)

    im = Image.open('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\ATM.PNG')
    im_resize = im.resize((1974,2889))
    im_resize.save('G:\Virtual Blue Board\Daily Hoshins\Daily Hoshins\TV Display Hoshin\ATM.PNG', quality=110)

    pyautogui.leftClick(x=31, y=15)
    pyautogui.leftClick(x=960, y=595)
    pyautogui.leftClick(x=1886, y=11)

    


second()

print("Process Completed!!!!!!!!!")
