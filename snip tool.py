import os
from pathlib import Path
import time
import pyautogui
from PIL import Image


absolutePath = Path('G:\Example folder for BC\Snip Program\Electric New hoshin.xlsm').resolve()
os.system(f'start excel.exe "{absolutePath}"')

time.sleep(5)

pyautogui.leftClick(x=528, y=1003)
time.sleep(3)
pyautogui.leftClick(x=528, y=1003)
time.sleep(3)

myScreenshot = pyautogui.screenshot()   
myScreenshot.save(r'G:\Example folder for BC\Snip Program\E4.PNG')

im = Image.open('G:\Example folder for BC\Snip Program\E4.PNG')
im_crop = im.crop((40,125,637,903))
im_crop.save('G:\Example folder for BC\Snip Program\E4.PNG', quality=110)


pyautogui.leftClick(x=601, y=1001)
time.sleep(3)
pyautogui.leftClick(x=601, y=1001)

myScreenshot = pyautogui.screenshot()   
myScreenshot.save(r'G:\Example folder for BC\Snip Program\EWP.PNG')

im = Image.open('G:\Example folder for BC\Snip Program\EWP.PNG')
im_crop = im.crop((40,125,637,893))
im_crop.save('G:\Example folder for BC\Snip Program\EWP.PNG', quality=110)

pyautogui.leftClick(x=31, y=15)
pyautogui.leftClick(x=960, y=595)
pyautogui.leftClick(x=1886, y=11)
