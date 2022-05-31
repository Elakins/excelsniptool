import os
from pathlib import Path
import time
import pyautogui
from PIL import Image


absolutePath = Path('G:\Elakins\Python Files\Snip Tool Trial\Sample.xlsx').resolve()
os.system(f'start excel.exe "{absolutePath}"')

time.sleep(2)

pyautogui.leftClick(x=204, y=1002)
time.sleep(3)
pyautogui.leftClick(x=204, y=1002)

myScreenshot = pyautogui.screenshot()   
myScreenshot.save(r'G:\Elakins\Python Files\Snip Tool Trial\Sample picture.png')

im = Image.open('G:\Elakins\Python Files\Snip Tool Trial\Sample picture.png')
im_crop = im.crop((25,226,377,974))
im_crop.save('G:\Elakins\Python Files\Snip Tool Trial\Sample picture.png', quality=95)

pyautogui.leftClick(x=31, y=15)
pyautogui.leftClick(x=1886, y=11)
