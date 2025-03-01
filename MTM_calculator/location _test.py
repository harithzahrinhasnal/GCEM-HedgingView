import time
import pyautogui

time.sleep(3)
pyautogui.moveTo([390,580], duration = 0.5)
pyautogui.press('enter')
pyautogui.write('TEST')