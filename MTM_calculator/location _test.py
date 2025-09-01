import time
import pyautogui

time.sleep(3)
Monitor = 'BBG Laptop' #options: 'BBG Laptop', 'External Monitor'
if Monitor == 'BBG Laptop': #make sure resolution is (1280,720)
        search_bar = [44, 112]
        CO_1 = [81, 300]
        date_input = [220, 320]
        period = [330 ,330]
        delivery = [280, 340]
        barrel_1 = [280, 480]
        premium_1 = [780, 480]
        calc_time = [361 ,270]
        Upper_strike = [570, 480]
        calc = [82, 181] 
        swap = [399,580]

pyautogui.moveTo(premium_1, duration = 0.5)
pyautogui.click()
pyautogui.click()
pyautogui.write('1')