import pyautogui
import time

time.sleep(2)

while True:
    time.sleep(1)
    pyautogui.moveTo(970, 461)
    pyautogui.rightClick()
    pyautogui.moveTo(1100, 517)
    pyautogui.moveTo(1332, 538, 0.5)
    pyautogui.click()
    time.sleep(2.5)
    for _ in range(28):
        pyautogui.press('pageup')