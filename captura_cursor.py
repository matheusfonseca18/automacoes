import pyautogui
import time

print("Capturando a posição do cursor. Pressione Ctrl+C para parar.")
time.sleep(2)

while True:
    pos = pyautogui.position()
    print(f"Posição do cursor: {pos}")
    time.sleep(3)