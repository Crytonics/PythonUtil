import pyautogui
import sys

PeaZip_1 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-1.png', confidence=0.9)

while PeaZip_1 is None:
    PeaZip_1 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-1.png', confidence=0.9)
    pyautogui.sleep(0.5)  # Add a delay to avoid excessive CPU usage

print(pyautogui.center(PeaZip_1))
click_1 = pyautogui.center(PeaZip_1)

pyautogui.moveTo(click_1)
pyautogui.sleep(0.2)  # Add a delay before clicking
pyautogui.click()
pyautogui.sleep(0.2)

PeaZip_2 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-2.png', confidence=0.9)

while PeaZip_2 is None:
    PeaZip_2 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-2.png', confidence=0.9)
    pyautogui.sleep(0.5)  # Add a delay to avoid excessive CPU usage

print(pyautogui.center(PeaZip_2))
click_2 = pyautogui.center(PeaZip_2)

pyautogui.moveTo(click_2)
pyautogui.sleep(0.2)  # Add a delay before clicking
pyautogui.click()
pyautogui.sleep(0.2)

pyautogui.click()
pyautogui.sleep(0.2)

pyautogui.click()
pyautogui.sleep(0.2)

PeaZip_3 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-3.png', confidence=0.9)

while PeaZip_3 is None:
    PeaZip_3 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-3.png', confidence=0.9)
    pyautogui.sleep(0.5)  # Add a delay to avoid excessive CPU usage

print(pyautogui.center(PeaZip_3))
click_3 = pyautogui.center(PeaZip_3)

pyautogui.moveTo(click_3)
pyautogui.sleep(0.2)  # Add a delay before clicking
pyautogui.click()
pyautogui.sleep(0.2)

PeaZip_4 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-4.png', confidence=0.9)

while PeaZip_4 is None:
    PeaZip_4 = pyautogui.locateOnScreen('functions/automate/PeaZip/PeaZip-4.png', confidence=0.9)
    pyautogui.sleep(0.5)  # Add a delay to avoid excessive CPU usage

print(pyautogui.center(PeaZip_4))
click_4 = pyautogui.center(PeaZip_4)

pyautogui.moveTo(click_4)
pyautogui.sleep(0.2)  # Add a delay before clicking
pyautogui.click()
pyautogui.sleep(0.2)

sys.exit()