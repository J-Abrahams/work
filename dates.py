import keyboard
import pyautogui
import time
import pyperclip

image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                          region=(514, 245, 889, 566))
while image == None:
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_tour_menu.png',
                                               region=(514, 245, 889, 566))
x, y = image

notes = []
for i in range(5):
    pyautogui.doubleClick(x, y + 58)
    y += 14
    image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_changing_note_menu.png',
                                          region=(514, 245, 889, 566))
    while image == None:
        image = pyautogui.locateCenterOnScreen('C:\\Users\\Jared.Abrahams\\Screenshots\\sc_changing_note_menu.png',
                                               region=(514, 245, 889, 566))
    x_1, y_1 = image
    pyautogui.rightClick(x_1 + 50, y_1 + 75)
    pyautogui.click(x_1 + 60, y_1 + 190)
    pyautogui.rightClick(x_1 + 50, y_1 + 75)
    pyautogui.click(x_1 + 60, y_1 + 130)
    pyautogui.click(x_1 + 60, y_1 + 250)
    notes.append(pyperclip.paste())
print(notes)
