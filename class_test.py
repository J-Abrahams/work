class Deposit:

    def __init__(self, deposit_type, amount):
        self.deposit_type = deposit_type
        self.amount = amount

    def change_description(self):
        sc.get_m3_coordinates()
        amount = 0
        old_title = 'old'
        x_2, y_2 = m3['deposit_1']
        while amount != price:
            pyautogui.click(m3['tour_packages'])
            pyautogui.click(x_2, y_2)
            pyautogui.click(m3['change_deposit'])
            sc.get_m6_coordinates()
            with mss.mss() as sct:
                x, y = m6['title']
                monitor = {'top': y + 187, 'left': x + 168, 'width': 51, 'height': 11}
                im = sct.grab(monitor)
                try:
                    amount = sc.screenshot_dict[str(mss.tools.to_png(im.rgb, im.size))]
                except KeyError:
                    amount = 0
            if amount != price:
                pyautogui.click(m6['ok'])
                y_2 += 13
        while 'ref' not in old_title.lower():
            pyautogui.click(m6['description'])
            keyboard.send('ctrl + z')
            keyboard.send('ctrl + c')
            r = Tk()
            old_title = r.selection_get(selection="CLIPBOARD")
        new_title = old_title.replace("able", "ed")
        new_title = new_title.replace("ABLE", "ED")
        new_title = new_title.replace(" /", "/")
        new_title = new_title.replace("/ ", "/")
        keyboard.write(new_title)

