import pyautogui
import mss
import mss.tools
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import datetime

sc.get_m8_coordinates()
x, y = m8['title']
with mss.mss() as sct:
    # The screen part to capture
    monitor = {'top': y, 'left': x, 'width': 750, 'height': 500}
    now = datetime.datetime.now()
    output = 'screenshotf9.png'

    # Grab the data
    sct_img = sct.grab(monitor)

    # Save to the picture file
    mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)