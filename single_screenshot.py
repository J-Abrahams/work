import pyautogui
import mss.tools
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import datetime

window = 0

with mss.mss() as sct:
    filename = sct.shot()
if window != 0:
    sc.get_m2_coordinates()
    x, y = m2['title']
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {'top': y + 197, 'left': x + 38, 'width': 12, 'height': 9}
        now = datetime.datetime.now()
        output = 'monitor-1-crop.png'

        # Grab the data
        sct_img = sct.grab(monitor)

        # Save to the picture file
        mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)