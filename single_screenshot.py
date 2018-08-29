import pyautogui
import mss.tools
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8
import datetime
import importlib
"""importlib.reload(sc)

window = 0

with mss.mss() as sct:
    filename = sct.shot()
if window != 0:
    sc.get_m6_coordinates()
    x, y = m6['title']
    with mss.mss() as sct:
        # The screen part to capture
        monitor = {'top': y, 'left': x, 'width': 700, 'height': 500}
        now = datetime.datetime.now()
        output = 'monitor-1-crop.png'

        # Grab the data
        sct_img = sct.grab(monitor)

        # Save to the picture file
        mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)"""

sc.get_m6_coordinates()
x, y = m6['title']
with mss.mss() as sct:
    # The screen part to capture
    monitor = {'top': y + 188, 'left': x + 244, 'width': 7, 'height': 10}
    now = datetime.datetime.now()
    output = 'monitor-1-crop.png'

    # Grab the data
    sct_img = sct.grab(monitor)

    # Save to the picture file
    mss.tools.to_png(sct_img.rgb, sct_img.size, output=output)
    print(str(mss.tools.to_png(sct_img.rgb, sct_img.size)))
    #print(sc.m3_tour_status[str(mss.tools.to_png(sct_img.rgb, sct_img.size))])