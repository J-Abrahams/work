import mss
import mss.tools
import datetime
import screenshot_data as sc
from screenshot_data import m1, m2, m3, m4, m5, m6, m7, m8, m9, m10
import pyautogui
# import importlib
# importlib.reload(sc)


def take_screenshot(x, y, width, height, save_file=False):
    with mss.mss() as sct:
        monitor = {'top': y, 'left': x, 'width': width, 'height': height}
        im = sct.grab(monitor)
        screenshot = str(mss.tools.to_png(im.rgb, im.size))
        if save_file:
            now = datetime.datetime.now()
            output = now.strftime("%m-%d-%H-%M-%S.png".format(**monitor))
            mss.tools.to_png(im.rgb, im.size, output=output)
        return screenshot


def check_for_upgrade():
    sc.get_m3_coordinates()
    x, y = m3['title']
    pyautogui.click(m3['events'])
    while not pyautogui.pixelMatchesColor(x + 290, y + 260, (8, 36, 107)):
        pass
    n = 0
    for i in range(11):
        prior_value = take_screenshot(x + 431, y + 256 + 13 * n, 55, 10)
        new_value = take_screenshot(x + 751, y + 256 + 13 * n, 55, 10)
        if prior_value == sc.day_drive_event and new_value == sc.minivac_event:
            return 'upgrade'
        else:
            n += 1


def test_screenshot():
    sc.get_m3_coordinates()
    x, y = m3['title']
    assert take_screenshot(x + 431, y + 256 + 13 * 6, 55, 10) == sc.day_drive_event
    assert take_screenshot(x + 751, y + 256 + 13 * 6, 55, 10) == sc.minivac_event


def test_upgrade():
    assert check_for_upgrade() == 'upgrade'
