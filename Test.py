from tkinter import Tk
import clipboard

clipboard.copy("abc")
text = clipboard.paste()
print(text)