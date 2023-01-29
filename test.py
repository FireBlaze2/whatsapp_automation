import pyperclip
import webbrowser

messagefile = webbrowser.open("C:/Users/Raghav/Desktop/message.txt")
pyperclip.copy(messagefile)
print(pyperclip.paste)
