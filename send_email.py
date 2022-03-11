import pyautogui
import pyperclip
import time
import pandas
import openpyxl
import numpy

pyautogui.PAUSE = 2

# Accessing drive & download file
pyautogui.hotkey("alt", "tab")
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://...")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(2)
pyautogui.click(x=603, y=503, clicks=2)
pyautogui.click(x=655, y=500, button="right")
pyautogui.click(x=944, y=1275)

# Accessing file & handling the data
time.sleep(5)
file1 = pandas.read_excel(r"C:\...")
total_sales = file1["Quantidade"].sum()
total_value = file1["Valor Final"].sum()

# Send email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://...")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(2)
pyautogui.click(x=273, y=340)
pyperclip.copy("email@email.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press('tab')
pyperclip.copy("Titulo do email")
pyautogui.hotkey("ctrl", "v")
pyautogui.press('tab')
message = f"""
Hi,

This is a test automated email.

Total sales = {total_sales}
total_value = R$ {total_value:.2f}

Thank you for your cooperation.

Carry on.
"""
pyperclip.copy(message)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
