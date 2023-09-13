import os
import time
import csv
import pandas as pd
from pywinauto.application import Application
import pyperclip
from pywinauto.keyboard import send_keys
import pywinauto
import pyautogui
from PIL import Image
from io import BytesIO
import win32clipboard

# Program Initialisation
if not os.path.exists("./excels"):
    os.mkdir("./excels")
if not os.path.exists("./messages"):
    os.mkdir("./messages")
if not os.path.exists("./images"):
    os.mkdir("./images")

global selected_text
global image_select

EXCEL_SHEET_PATH = f"{os.getcwd()}/excels"
MESSAGE_PATH = f"{os.getcwd()}/messages"
IMAGE_PATH = f"{os.getcwd()}/images"
PROGRAM_FUNCTIONS = ["Upload Excel", "Draft a Message", "Send a Message", "Upload Image", "Exit"]


def send_to_clipboard(clip_type, filepath):
    image = Image.open(filepath)

    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def menu():
    print("Welcome to Automated Whatsapp Message Sender")
    for index, function in enumerate(PROGRAM_FUNCTIONS):
        print(index + 1, function)
    function_select = input("Please select what you would like to do next \n")
    return function_select

def app_functions(function_select):
    match function_select:
        case "1":
            os.startfile(EXCEL_SHEET_PATH)

        case "2":
            messageName = input("Please give your draft message a name. \n")
            os.system(f"start notepad.exe ./messages/{messageName}")

        case "3":
            # Choosing message to send
            for index, message in enumerate(os.listdir(MESSAGE_PATH)):
                print(index+1, message)
            message_select = input("Please choose the message you would like to send \n")

            # Reading message text
            with open(f"{MESSAGE_PATH}/{os.listdir(MESSAGE_PATH)[int(message_select)-1]}") as file:
                selected_text = file.read()

            # Attaching image option
            image_option = input("Would you like to attach a image to your message? Y or N \n")
            if image_option == 'Y':
                for index, image in enumerate(os.listdir(IMAGE_PATH)):
                    print(index+1, image)
                image_select = input("Please select the image to be attached. \n")

            # Excel interaction
            for index, excel in enumerate(os.listdir(EXCEL_SHEET_PATH)):
                print(index+1, excel)
            excel_select = input("Please choose the excel that you are using \n")
            df = pd.read_csv(f"{EXCEL_SHEET_PATH}/{os.listdir(EXCEL_SHEET_PATH)[int(excel_select)-1]}")
            print(df.head())
            for index, columnName in enumerate(list(df.columns.values)):
                print(index, columnName)
            num_col_select = int(input("Please choose the column with the numbers \n"))
            cc_col_select = int(input("Please choose the column with the country code \n"))

            # Reading the numbers
            with open(f"{EXCEL_SHEET_PATH}/{os.listdir(EXCEL_SHEET_PATH)[int(excel_select)-1]}") as excel_file:
                name_reader = csv.reader(excel_file, delimiter=",")
                next(name_reader)

                os.system("start chrome")
                time.sleep(2)
                chromeApp = Application(backend='uia').connect(title_re='.*Chrome.*')
                element_name = "Address and search bar"
                dlg = chromeApp.top_window()
                url = dlg.child_window(title=element_name, control_type="Edit")

                for row in name_reader:
                    url.set_edit_text(f"https://wa.me/{row[cc_col_select]}{row[num_col_select]}")
                    send_keys("{ENTER}")
                    time.sleep(2)

                    whatsapp = Application(backend='uia').connect(title='WhatsApp')
                    dlg_wa = whatsapp.top_window()
                    message_textbox = dlg_wa.child_window(auto_id="InputBarTextBox", control_type="Edit").click_input()
                    time.sleep(2)

                    send_to_clipboard(win32clipboard.CF_DIB, f"{IMAGE_PATH}/{os.listdir(IMAGE_PATH)[int(image_select) - 1]}")
                    pywinauto.keyboard.send_keys('^v')
                    time.sleep(2)
                    pyperclip.copy(selected_text)
                    pywinauto.keyboard.send_keys('^v')
                    time.sleep(1)
                    pyautogui.press('enter')
                    time.sleep(1)

        case "4":
                os.startfile(IMAGE_PATH)

        case "5":
            quit()

def main():
    while menu != "5":
        app_functions(menu())
        os.system('cls') # Clear the command prompt to keep it neat

main()


