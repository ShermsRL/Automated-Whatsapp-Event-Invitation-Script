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

SLEEP_TIME_LONG = 2
SLEEP_TIME_SHORT = 1

def get_selected_path(main_path, selected_index):
    return f"{main_path}/{os.listdir(main_path)[int(selected_index)-1]}"

def get_option(path, prompt):
    for index, file in enumerate(os.listdir(path)):
        print(index + 1, file)
    option = input(prompt)
    return get_selected_path(path, option)

def paste():
    pywinauto.keyboard.send_keys('^v')

def send(clip_type, data): 
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(clip_type, data)
    win32clipboard.CloseClipboard()

def send_to_clipboard(clip_type, filepath):
    image = Image.open(filepath)

    output = BytesIO()
    image.convert("RGB").save(output, "BMP")
    data = output.getvalue()[14:]
    output.close()

    send(clip_type, data)

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
            message_path = get_option(MESSAGE_PATH, "Please choose the message you would like to send \n")
            
            # Reading message text
            with open(message_path) as file:
                selected_text = file.read()

            # Attaching image option
            image_option = input("Would you like to attach a image to your message? Y or N \n")
            if image_option == 'Y':
                image_path = get_option(IMAGE_PATH, "Please select the image to be attached. \n")

            # Excel interaction
            excel_path = get_option(EXCEL_SHEET_PATH, "Please choose the excel that you are using \n")

            data_frame = pd.read_csv(excel_path)
            print(data_frame.head())

            for index, columnName in enumerate(list(data_frame.columns.values)):
                print(index, columnName)

            num_col_select = int(input("Please choose the column with the numbers \n"))
            cc_col_select = int(input("Please choose the column with the country code \n"))

            # Reading the numbers
            with open(excel_path) as excel_file:
                name_reader = csv.reader(excel_file, delimiter=",")
                next(name_reader)

                os.system("start chrome")
                time.sleep(2) #wait for chrome to launch
                chromeApp = Application(backend='uia').connect(title_re='.*Chrome.*')
                element_name = "Address and search bar"
                dlg = chromeApp.top_window()
                url = dlg.child_window(title=element_name, control_type="Edit")

                for row in name_reader:
                    url.set_edit_text(f"https://wa.me/{row[cc_col_select]}{row[num_col_select]}")
                    send_keys("{ENTER}")
                    time.sleep(SLEEP_TIME_LONG) #wait for whatsapp to launch

                    whatsapp = Application(backend='uia').connect(title='WhatsApp')
                    dlg_wa = whatsapp.top_window()
                    message_textbox = dlg_wa.child_window(auto_id="InputBarTextBox", control_type="Edit").click_input()

                    send_to_clipboard(win32clipboard.CF_DIB, image_path)
                    paste()

                    time.sleep(SLEEP_TIME_LONG) #time between pasting image and copying text
                    pyperclip.copy(selected_text)
                    paste()

                    time.sleep(SLEEP_TIME_SHORT) #time between pasting text sending the message
                    pyautogui.press('enter')

                    time.sleep(SLEEP_TIME_SHORT) #time after sending the message and getting next number

        case "4":
                os.startfile(IMAGE_PATH)

        case "5":
            quit()

def main():
    while menu != "5":
        app_functions(menu())
        os.system('cls') # Clear the command prompt to keep it neat

main()