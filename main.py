import pywhatkit
import time
import pyautogui
import keyboard as kp
import argparse
import datetime
import pandas as pd


def read_excel():
    devs_df2 = pd.read_excel(r'devs.xlsx', converters={'Name': str, 'Phone': str, 'Email': str})
    return devs_df2


def read_message():
    f = open("message.txt", "r")
    msg = f.read()
    f.close()
    return msg


def format_message(msg, name, email):
    msg = msg.replace('{{name}}', name)
    msg = msg.replace('{{email}}', email)
    return msg


def get_arguments():
    parser = argparse.ArgumentParser(
        description="This a program that helps Lazymerlada to send a certain message to people saved on an excel sheet."
        , exit_on_error=True)
    # parser.add_argument(
    #     '-m',
    #     '--message',
    #     help='Enter the message to be sent',
    #     required=True
    # )
    # parser.add_argument(
    #     '-n',
    #     '--number',
    #     help='Phone number to send a message to (international format with)',
    #     required=True
    # )
    # parser.add_argument(
    #     '-t',
    #     '--time',
    #     help='When to send message in (hh:mm) in 24 hours format',
    #     required=True
    # )
    # parser.add_argument(
    #     '-v',
    #     '--verbose',
    #     help='it will print descriptive messages of the script status, if it was provided.',
    #     nargs='?',
    #     const=True,
    #     type=bool
    # )
    return parser.parse_args()


def greet(name):
    print(f'Hi, {name} the script is running now, sending messages...')


def format_name(name):
    name = name.capitalize().split(' ')
    return name[0] if len(name) > 1 else name


def format_email(email):
    return email.lower()


def format_phone(phone_no):
    return phone_no if phone_no[0] == '+' else '+' + phone_no


def send_messages(devs, msg):
    for index, dev in devs.iterrows():
        name = format_name(dev['Name'])
        phone_no = format_phone(dev['Phone'])
        email = format_email(dev['Email'])
        temp_message = format_message(msg, name, email)
        send_message(temp_message, phone_no, name, email)


def send_message(msg, phone_no, name, email):
    now = datetime.datetime.now()

    pywhatkit.sendwhatmsg(
        phone_no=phone_no
        , message=msg
        , time_hour=now.hour
        , time_min=now.minute + 1
        , wait_time=7
        , tab_close=True
        , close_time=2
    )
    # pyautogui.click(155, 1000)  # this might need be adjust depending on the screen size
    # time.sleep(2)
    # kp.press_and_release('enter')
    press_enter()

    print(f'Message has been sent to {name} with email: {email} and {phone_no}')


def press_enter():
    chrome_window = pyautogui.getWindowsWithTitle('Chrome')[0]
    chrome_window.activate()
    dims = pyautogui.size()
    pyautogui.click(dims.width/2, dims.height/2)
    time.sleep(1)
    pyautogui.press('enter')


if __name__ == '__main__':
    greet('Lazymeralda')
    # args = get_arguments()

    message = read_message()
    devs_df = read_excel()
    send_messages(devs_df, message)


