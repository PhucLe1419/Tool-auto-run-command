import os
import pysftp
import time
from datetime import datetime
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
from tkinter import *
from tkinter import ttk, filedialog
from tkinter.filedialog import askopenfile
import pandas as pd
from query import data


def select_file():
    win = Tk()
    win.geometry("500x150")

    def open_file():
        file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx')])
        if file:
            global filepath
            filepath = os.path.abspath(file.name)
            win.destroy()

    # Add a Label widget
    label = Label(win, text="Click the Button to browse the Files", font=('Georgia 13'))
    label.pack(pady=10)

    # Create a Button
    ttk.Button(win, text="Browse", command=open_file).pack(pady=20)

    win.mainloop()
    return filepath


def create_command(filepath):
    site_name = "Name"  # Get name of the first column
    df_command = pd.read_excel(filepath, engine="openpyxl")
    df_command.dropna(inplace=True)  # del value Null
    df_command.drop_duplicates(inplace=True)  # del value duplicated

    df_site = df_command[site_name]
    number_command = len(df_site)
    list_teles = list(df_site.drop_duplicates())
    print("Numbers of command: ", number_command)
    print("Numbers of tele_station: ", list_teles)

    file_list_teles = "file_teles.txt"
    with open(file_list_teles, "w") as file:
        for teles in list_teles:
            file.write(teles + '\n')

    df_command['command'] = 'set ' + df_command['MO'] + ' ' + df_command['Parameter'] + ' ' + df_command[
        'New Value'].astype(
        str) + '\n'
    df_command_new = df_command[[site_name, 'command']].copy()
    list_group = [x for _, x in df_command_new.groupby(site_name)]
    for group in list_group:
        file_name = group.iloc[0, 0]  # Get first cell as name of file
        file = open(file_name + '.txt', 'w')
        file.write('lt all\n')
        for row in group['command']:
            file.write(row)
        file.write('\nalt')
    print('Create file succesfully!')
    return list_teles, number_command, file_list_teles


def upload(list_teles, file_list_teles):
    # create file name today
    today = str(datetime.today().strftime('%Y%m%d'))  
    os.makedirs(today, exist_ok=True)

    targetpath = data.path_ftp
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None
    with pysftp.Connection(host=data.host_ftp, username=data.user_name, password=data.pass_word, port=data.port_ftp,
                           cnopts=cnopts) as sftp:
        sftp.cwd(targetpath)
        sftp.put(file_list_teles)  # up file teles to server
        shutil.copy(file_list_teles, today)
        os.remove(file_list_teles)
        for i in list_teles:
            sftp.put(i + '.txt')
            shutil.copy(i + '.txt', today)
            os.remove(i + '.txt')
    print('Upload success!')


def shoot_command(file_list_teles, number_command, list_teles):
    driver = webdriver.Chrome()
    driver.get(data.url_web)
    driver.maximize_window()
    time.sleep(2)

    # login to web
    alarm_status = driver.find_element(By.ID, 'details-button')
    alarm_status.click()
    time.sleep(2)
    next_but = driver.find_element(By.ID, 'proceed')
    next_but.click()
    time.sleep(2)
    ok = driver.find_element(By.ID, 'login')
    ok.click()
    time.sleep(2)
    username = driver.find_element(By.ID, 'loginUsername')
    username.send_keys(data.user_key)
    time.sleep(2)
    password = driver.find_element(By.ID, 'loginPassword')
    password.send_keys(data.pass_key)
    time.sleep(2)
    submit = driver.find_element(By.ID, 'submit')
    submit.click()
    time.sleep(2)
    continue_ = driver.find_element(By.ID, 'continueButton')
    continue_.click()
    time.sleep(5)

    # set file to shell terminal
    command_ = file_list_teles + " 'run name.txt'"
    pyautogui.typewrite(command_)
    pyautogui.press('enter')
    time_run_command = number_command + len(list_teles) * 90
    print('Time is:', time_run_command)
    time.sleep(time_run_command)
    driver.close()


# main
filepath = select_file()
print('Program is running...')
time.sleep(2)
list_teles, number_command, file_list_teles = create_command(filepath)
time.sleep(2)
upload(list_teles, file_list_teles)
time.sleep(2)
shoot_command(file_list_teles, number_command, list_teles)
print('Finished!')
time.sleep(10)
