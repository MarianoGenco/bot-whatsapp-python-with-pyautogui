import webbrowser
import pyautogui
from time import sleep
import pathlib
import pandas as pd
import os

ruta = str(pathlib.Path(__file__).parent.absolute()) 

#SELECT SPREADSHEET
spreadsheet = pd.read_excel(ruta + "/contacts.xlsx")
contacts = pd.ExcelFile(ruta + "/contacts.xlsx").parse(0) 
spreadsheet['name_user'] = spreadsheet['name_user'].astype(str)
spreadsheet['cell_phone'] = spreadsheet['cell_phone'].astype(str).str.replace(" ", "").str.replace("-", "")

#DELETE THE ROW IF THE PHONE IS REPEATED. 
#IF YOU NEED TO PLACE MORE THAN ONE PHONE, ENTER A ',' IN THE SAME CELL. 
#IN THE SPREADSHEET YOU CAN SEE EXAMPLES.
eliminar_duplicados = contacts.drop_duplicates() 
eliminar_duplicados.to_excel(ruta + "/contacts.xlsx", index=False)
spreadsheet['name_user'] = spreadsheet['name_user'].astype(str)
spreadsheet['cell_phone'] = spreadsheet['cell_phone'].astype(str)
df = pd.DataFrame(columns=['cell_phone','name_user'])

#COUNT NUMBERS REMOVED
x=contacts.shape[0]
contacts = contacts.drop_duplicates()
y=contacts.shape[0]
z=x-y
print(z,'duplicate contacts found and removed.')
print(y, 'contacts left.')

#FOR EACH PHONE NUMBER IN THE SPREADSHEET, 
#IF THERE IS A REPEATED NUMBER IN THE SAME CELL, 
#ONLY ONE IS SENT. (NOT REMOVED)
for each_cell in range(0, y):
    cell_phones = (spreadsheet['cell_phone'][each_cell].split(','))
    clean_list = set(cell_phones)
    list = list(clean_list)  

    #PLEASE CHOOSE YOUR NUMBER CODE COUNTRY
    code_country = str(54)   
    for cell_x_cell in list: 
        name = spreadsheet['name_user'][each_cell]
        print(f'Sending to {name} with the number {cell_x_cell}') 
        sleep(5) 
        #WRITE YOUR MESSAGGE
        message = f'Hi *{name}*, how are you?'
        webbrowser.open("https://api.whatsapp.com/send?phone=" + code_country + cell_x_cell)
        sleep(2)
        print(f'Write message... to {name}\n')    
        pyautogui.typewrite(message)
        sleep(2) 
        pyautogui.press('enter')
        print('Message sent.')
        sleep(2) 
    sleep(2)
    #CLOSE CHROME.EXE, YOU CAN DELETE THIS
    os.system('taskkill /F /IM chrome.exe')
    sleep(2)


print('The execution has ended. You can close the program.')