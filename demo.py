# Version 0.0.1
# April 20, 2023
import re # Python regular expression support
from datetime import datetime # Python datetime conversion support
from concurrent.futures import process # Asynchronous callable execution
from tkinter import * # For user interface
from tkinter.ttk import * # UI widgets
from tkinter.filedialog import askopenfile, askopenfilename, askdirectory # UI-file system interaction
import pandas as pd # For importing, manipulating, and exporting data

ws = Tk()
ws.title('Nova Support Home Data Processor (Demo)')
ws.geometry('500x250')

def process_file(fpath):
    global df

    # Import data
    df=pd.read_excel(fpath)

    # Keep useful columns
    df = df[['Service 1 Description (Code)', 
            'Service Provider','Check-In Date',
            'Check-In Time',
            'Updated Check-In Date',
            'Updated Check-In Time',
            'Check-Out Date',
            'Check-Out Time',
            'Updated Check-Out Date',
            'Updated Check-Out Time',
            'Staff Worked Duration',
            'Staff Worked Duration (Minutes)']]
    
    # Remove parentheses and everything within them
    df['Service 1 Description (Code)'] = df['Service 1 Description (Code)'].str.replace(r'\(.*\)', '')
    # Remove prefix if it exists
    prefix = 'RC-SDP-CLS-320 '
    df['Service 1 Description (Code)'] = df['Service 1 Description (Code)'].apply(lambda x: x[len(prefix):] if x.startswith(prefix) else x)
    # Remove everything after " /"
    df['Service Provider'] = df['Service Provider'].str.split(' /', n=1).str[0]
    # Replace Date/Time with Updated Date/Time if the latter is not NaN
    df['Check-In Date'] = df['Updated Check-In Date'].fillna(df['Check-In Date'])
    df['Check-In Time'] = df['Updated Check-In Time'].fillna(df['Check-In Time'])
    df['Check-Out Date'] = df['Updated Check-Out Date'].fillna(df['Check-Out Date'])
    df['Check-Out Time'] = df['Updated Check-Out Time'].fillna(df['Check-Out Time'])
    df.drop(['Updated Check-In Date', 'Updated Check-In Time','Updated Check-Out Date', 'Updated Check-Out Time'], axis=1, inplace=True)
    # Create datetime series in Python format
    CIDT = df['Check-In Date'].str.cat(df['Check-In Time'], sep=' ')
    CODT = df['Check-Out Date'].str.cat(df['Check-Out Time'], sep=' ')
    CIDT = CIDT.apply(lambda x: datetime.strptime(x, r'%m/%d/%Y %I:%M %p'))
    CODT = CODT.apply(lambda x: datetime.strptime(x, r'%m/%d/%Y %I:%M %p'))
    # Calculate time difference from check-in and check-out datetimes
    CTD = (CODT - CIDT).dt.total_seconds() / 60
    # Convert Staff Work Duration from Hour:Minutes to Minutes
    SWD_min = df['Staff Worked Duration'].apply(lambda x: (int(x.split(':')[0]) * 60) + int(x.split(':')[1]))

    # Sanity check:
    # 1. Check if Staff Work Duration == Staff Work Duration (Minutes)
    sanity1 = (SWD_min == df['Staff Worked Duration (Minutes)'])
    # 2. Check if |Staff Work Duration (Minutes) - Calculated Time Difference| <= 1
    sanity2 = ((df['Staff Worked Duration (Minutes)'] - CTD).abs() <= 1.1) # 1.1 to avoid float precision issues
    df["Sanity Check"] = (sanity1 & sanity2) # The data is "sane" only when both checks are passed
    processsed_lab.configure(text='File Processed Sucessfully!')
    save_btn["state"] = "enabled" # Enable saving botton
    if df['Sanity Check'].all():
        problem_lab.configure(text='We Found No Sanity Issues.')
    else:
        counts = df['Sanity Check'].value_counts()
        false_count = counts[False]
        problem_lab.configure(text= str(false_count) +' Sanity Issue(s) Found. ')

def open_file():
    try:
        filepath = askopenfilename()
        problem_lab.config(text="")
        success_lab.config(text="")
        if filepath is not None:
            process_file(filepath)
    except:
        processsed_lab.configure(text='File Cannot Be Processed!')


def save_file():
    save_path = askdirectory()
    try:
        writer = pd.ExcelWriter(save_path+"/data_processed.xlsx") 
        df.to_excel(writer, sheet_name='sheet1', index=False, na_rep='NaN')
        for column in df:
            column_length = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['sheet1'].set_column(col_idx, col_idx, column_length)
        writer.save() 
        success_lab.configure(text='File Saved Sucessfully!')
    except:
        success_lab.configure(text='File Not Saved!')

hello_lab = Label(ws, text='Clean Shift Data and Perform Sanity Checks')
hello_lab.grid(row=1, column=0, padx=10, pady=10)

choose_lab = Label(ws, text='Upload file in excel format ')
choose_lab.grid(row=2, column=0, padx=10)

processsed_lab = Label(ws)
processsed_lab.grid(row=3,column=0, padx=10)

problem_lab = Label(ws)
problem_lab.grid(row=4,column=0, padx=10)

save_lab = Label(ws, text='Save Processed file')
save_lab.grid(row=5, column=0, padx=10)

success_lab = Label(ws)
success_lab.grid(row=6,column=0, padx=10)

choosebtn = Button(ws, text ='Choose File', command = lambda:open_file()) 
choosebtn.grid(row=2, column=1)

save_btn = Button(ws, text ='Choose Location ', state = "disabled", command = lambda: save_file())
save_btn.grid(row=5, column=1)

ws.mainloop()