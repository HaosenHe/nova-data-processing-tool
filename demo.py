# Version 0.0.1
# April 20, 2023
import re # Python regular expression support
from datetime import datetime # Python datetime conversion support
from concurrent.futures import process # Asynchronous callable execution
from tkinter import * # For user interface
from tkinter.filedialog import askopenfile, askopenfilename, askdirectory,asksaveasfile # UI-file system interaction
import pandas as pd # For importing, manipulating, and exporting data
#######################################
#Backend
df = pd.DataFrame()
rates = pd.DataFrame()

def open_shift_file():
    global df
    global rates
    filepath =''
    read_shift_lab.configure(text='')
    try:
        processed_lab.config(text="")
        save_lab.config(text="")
        filepath = askopenfilename()
        df=pd.read_excel(filepath)
        try:    
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
            read_shift_lab.configure(text='Shift record loaded successfully.')
        except:
            read_shift_lab.configure(text='The file does not contain all columns needed.')
    except:
        if filepath!='':
            read_shift_lab.configure(text='The file must be an excel.')

def open_rates_file():
    global df
    global rates
    filepath =''
    read_wage_lab.configure(text='')
    try:
        processed_lab.config(text='')
        save_lab.config(text='')
        filepath = askopenfilename()
        rates=pd.read_excel(filepath)
        try:    
            rates = rates[['Service 1 Description (Code)', 
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
            read_wage_lab.configure(text='Shift record loaded successfully.')
        except:
            read_wage_lab.configure(text='The file does not contain all columns needed.')
    except:
        if filepath!='':
            read_wage_lab.configure(text='The file must be an excel.')

def process_file():
    global df
    global rates
    processed_lab.configure(text='')
    # Remove parentheses and everything within them
    try:
        df['Service 1 Description (Code)'] = df['Service 1 Description (Code)'].str.replace(r'\(.*\)', '')
        # Remove prefix if it exists
        prefix = 'RC-SDP-CLS-320 '
        df['Service 1 Description (Code)'] = df['Service 1 Description (Code)'].apply(lambda x: x[len(prefix):] if x.startswith(prefix) else x)
        # Remove everything after " /"
        df['Service Provider'] = df['Service Provider'].str.split(' /', n=1).str[0]
    except:
        processed_lab.configure(text='Problem Detected in Service 1 Description (Code)/Service Provider')
        return
    # Replace Date/Time with Updated Date/Time if the latter is not NaN
    try:
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
    except:
        processed_lab.configure(text='Problem Detected in Check-In/Check-Out Dates and Times')
        return
    # Convert Staff Work Duration from Hour:Minutes to Minutes
    try:
        SWD_min = df['Staff Worked Duration'].apply(lambda x: (int(x.split(':')[0]) * 60) + int(x.split(':')[1]))
    except:
        processed_lab.configure(text='Problem Detected in Staff Worked Duration/Staff Worked Duration (Minutes)')
        return
    # Error check:
    # 1. Check if Staff Work Duration == Staff Work Duration (Minutes)
    error1 = (SWD_min == df['Staff Worked Duration (Minutes)']) #True means no error
    # 2. Check if |Staff Work Duration (Minutes) - Calculated Time Difference| <= 1
    error2 = ((df['Staff Worked Duration (Minutes)'] - CTD).abs() <= 1.1) # 1.1 to avoid float precision issues
    df["Passed Error Check"] = (error1 & error2) # The data is "sane" only when both checks are passed
    save_btn["state"] = "normal" # Enable saving botton
    if df['Passed Error Check'].all():
        processed_lab.configure(text='File Processed. We Found No Errors.')
    else:
        counts_1 = error1.value_counts()
        counts_2 = error2.value_counts()
        false_count_1 = counts_1[False]
        false_count_2 = counts_2[False]
        processed_lab.configure(text= f'File Processed. {false_count_1} Staff Work Duration Unequal to Staff Work Duration (Minutes).\n{false_count_2} Staff Work Duration (Minutes) Unequal to Calculated Time Difference.')

def save_file():
    global df
    global rates
    save_lab.configure(text='')
    save_path = asksaveasfile(mode='w', defaultextension=".xlsx").name
    print(save_path)
    if save_path is None: # asksaveasfile return `None` if dialog closed with "cancel".
        return
    try:
        writer = pd.ExcelWriter(save_path) 
        df.to_excel(writer, sheet_name='sheet1', index=False, na_rep='NaN')
        for column in df:
            column_length = max(df[column].astype(str).map(len).max(), len(column))
            col_idx = df.columns.get_loc(column)
            writer.sheets['sheet1'].set_column(col_idx, col_idx, column_length)
        writer.save() 
        save_lab.configure(text='File Saved Successfully!')
    except:
        save_lab.configure(text='File Not Saved!')

#########################################################
# UI
window = Tk()
window.title('Nova Support Home Data Processor (Demo)')
window.geometry("1280x832")


canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 832,
    width = 1280,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)
canvas.place(x = 0, y = 0)

canvas.create_rectangle(
    0.0,
    0.0,
    1280.0,
    136.0,
    fill="#1A8FDD",
    outline="")

canvas.create_text(
    55.0,
    12.0,
    anchor="nw",
    text="Nova Home Support Data Processing Tool",
    fill="#FFFFFF",
    font=("Inter", 40 * -1)
)

canvas.create_text(
    54.0,
    69.0,
    anchor="nw",
    text="Clean shift records, check for errors, and generate payroll.\nBeta version 0.2",
    fill="#FFFFFF",
    font=("Inter Bold", 20 * -1)
)

canvas.create_text(
    55.0,
    136.0,
    anchor="nw",
    text="Step 1: Upload Files\n",
    fill="#000000",
    font=("Inter SemiBold", 30 * -1)
)
canvas.create_text(
    55.0,
    176.0,
    anchor="nw",
    text="Click buttons below to upload shift records and the latest billing & wage rates in excel format. \n",
    fill="#000000",
    font=("Inter SemiBold", 20 * -1)
)

canvas.create_text(
    55.0,
    200.0,
    anchor="nw",
    text="Note: shift records must contain Service 1 Description (Code), Service Provider, Check-In Date, Check-In Time, Updated Check-In Date,\nUpdated Check-In Time, Check-Out Date, Check-Out Time, Updated Check-Out Date, Updated Check-Out Time, Staff Worked Duration,\nand Staff Worked Duration (Minutes).",
    fill="#000000",
    font=("Inter Light", 15 * -1)
)

canvas.create_text(
    55.0,
    416.0,
    anchor="nw",
    text="Step 2: Process File and Check for Errors\n",
    fill="#000000",
    font=("Inter SemiBold", 30 * -1)
)

canvas.create_text(
    55.0,
    456.0,
    anchor="nw",
    text="Click the button below to process shift records and generate the processed file. Results for error check will appear on the right.",
    fill="#000000",
    font=("Inter", 20 * -1)
)

canvas.create_text(
    55.0,
    614.0,
    anchor="nw",
    text="Step 3: Save Processed File\n",
    fill="#000000",
    font=("Inter SemiBold", 30 * -1)
)

canvas.create_text(
    55.0,
    654.0,
    anchor="nw",
    text="Click the button below to save the processed file.",
    fill="#000000",
    font=("Inter", 20 * -1)
)

# Initialize UI  Labels
cimg = PhotoImage(width=1, height=1)

read_shift_lab = Label(window, text="", font=('lucida', 20),bg = "white")
read_shift_lab.place(x=696.0,y=277.0)

read_wage_lab = Label(window, text="", font=('lucida', 20),bg = "white")
read_wage_lab.place(x=696.0,y=362.0)

processed_lab = Label(window, text="", font=('lucida', 20),bg = "white")
processed_lab.place(x=696.0,y=550.0)

save_lab = Label(window, text="", font=('lucida', 20), bg = "white")
save_lab.place(x=696.0, y=740.0)

# Initialize UI  Buttons
load_shift_btn = Button(window, text = 'upload shift records',command = lambda:open_shift_file(), image=cimg, width=324, height=43,compound='c',font=('lucida', 20))
load_billing_btn = Button(window, text = 'upload billing & wage rates',image=cimg, width=324, height=43,compound='c',font=('lucida', 20))
process_btn = Button(window, text = 'process file', command = lambda:process_file(), image=cimg, width=324, height=43,compound='c', font=('lucida', 20))
save_btn = Button(window, text = 'save as...', state = "disabled", command = lambda: save_file(), image=cimg,width=324, height=43,compound='c',font=('lucida', 20))

load_shift_btn.place(x=188.0,y=272.0)
load_billing_btn.place(x=188.0,y=347.0)
process_btn.place(x=188.0,y=541.0)
save_btn.place(x=188.0,y=735.0)

# Initialize User Interface
window.resizable(False, False)
window.mainloop()