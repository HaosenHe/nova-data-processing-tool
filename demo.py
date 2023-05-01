# Version 0.0.4
# April 28, 2023
import pandas as pd # For importing, manipulating, and exporting data
import re # Python regular expression support
import datetime # Python datetime conversion support
from concurrent.futures import process # Asynchronous callable execution
from dateutil.easter import * # Identify Easter Sunday
import holidays # Identify holiday date
from tkinter import * # For user interface
from tkinter.filedialog import askopenfile, askopenfilename, askdirectory,asksaveasfile # UI-file system interaction
from copy import deepcopy 
from collections import namedtuple # for date range object
from helpers import * #helper functions

# define time range object
Range = namedtuple('Range', ['start', 'end'])

# GLOBAL VARIABLES:
# shift data
df = pd.DataFrame()
# manager rates data
manager_rates = pd.DataFrame()
# non-manager rates data
non_manager_rates = pd.DataFrame()
#other wages
other_rates = pd.DataFrame()
# shift data merged with non-manager rates
df_aug = pd.DataFrame()
# payroll for non-managers
non_mgr_payroll = []
# payroll for managers
mgr_payroll = []
# a string indicating pay period
PAY_PERIOD = ''

def open_shift_file():
    '''
    Open an Excel file containing shift records and check formatting.
    '''
    global df
    filepath =''
    read_shift_lab.configure(text='')
    try:
        processed_lab.config(text="")
        save_lab.config(text="")
        filepath = askopenfilename()
        df=pd.read_excel(filepath)
        try:
            # Subsetting useful columns
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
            # Success message
            read_shift_lab.configure(text='Shift Record Loaded Successfully.')
        except:
            read_shift_lab.configure(text='The File Does Not Contain All Columns Needed.')
    except: # Non-excel
        if filepath != '': # File selection is not "cancelled"
            read_shift_lab.configure(text='The File Must Be An Excel.')
            

def open_rates_file():
    '''
    Open an Excel file containing manager and non-manager rates and check formatting.
    '''
    global manager_rates, non_manager_rates, other_rates
    filepath =''
    read_wage_lab.configure(text='')
    try:
        processed_lab.config(text='')
        save_lab.config(text='')
        filepath = askopenfilename()
        # Read "Manger Rates" sheet
        manager_rates = pd.read_excel(filepath, sheet_name="Manager Rates")
        # Read "Non-manager Rates" sheet
        non_manager_rates = pd.read_excel(filepath, sheet_name="Non-manager Rates")
        # Read "Other Wages"
        other_rates = pd.read_excel(filepath, sheet_name="Other Wages")
        try:
            non_manager_rates = non_manager_rates[['Shift Code', 'Regular Hourly Wage', 'BOT Hourly Wage', 'Accrual Rate']]
            manager_rates = manager_rates[['Name', 'Non-exempt Hourly Wage', 'Exempt Weekly Wage', 'Exempt Biweekly Wage', 'Accrual Rate']]
            read_wage_lab.configure(text='Billing & Wage Rates Loaded Successfully.')
        except:
            read_wage_lab.configure(text='The File Does Not Contain All Columns Needed.')
    except:
        if filepath=='':
            return
        else:
            read_wage_lab.configure(text='The File Must Be An Excel Containing Manager, Non-Manager Rates, and Other Wages.')

def process_shift():
    '''
    Process shift records
    '''
    global df, PAY_PERIOD
    processed_lab.configure(text='')
    # Remove parentheses and everything within them
    try:
        df['Service 1 Description (Code)'] = df['Service 1 Description (Code)'].str.replace(r'\(.*\)', '')
        # Remove prefix if it exists
        prefix = 'RC-SDP-CLS-320 '
        df['Service 1 Description (Code)'] = df['Service 1 Description (Code)'].apply(lambda x: x[len(prefix):] if x.startswith(prefix) else x)
        df['Service 1 Description (Code)'] = df['Service 1 Description (Code)'].apply(lambda x: x.rstrip())
        # Remove everything after " /"
        df['Service Provider'] = df['Service Provider'].str.split(' /', n=1).str[0]
    except:
        processed_lab.configure(text='Problem Detected in Service 1 Description (Code)/Service Provider')
        return
    try:
        # Replace Date/Time with Updated Date/Time if the latter is not NaN
        df['Check-In Date'] = df['Updated Check-In Date'].fillna(df['Check-In Date'])
        df['Check-In Time'] = df['Updated Check-In Time'].fillna(df['Check-In Time'])
        df['Check-Out Date'] = df['Updated Check-Out Date'].fillna(df['Check-Out Date'])
        df['Check-Out Time'] = df['Updated Check-Out Time'].fillna(df['Check-Out Time'])
        df.drop(['Updated Check-In Date', 'Updated Check-In Time','Updated Check-Out Date', 'Updated Check-Out Time'], axis=1, inplace=True)
        # Create datetime series in Python format
        CIDT = df['Check-In Date'].str.cat(df['Check-In Time'], sep=' ')
        CODT = df['Check-Out Date'].str.cat(df['Check-Out Time'], sep=' ')
        CIDT = CIDT.apply(lambda x: datetime.datetime.strptime(x, r'%m/%d/%Y %I:%M %p'))
        CODT = CODT.apply(lambda x: datetime.datetime.strptime(x, r'%m/%d/%Y %I:%M %p'))
        df['CIDT'] = CIDT
        df['CODT'] = CODT
        PAY_PERIOD = str(CIDT.min().date()) + ' - ' + str(CODT.max().date())
        # Calculate time difference from check-in and check-out datetimes
        CTD = (CODT - CIDT).dt.total_seconds() / 60
    except:
        processed_lab.configure(text='Problem Detected in Check-In/Check-Out Dates and Times')
        return
    # Convert Staff Work Duration from Hour:Minutes to Minutes
    try:
        SWD_min = df['Staff Worked Duration'].apply(lambda x: (int(x.split(':')[0]) * 60) + int(x.split(':')[1]))
    except:
        SWD_min = CTD
    df = df.rename(columns={'Service 1 Description (Code)': 'Shift Code'})
    # Error check:
 # 1. Check if Staff Work Duration ==  Staff Work Duration (Minutes)
    error1 = (SWD_min - df['Staff Worked Duration (Minutes)']).abs() <= 1.1
    # 2. Check if |Staff Work Duration (Minutes) - Calculated Time Difference| <= 1
    error2 = ((df['Staff Worked Duration (Minutes)'] - CTD).abs() <= 1.1) # 1.1 to avoid float precision issues
    df["Passed Error Check"] = (error1 & error2) # The data is "sane" only when both checks are passed
    save_btn["state"] = "normal" # Enable saving botton
    if df['Passed Error Check'].all():
        processed_lab.configure(text='File Processed. We Found No Errors.')
    else:
        false_count_1 = len(error1[error1 == False])
        false_count_2 = len(error2[error2 == False])
        processed_lab.configure(text= f'Shift File Processed. {false_count_1} Staff Work Duration Unequal to \nStaff Work Duration (Minutes).{false_count_2} Staff Work Duration (Minutes) Unequal \nto Calculated Time Difference.')

def process_payroll():
    '''
    Process payroll for managers and non-managers
    '''
    global df, manager_rates, non_manager_rates, other_rates, df_aug, non_mgr_payroll, mgr_payroll, PAY_PERIOD
    try:
        # prepare for merging
        df_aug = df.rename(columns={'Service 1 Description (Code)': 'Shift Code'})
        df_aug = deepcopy(df_aug.loc[~df_aug['Shift Code'].str.contains('Adaptive Behavior Treatment')])
        # split the 'Service Provider' column on comma separator and extract First name and Surname
        df_aug[['Surname', 'First Name']] = df_aug['Service Provider'].str.split(', ', expand=True)
        # concatenate First name and Surname columns in the desired order
        df_aug['Service Provider'] = df_aug['First Name'] + ' ' + df_aug['Surname']
        # drop First name and Surname columns
        df_aug.drop(columns=['First Name', 'Surname'], inplace=True)
        df_aug=pd.merge(df_aug, non_manager_rates, how='left', on='Shift Code')
        admin_shifts = df_aug[df_aug['Shift Code'] == 'Admin']
        admin_names = admin_shifts['Service Provider'].tolist()
        # Filter the rows in other_rates where the NAME column matches the names of the Service Providers in the Admin shifts
        admin_rates = other_rates[other_rates['Service Provider'].isin(admin_names)]\
        # Merge the Admin rates back into the Admin shifts dataframe
        admin_shifts_merged = pd.merge(admin_shifts, admin_rates, on='Service Provider', how='left')
        # Fill the Regular Hourly Wage and Overtime Hourly Wage columns with the values from the ADMIN/VACAY/SICK WAGE column
        admin_shifts_merged.loc[:, 'Regular Hourly Wage'] = admin_shifts_merged['ADMIN/VACAY/SICK WAGE']
        admin_shifts_merged.loc[:, 'BOT Hourly Wage'] = admin_shifts_merged['ADMIN/VACAY/SICK WAGE']
        admin_shifts_merged.loc[:, 'Accrual Rate'] = 0.04
        admin_shifts_merged = admin_shifts_merged.drop("ADMIN/VACAY/SICK WAGE", axis=1)
        df_aug = df_aug[df_aug['Shift Code'] != 'Admin']
        df_aug = pd.concat([df_aug, admin_shifts_merged], ignore_index=True)
        # append accure rates
        for name, accrued in zip(manager_rates['Name'], manager_rates['Accrual Rate']):
            df_aug.loc[(df_aug['Service Provider'] == name), ['Accrual Rate']] = accrued
        # add holiday hours
        calc_worked_holiday(df_aug)
        # identify staff roles:
        staff_names = df_aug['Service Provider'].unique()
        manager_status = [is_manager(i, manager_rates) for i in staff_names]
        non_mgr = [] #list of names
        mgr = [] #list of names
        for i, name in enumerate(staff_names):
            if manager_status[i]:
                mgr.append(name)
            else:
                non_mgr.append(name)
        # generate payrolls
        non_mgr_payroll = non_manager_payroll(non_mgr, df_aug, PAY_PERIOD)
        mgr_payroll = manager_payroll(mgr, manager_rates, df_aug, PAY_PERIOD)
    except:
        processed_lab.configure(text= 'Unknown Error Encountered When Processing Payroll.')

def process_file():
    '''
    Process shift and payroll in a row
    '''
    global df, manager_rates, non_manager_rates, df_aug, non_mgr_payroll, mgr_payroll, PAY_PERIOD
    process_shift()
    process_payroll()
    df_aug = df_aug.round(decimals=2)
    non_mgr_payroll = non_mgr_payroll.round(decimals=2)
    mgr_payroll = mgr_payroll.round(decimals=2)


def save_file():
    '''
    Save processed spreadsheets
    '''
    save_lab.configure(text='')
    save_path = asksaveasfile(mode='w', defaultextension=".xlsx").name
    print(save_path)
    if save_path is None: # asksaveasfile return `None` if dialog closed with "cancel".
        return
    try:
        writer = pd.ExcelWriter(save_path) 
        startrow = 0
        for person in [ *non_mgr_payroll, *mgr_payroll ]:
            for df in [person['header'], person['payroll'],person['summary']]:
                df.to_excel(writer, engine="xlsxwriter",sheet_name='Payroll', startrow=startrow, index=False)
                startrow += (df.shape[0] + 1)
            startrow += 3
        writer.sheets['Payroll'].set_column(0, 4, 28)

        df_aug.to_excel(writer, sheet_name='Shift with Pay Rates', index=False, na_rep='NaN')
        non_manager_rates.to_excel(writer, sheet_name='Non-Manager Pay Rates', index=False, na_rep='NaN')
        manager_rates.to_excel(writer, sheet_name='Manager Pay Rates', index=False, na_rep='NaN')
        
        for column in df_aug:
            column_length = max(df_aug[column].astype(str).map(len).max(), len(column))
            col_idx = df_aug.columns.get_loc(column)
            writer.sheets['Shift with Pay Rates'].set_column(col_idx, col_idx, column_length)
        
        for column in non_manager_rates:
            column_length = max(non_manager_rates[column].astype(str).map(len).max(), len(column))
            col_idx = non_manager_rates.columns.get_loc(column)
            writer.sheets['Non-Manager Pay Rates'].set_column(col_idx, col_idx, column_length)
        
        for column in manager_rates:
            column_length = max(manager_rates[column].astype(str).map(len).max(), len(column))
            col_idx = manager_rates.columns.get_loc(column)
            writer.sheets['Manager Pay Rates'].set_column(col_idx, col_idx, column_length)

        writer.save() 
        save_lab.configure(text='File Saved Successfully!')
    except:
        save_lab.configure(text='File Not Saved!')

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
    text="Clean shift records, check for errors, and generate payroll.\nBeta version 0.3",
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
read_shift_lab.place(x=600.0,y=277.0)

read_wage_lab = Label(window, text="", font=('lucida', 20),bg = "white")
read_wage_lab.place(x=600.0,y=362.0)

processed_lab = Label(window, text="", font=('lucida', 20),bg = "white")
processed_lab.place(x=600.0,y=550.0)

save_lab = Label(window, text="", font=('lucida', 20), bg = "white")
save_lab.place(x=600.0, y=740.0)

# Initialize UI  Buttons
load_shift_btn = Button(window, text = 'upload shift records',command = lambda:open_shift_file(), image=cimg, width=324, height=43,compound='c',font=('lucida', 20))
load_billing_btn = Button(window, text = 'upload billing & wage rates',command = lambda:open_rates_file(), image=cimg, width=324, height=43,compound='c',font=('lucida', 20))
process_btn = Button(window, text = 'process file', command = lambda:process_file(), image=cimg, width=324, height=43,compound='c', font=('lucida', 20))
save_btn = Button(window, text = 'save as...', state = "disabled", command = lambda: save_file(), image=cimg,width=324, height=43,compound='c',font=('lucida', 20))

load_shift_btn.place(x=188.0,y=272.0)
load_billing_btn.place(x=188.0,y=347.0)
process_btn.place(x=188.0,y=541.0)
save_btn.place(x=188.0,y=735.0)

# Initialize User Interface
window.resizable(False, False)
window.mainloop()