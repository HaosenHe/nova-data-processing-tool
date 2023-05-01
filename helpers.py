# Version 0.0.4
import pandas as pd # For importing, manipulating, and exporting data
import re # Python regular expression support
import datetime # Python datetime conversion support
from concurrent.futures import process # Asynchronous callable execution
from dateutil.easter import *
import holidays
from copy import deepcopy
from collections import namedtuple

Range = namedtuple('Range', ['start', 'end'])

def approved_holiday(YEARS):
    approved_holiday = set()
    holiday_set = set(["Thanksgiving", "Christmas Day"])
    for year in YEARS:
        ah = {k for k, v in holidays.US(years=year).items() if v in holiday_set}
        approved_holiday = approved_holiday.union(ah)
        approved_holiday.add(easter(year))
        approved_holiday.add(datetime.date(year, 12, 24))
        approved_holiday.add(datetime.date(year, 12, 31))
        approved_holiday.add(datetime.date(year-1, 12, 31))
    return approved_holiday

def approved_holiday_hours(YEARS):
    approved_holiday_dt = []
    for x in approved_holiday(YEARS):
        if (x.month == 12 and x.day == 31):
            approved_holiday_dt.append(Range(start = datetime.datetime.combine(x, datetime.time(hour=23)),end = datetime.datetime.combine(x+datetime.timedelta(days=1), datetime.time(hour=23))))
        else:
            approved_holiday_dt.append(Range(start = datetime.datetime.combine(x, datetime.time(hour=7)), end = datetime.datetime.combine(x, datetime.time(hour=23))))
    return approved_holiday_dt

def work_holiday_overlap(work_range, ahh_list):
    total_overlap = 0.0
    for ahh in ahh_list:
        latest_start = max(work_range.start, ahh.start)
        earliest_end = min(work_range.end, ahh.end)
        delta = (earliest_end - latest_start)
        overlap = max(0, delta.total_seconds())
        total_overlap += overlap
    return total_overlap/60 #in minutes

def calc_worked_holiday(df_shift_merged):
        years = set([min([df_shift_merged['CIDT'][i].year for i in range(len(df_shift_merged))]), max([df_shift_merged['CODT'][i].year for i in range(len(df_shift_merged))])])
        duration_range = [Range(start = df_shift_merged['CIDT'][i].to_pydatetime(), end = df_shift_merged['CODT'][i].to_pydatetime()) for i in range(len(df_shift_merged))]
        ahh_list = approved_holiday_hours(years)
        holiday_work_time = [work_holiday_overlap(duration_range[i], ahh_list) for i in range(len(df_shift_merged))]
        df_shift_merged['Holiday Worked Duration (Minutes)'] = holiday_work_time
        
def is_manager(name, manager_rates):
    return name in set(manager_rates['Name'])

def manager_is_exempt(name, df_shift_merged):
    exempt_mins = df_shift_merged.loc[(df_shift_merged['Service Provider'] == name) & (df_shift_merged['Shift Code'] != 'MGR Direct Care')]['Staff Worked Duration (Minutes)'].sum()
    non_exempt_mins = df_shift_merged.loc[(df_shift_merged['Service Provider'] == name) & (df_shift_merged['Shift Code'] == 'MGR Direct Care')]['Staff Worked Duration (Minutes)'].sum()
    return (exempt_mins >= non_exempt_mins)

def worked_overtime(name, df_shift_merged):
    return df_shift_merged.loc[df_shift_merged['Service Provider'] == name]['Staff Worked Duration (Minutes)'].sum() > (40*60)

def non_manager_payroll(non_mgr, df_shift_merged, PAY_PERIOD):   

    non_mgr_payroll = []

    for i, name in enumerate(non_mgr):
        
        df_indiv = df_shift_merged.loc[df_shift_merged['Service Provider'] == name]
        aggregations = {'Staff Worked Duration (Minutes)': 'sum', 'Regular Hourly Wage': 'first', 'Service Provider': 'first'}
        df_payroll = df_indiv[['Shift Code', 'Staff Worked Duration (Minutes)', 'Regular Hourly Wage', 'Service Provider']]
        df_payroll = df_payroll.groupby('Shift Code').agg(aggregations)
        df_payroll = df_payroll.reset_index()

        total_hours_worked = df_indiv['Staff Worked Duration (Minutes)'].sum()/60

        overtime_hours = max(0, total_hours_worked-40)
        if overtime_hours > 0:
            BOT_pay_rate = 0.5 * (df_indiv['BOT Hourly Wage']*df_indiv['Staff Worked Duration (Minutes)']/60).sum()/total_hours_worked
            df_overtime = pd.DataFrame({'Service Provider': name, 'Shift Code': ['Blended Overtime'], 'Staff Worked Duration (Minutes)': [overtime_hours], 
                                        'Regular Hourly Wage': [BOT_pay_rate]})
            df_payroll= pd.concat([df_payroll, df_overtime], ignore_index=True)
        
        if df_indiv['Holiday Worked Duration (Minutes)'].sum() > 0:
            df_holiday_pay = df_indiv[['Service Provider', 'Shift Code', 'Holiday Worked Duration (Minutes)', 'Regular Hourly Wage']]
            df_holiday_pay['Regular Hourly Wage'] = df_holiday_pay['Regular Hourly Wage']*0.5
            df_holiday_pay['Shift Code'] = df_holiday_pay['Shift Code'].apply(lambda x: x + ' HOL')
            df_holiday_pay = df_holiday_pay.rename(columns={'Holiday Worked Duration (Minutes)': 'Staff Worked Duration (Minutes)'})
            df_holiday_pay = df_holiday_pay.groupby('Shift Code').agg(aggregations)
            df_holiday_pay = df_holiday_pay.reset_index()
            df_payroll= pd.concat([df_payroll, df_holiday_pay], ignore_index=True)
        
        df_payroll = df_payroll.rename(columns={'Regular Hourly Wage': 'Hourly Rate'})
        df_payroll['Staff Worked Duration (Hours)'] = df_payroll['Staff Worked Duration (Minutes)']/60
        df_payroll = df_payroll.reindex(columns=['Service Provider', 'Shift Code', 'Staff Worked Duration (Minutes)', 'Staff Worked Duration (Hours)',  'Hourly Rate'])
        total_gross_wage = (df_payroll['Hourly Rate'] * df_payroll['Staff Worked Duration (Hours)']).sum()
        weekly_accrued_hours = (df_indiv['Staff Worked Duration (Minutes)']/60 * df_indiv['Accrual Rate']).sum()

        non_mgr_payroll.append({'header': pd.DataFrame(columns=[name]), 'summary': pd.DataFrame({'Total Hours Worked': total_hours_worked, 
                                    'Total Gross Wage': total_gross_wage, 'Weekly Accrued Hours': weekly_accrued_hours, 'Pay Period': PAY_PERIOD}, index=[0]).round(decimals=2), 
                                    'payroll': df_payroll.round(decimals=2)})      
    return non_mgr_payroll

def manager_payroll(mgr, manager_rates, df_shift_merged, PAY_PERIOD):
    mgr_payroll = []
    for name in mgr:
      df_indiv = df_shift_merged.loc[df_shift_merged['Service Provider'] == name]
      total_hours_worked = df_indiv['Staff Worked Duration (Minutes)'].sum()/60
      exempt_hours_worked = df_indiv.loc[df_indiv['Shift Code'] != 'MGR Direct Care']['Staff Worked Duration (Minutes)'].sum()/60
      non_exempt_rate = manager_rates.loc[manager_rates['Name'] == name]['Non-exempt Hourly Wage'].iloc[0]
      weekly_accrued_hours = (df_indiv['Staff Worked Duration (Minutes)']/60 * df_indiv['Accrual Rate']).sum()
      total_gross_wage = 0
      df_payroll = pd.DataFrame()
      if exempt_hours_worked >= (total_hours_worked - exempt_hours_worked): # Exempt
        MGR_weekly_salary = manager_rates.loc[manager_rates['Name'] == name]['Exempt Weekly Wage'].iloc[0]
        if name == 'Mikayla Napier':
          MGR_weekly_salary = manager_rates.loc[manager_rates['Name'] == name]['Exempt Biweekly Wage'].iloc[0]
        df_payroll = pd.DataFrame({'Service Provider': name, 'Shift Code': ['MGR Salary'], 'Staff Worked Duration (Minutes)': [60], 'Hourly Rate': [MGR_weekly_salary]})
        holiday_work_time = df_indiv['Holiday Worked Duration (Minutes)'].sum()
        if holiday_work_time > 0:
          df_holiday_pay = pd.DataFrame({'Service Provider': name, 'Shift Code': ['MGR HOL BON'], 'Staff Worked Duration (Minutes)': [holiday_work_time], 
                                         'Hourly Rate': [0.5 * non_exempt_rate]})
          df_payroll= pd.concat([df_payroll, df_holiday_pay], ignore_index=True)
      else: # Non-exempt
        aggregations = {'Staff Worked Duration (Minutes)': 'sum',  'Service Provider': 'first'}
        df_payroll = df_indiv[['Service Provider', 'Shift Code', 'Staff Worked Duration (Minutes)']]
        df_payroll = df_payroll.groupby('Shift Code').agg(aggregations)
        df_payroll = df_payroll.reset_index() 
        df_payroll['Regular Hourly Wage'] = [non_exempt_rate]*len(df_payroll)
        overtime_hours = max(0, total_hours_worked-40)
        if overtime_hours > 0:
          BOT_pay_rate = 0.5 * non_exempt_rate
          df_overtime = pd.DataFrame({'Service Provider': name, 'Shift Code': ['Blended Overtime'], 'Staff Worked Duration (Minutes)': [overtime_hours], 
                                      'Regular Hourly Wage': [BOT_pay_rate]})
          df_payroll= pd.concat([df_payroll, df_overtime], ignore_index=True)
        if df_indiv['Holiday Worked Duration (Minutes)'].sum() > 0:
            df_holiday_pay = df_indiv[['Shift Code', 'Holiday Worked Duration (Minutes)']]
            df_holiday_pay['Shift Code'] = df_holiday_pay['Shift Code'].apply(lambda x: x + ' HOL')
            df_holiday_pay = df_holiday_pay.rename(columns={'Holiday Worked Duration (Minutes)': 'Staff Worked Duration (Minutes)'})
            df_holiday_pay = df_holiday_pay.groupby('Shift Code').agg(aggregations)
            df_holiday_pay = df_holiday_pay.reset_index()
            df_holiday_pay['Regular Hourly Wage'] = [0.5 * non_exempt_rate] * len(df_holiday_pay)
            df_payroll= pd.concat([df_payroll, df_holiday_pay], ignore_index=True)
      df_payroll = df_payroll.rename(columns={'Regular Hourly Wage': 'Hourly Rate'})
      df_payroll['Staff Worked Duration (Hours)'] = df_payroll['Staff Worked Duration (Minutes)']/60
      df_payroll = df_payroll.reindex(columns=['Service Provider', 'Shift Code', 'Staff Worked Duration (Minutes)', 'Staff Worked Duration (Hours)',  'Hourly Rate'])
      total_gross_wage = (df_payroll['Hourly Rate'] * df_payroll['Staff Worked Duration (Hours)']).sum()
      mgr_payroll.append({'header': pd.DataFrame(columns=[name]), 
                          'summary': pd.DataFrame({'Total Hours Worked': total_hours_worked, 'Total Gross Wage': total_gross_wage, 'Weekly Accrued Hours': weekly_accrued_hours,
                                                    'Pay Period': PAY_PERIOD}, index=[0]).round(decimals=2), 
                                                    'payroll': df_payroll.round(decimals=2)})      
    return mgr_payroll