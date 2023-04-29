# Version 0.03
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

def calc_holiday_hours(df_payroll):
        years = set([min([df_payroll['CIDT'][i].year for i in range(len(df_payroll))]), max([df_payroll['CODT'][i].year for i in range(len(df_payroll))])])
        duration_range = [Range(start = df_payroll['CIDT'][i].to_pydatetime(), end = df_payroll['CODT'][i].to_pydatetime()) for i in range(len(df_payroll))]
        ahh_list = approved_holiday_hours(years)
        holiday_hours = [work_holiday_overlap(duration_range[i], ahh_list) for i in range(len(df_payroll))]
        df_payroll['Holiday Hours'] = holiday_hours
        
def is_manager(name, manager_rates):
    return name in set(manager_rates['Name'])

def manager_is_exempt(name, df_payroll):
    exempt_mins = df_payroll.loc[(df_payroll['Service Provider'] == name) & (df_payroll['Shift Code'] != 'MGR Direct Care')]['Staff Worked Duration (Minutes)'].sum()
    non_exempt_mins = df_payroll.loc[(df_payroll['Service Provider'] == name) & (df_payroll['Shift Code'] == 'MGR Direct Care')]['Staff Worked Duration (Minutes)'].sum()
    return (exempt_mins >= non_exempt_mins)

def worked_overtime(name, df_payroll):
    return df_payroll.loc[df_payroll['Service Provider'] == name]['Staff Worked Duration (Minutes)'].sum() > (40*60)

def non_manager_payroll(non_mgr, df_payroll, PAY_PERIOD):
    TGW = [] #Total Gross Wages
    THW = [] #Total Hours Worked
    THA = [] #Total Hours Accured
    for i, name in enumerate(non_mgr):
        df_indiv = df_payroll.loc[df_payroll['Service Provider'] == name]
        total_hours_worked = df_indiv['Staff Worked Duration (Minutes)'].sum()/60
        THW.append(total_hours_worked)
        THA.append((df_indiv['Staff Worked Duration (Minutes)']/60 * df_indiv['Accrue Rate']).sum())
        base_salary = (df_indiv['Regular Hourly Wage']*df_indiv['Staff Worked Duration (Minutes)']/60).sum()
        blended_overtime = 0
        if worked_overtime(name, df_payroll):
            BOT_base_rate = (df_indiv['BOT Hourly Wage']*df_indiv['Staff Worked Duration (Minutes)']/60).sum()/total_hours_worked
            blended_overtime = 0.5 * BOT_base_rate * (total_hours_worked-40)
        holiday_bonus = 0.5 * (df_indiv['Holiday Hours']*df_indiv['Regular Hourly Wage']).sum() 
        TGW.append(base_salary + blended_overtime + holiday_bonus)
    return pd.DataFrame({'Employee Name': non_mgr, 'Pay Peiord': [PAY_PERIOD]*len(THA), 'Total Gross Wages': TGW, 'Total Hours Worked': THW, 'Accured Time Off':THA})

def manager_payroll(mgr, manager_rates, df_payroll, PAY_PERIOD):
    TGW = [] #Total Gross Wages
    THW = [] #Total Hours Worked
    THA = [] #Total Hours Accured
    Exempt = [] #Exempt status
    for name in mgr:
      df_indiv = df_payroll.loc[df_payroll['Service Provider'] == name]
      total_hours_worked = df_indiv['Staff Worked Duration (Minutes)'].sum()/60
      exempt_hours_worked = df_indiv.loc[df_indiv['Shift Code'] != 'MGR Direct Care']['Staff Worked Duration (Minutes)'].sum()/60
      non_exempt_rate = manager_rates.loc[manager_rates['Name'] == name]['Non-exempt Hourly Wage'][0]
      THW.append(total_hours_worked)
      THA.append((df_indiv['Staff Worked Duration (Minutes)']/60 * df_indiv['Accrue Rate']).sum())
      if manager_is_exempt(name, df_payroll):
        MGR_weekly_salary = manager_rates.loc[manager_rates['Name'] == name]['Exempt Weekly Wage'][0]
        if name == 'Mikayla Napier':
          MGR_weekly_salary = manager_rates.loc[manager_rates['Name'] == name]['Exempt Biweekly Wage'][0]
        non_exempt_salary = (df_indiv['Regular Hourly Wage']*(total_hours_worked - exempt_hours_worked)).sum()
        holiday_bonus = 0.5 * (df_indiv['Holiday Hours'] * non_exempt_rate).sum()
        TGW.append(max(MGR_weekly_salary, non_exempt_salary) + holiday_bonus)
        Exempt.append('Exempt')
      else: #non-exempt
        base_salary = (non_exempt_rate * df_indiv['Staff Worked Duration (Minutes)']/60).sum()
        blended_overtime = 0
        if worked_overtime(name, df_payroll):
            BOT_base_rate = non_exempt_rate
            blended_overtime = 0.5 * BOT_base_rate * (total_hours_worked - 40)
        holiday_bonus = 0.5 * (df_indiv['Holiday Hours'] * non_exempt_rate).sum()
        TGW.append(base_salary + blended_overtime + holiday_bonus)
        Exempt.append('Non-exempt')
    return pd.DataFrame({'Employee Name': mgr, 'Pay Peiord': [PAY_PERIOD]*len(THA), 'Exempt Status': Exempt, 'Total Gross Wages': TGW, 'Total Hours Worked': THW, 'Accured Time Off':THA})
