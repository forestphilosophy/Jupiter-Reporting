import pandas as pd
from datetime import datetime,date, timedelta
from time import strptime

Missing_hours_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Missings 22-02-2021.xlsx"
Parked_hours_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Parked 22-02-2021.xlsx"
Employee_levels_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Table payscale group vs level.xlsx"
Cyber_EGM_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Cyber GM analyse.xlsx"
Details_and_people_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Opport per KRA and INDUSTRY Final.xlsx"
Net_revenue_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/GM analyse Risk Advisory NL (Maart).xlsx"
KRA_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Bestand KRA's (Test).xlsx"
Targets_KRA_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/FY21 hours targets from 6+6.xlsx"
Employee_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Employee  (Juist medewerker).xlsx"

#Loading the missing hours source table
Missing_hours = pd.read_excel(Missing_hours_directory,sheet_name="Data")
Year = [str(date.year)[-2:] for date in Missing_hours['To Date']]
Missing_hours['Month & Year'] = list(map('-'.join, zip(Missing_hours['Month'], Year)))

#Loading the parked hours source table
Parked_hours = pd.read_excel(Parked_hours_directory,sheet_name="Data incl. comments on hours")

#Loading the employee levels source table
Employee_levels = pd.read_excel(Employee_levels_directory)
Employee_levels = Employee_levels.iloc[1:,1:]
Employee_levels.columns = Employee_levels.iloc[0]
Employee_levels = Employee_levels[1:]

#Loading the Cyber EGM source table
Cyber_EGM = pd.read_excel(Cyber_EGM_directory,sheet_name="QRT004")
col_start = int(Cyber_EGM.iloc[5][Cyber_EGM.iloc[5].str.find("GM UHC") == 0].index.to_list()[0].split()[-1])
col_end = int(Cyber_EGM.iloc[5][Cyber_EGM.iloc[5].str.find("Total Chargeable Hours") == 0].index.to_list()[0].split()[-1])
Cyber_EGM = Cyber_EGM.iloc[5:,col_start:col_end+1]
Cyber_EGM.columns = Cyber_EGM.iloc[0]
Cyber_EGM = Cyber_EGM[1:]

#Loading the Details source table
Details = pd.read_excel(Details_and_people_directory,sheet_name="data")
People = pd.read_excel(Details_and_people_directory,sheet_name="People")
People = People.dropna(how='all')
KRA_map = dict(People[['Naam schoon','Service Line']].values)
Details['KRA'] = [KRA_map[x] if x in KRA_map else 'not found' for x in Details['Opportunity Leader']]
Details['Quarters'] = pd.PeriodIndex(Details['Expected Close Date'], freq='Q')
bins = [0, 25, 50, 75, 100]
Details['Probability buckets'] = pd.cut(Details['Probability (%)'],bins,labels=['0 to 25%', '25% to 50%', '50% to 75%', '75% to 100%'])
Details['Probability buckets'] = Details['Probability buckets'].astype(object).fillna("No probability input")

#Loading the net revenue source table
Net_revenue = pd.read_excel(Net_revenue_directory,sheet_name="QRT004")
col_start = int(Net_revenue.iloc[5][Net_revenue.iloc[5].str.find("GM UHC") == 0].index.to_list()[0].split()[-1])
col_end = int(Net_revenue.iloc[5][Net_revenue.iloc[5].str.find("Total Chargeable Hours") == 0].index.to_list()[0].split()[-1])
Net_revenue = Net_revenue.iloc[5:,col_start:col_end+1]
Net_revenue.columns = Net_revenue.iloc[0]
Net_revenue = Net_revenue[1:]

#Loading the KRA source table
KRA = pd.read_excel(KRA_directory)

#Loading the Target_KRA source table
Targets_KRA = pd.read_excel(Targets_KRA_directory,skiprows=2)
Targets_KRA = Targets_KRA.iloc[:,1:]
Targets_KRA.columns = Targets_KRA.iloc[0]
Targets_KRA = Targets_KRA[1:]

#Loading the employee source table
Employee = pd.read_excel(Employee_directory,sheet_name="QRT006")
col_start = int(Employee.iloc[6][Employee.iloc[6].str.find("Profit Centre Hierarchy Level 4") == 0].index.to_list()[0].split()[-1])
col_end = int(Employee.iloc[6][Employee.iloc[6].str.find("YTD FTE") == 0].index.to_list()[0].split()[-1])
Employee = Employee.iloc[6:,col_start:col_end+1]
Employee.columns = Employee.iloc[0]
Employee = Employee[1:]

#Creating the Table source table
Table = {'Identified':1, 'Contacted':2,'Qualified':3,'Request to Propose':4,'Proposal Submitted':5,'Orals':6,'Verbal Commit':7}
Table = pd.DataFrame.from_dict(list(Table.items()))
Table.columns = ['Stage','Order']

#Making the dates table
if datetime.now().month <= 5:
    year = datetime.now().year - 1
    sdate = date(year, 6, 1)   # fiscal start date
    edate = date(year+1, 5, 31)   # fiscal end date
    
else:
    year = datetime.now().year
    sdate = date(year, 6, 1)   # fiscal start date
    edate = date(year+1, 5, 31)   # fiscal end date
    
delta = edate - sdate       # as timedelta

dates = []
for i in range(delta.days + 1):
    dates.append(sdate + timedelta(days=i))

Dates = pd.DataFrame()
Dates['Dates'] = dates
Dates['Month & Year'] = [date.strftime("%b") + '-' + date.strftime("%y") for date in Dates['Dates']]
month_mapping = {'Jun':1, 'Jul':2, 'Aug':3, 'Sep':4, 'Oct':5, 'Nov':6, 'Dec':7, 'Jan':8, 'Feb':9, 'Mar':10, 'Apr':11, 'May':12}
Dates['FiscalMIndex'] = [month_mapping[Dates['Month & Year'][i].split('-')[0]] for i in range(len(Dates))]

#Making the Month Order table
Month_order = pd.DataFrame()
if datetime.now().month <= 5:
    year = datetime.now().year - 1
    sdate = date(year, 6, 1)   # fiscal start date
    edate = date(year+1, 5, 31)   # fiscal end date
else:
    year = datetime.now().year
    sdate = date(year, 6, 1)   # fiscal start date
    edate = date(year+1, 5, 31)   # fiscal end date
        
month_list = pd.date_range(datetime.now(),edate).strftime("%B").unique().tolist()
month_idx = [strptime(month,'%B').tm_mon for month in month_list]
Month_order['Month'] = month_list
Month_order['Order'] = month_idx
