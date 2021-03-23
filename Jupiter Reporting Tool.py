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

#Making the quarters MIndex column
quarters_list = list(set([str(Details['Quarters'][i]) for i in range(len(Details))]))
quarters_list = sorted(quarters_list,key=lambda element: (int(element[0:4]), int(element[-1])))
quarters_mapping = {}
idx = 1
for i in quarters_list:
    quarters_mapping[i] = idx
    idx += 1
Details['QuartersMIndex'] = [quarters_mapping[str(Details['Quarters'][i])] for i in range(len(Details))]

#Making the probability buckets column
bins = [0, 25, 50, 75, 100]
Details['Probability buckets'] = pd.cut(Details['Probability (%)'],bins,labels=['0 to 25%', '25% to 50%', '50% to 75%', '75% to 100%'])
Details['Probability buckets'] = Details['Probability buckets'].astype(object).fillna(0)
#Making the probability buckets MIndex column
probability_Mindex = {0:1,'0 to 25%':2,'25% to 50%':3,'50% to 75%':4, '75% to 100%':5} 
Details['Probability Buckets MIndex'] = [probability_Mindex[Details['Probability buckets'][i]] for i in range(len(Details))]

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
    last_year = datetime.now().year - 1
    sdate = date(last_year, 6, 1)   # fiscal start date
    edate = date(last_year+1, 5, 31)   # fiscal end date
else:
    current_year = datetime.now().year
    sdate = date(current_year, 6, 1)   # fiscal start date
    edate = date(current_year+1, 5, 31)   # fiscal end date
        
month_list = pd.date_range(datetime.now(),edate).strftime("%B").unique().tolist()
month_idx = [strptime(month,'%B').tm_mon for month in month_list]
Month_order['Month'] = month_list
Month_order['Order'] = month_idx

#Creating the table for Staffit report file for planning
df = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Staffit report 1703.xlsx")

practice_idx = df.columns.get_loc("Practice")
temp = df.iloc[:,practice_idx+1:]

if datetime.now().month <= 5:
    fiscal_end_year = datetime.now().year
else:
    fiscal_end_year = datetime.now().year + 1

current_month = datetime(datetime.now().year, datetime.now().month, 1)
end_fiscal = datetime(fiscal_end_year, 6, 1)

#Getting a list of irrelevant columns that need to be dropped from the dataset, i.e. columns that have dates that are outside of range between the current month and end of current fiscal year

columns_to_drop = [temp.columns[i] for i in range(len(temp.columns)) if not (current_month <= datetime.strptime(temp.columns[i][0:11], '%d-%b-%Y') < end_fiscal)]
df = df.drop(columns_to_drop,axis=1).iloc[:-1,:]

practice_idx = df.columns.get_loc("Practice")
temp = df.iloc[:,practice_idx+1:]
df[temp.columns] = df[temp.columns].fillna(value=0)

for col in temp.columns:
    work_week = int(col[-2:])
    work_year = int(col[-9:-5])
    #Check if the starting date and the ending date of certain workweek are in different months. And if so, we need to split the hours.
    if date.fromisocalendar(work_year, work_week, 1).month != date.fromisocalendar(work_year, work_week, 5).month:
        days_diff = (pd.Period(col[0:11],freq='M').end_time.date() - date.fromisocalendar(work_year, work_week, 1)).days
        df[col] *= (1 - days_diff / 5)
        
        try:
            idx = df.columns.get_loc(col)
            next_col = df.columns[idx + 1]
            df[next_col] = [df[next_col][i] * (days_diff / 5) if df[col][i] != 0 else df[next_col][i] for i in range(len(df[next_col]))]
            
        except:
            continue
