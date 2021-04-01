import pandas as pd
from datetime import datetime, date, timedelta
from time import strptime
import numpy as np
from os import listdir

data_folder_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/"

Missing_hours_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Missings 22-02-2021.xlsx"
Parked_hours_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Parked 22-02-2021.xlsx"
Employee_levels_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Table payscale group vs level.xlsx"
Cyber_EGM_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Cyber GM analyse.xlsx"
Details_and_people_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Opport per KRA and INDUSTRY Final.xlsx"
Net_revenue_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/GM analyse Risk Advisory NL (Maart).xlsx"
KRA_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Bestand KRA's (Test).xlsx"
Targets_KRA_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/FY21 hours targets from 6+6.xlsx"
Employee_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Employee  (Juist medewerker).xlsx"

def get_date(file_name):
    """
    Function to get the dates from the missing hours and parked hours filenames for sorting later.
    """
    return datetime.strptime(file_name.split()[1], '%d-%m-%Y')

def load_missing_hours(filename):
    """
    Function to perform loading and preprocessing of the missing hours files. 
    """
    Missing_hours = pd.read_excel(data_folder_directory + filename,sheet_name="Data")
    Year = [str(date.year)[-2:] for date in Missing_hours['To Date']]
    Missing_hours['Month & Year'] = list(map('-'.join, zip(Missing_hours['Month'], Year)))
    Missing_hours['Month & Year'] = Missing_hours['Month & Year'].str.title()
    
    return Missing_hours


list_of_files = listdir(data_folder_directory)

# Creating the missing hours tables
missing_hours_files = [f for f in list_of_files if "Missings" in f]

dates_dict = {}

for i in missing_hours_files:
    dates_dict[i] = get_date(i)

sorted_missing_hours = sorted(dates_dict, key = lambda x: x[1])

new_missing_hours = sorted_missing_hours[-1]
old_missing_hours = sorted_missing_hours[-2]

New_missing_hours = load_missing_hours(new_missing_hours)
Old_missing_hours = load_missing_hours(old_missing_hours)

# Creating the parked hours tables 
parked_hours_files = [f for f in list_of_files if "Parked" in f]

dates_dict = {}

for i in parked_hours_files:
    dates_dict[i] = get_date(i)

sorted_parked_hours = sorted(dates_dict, key = lambda x: x[1])

new_parked_hours = sorted_parked_hours[-1]
old_parked_hours = sorted_parked_hours[-2]

New_parked_hours = pd.read_excel(data_folder_directory + new_parked_hours,sheet_name="Data incl. comments on hours")
Old_parked_hours = pd.read_excel(data_folder_directory + old_parked_hours,sheet_name="Data incl. comments on hours")

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
Dates['Month'] = [Dates['Dates'][i].strftime("%B") for i in range(len(Dates))]
Dates['Month short'] = [Dates['Dates'][i].strftime("%b") for i in range(len(Dates))]

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

end_fiscal = datetime(fiscal_end_year, 6, 1)

#Getting a list of irrelevant columns that need to be dropped from the dataset, i.e. columns that have dates that are outside of range between the current month and end of current fiscal year

columns_to_drop = [temp.columns[i] for i in range(len(temp.columns)) if not ((datetime.strptime(temp.columns[i][0:11], '%d-%b-%Y') < end_fiscal) and (int(temp.columns[i][-2:]) >= date.today().isocalendar()[1]))]
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
        
        col_before = df[col].copy()
        df[col] *= (1 - days_diff / 5)
        col_after = df[col].copy()
        differences = [before - after for before, after in zip(col_before, col_after)]

        try:
            idx = df.columns.get_loc(col)
            next_col = df.columns[idx + 1]
            df[next_col] = [x + y for x,y in zip(df[next_col],differences)]
    
        except:
            continue

#Unpivot the dataframe to be able to visualize it
STAFFIT_report = pd.melt(df, id_vars=['Practitioner Name','Local Client ID','Client Description','Request Name', 'Booking Type', 'Cost Centre Desc',
       'EMP ID', 'Cost Centre', 'Capability Desc', 'Role Number',
       'Resource Requester', 'Engagement Manager', 'Business Desc Demand',
       'Role Name', 'Business Line', 'Assignment Type', 'Allocation Type',
       'Assignment Start Date', 'Assignment End Date', 'Engagement Code',
       'Engagement Description', 'Engagement Industry', 'Resource Manager',
       'Staffing Region', 'Federal Account', 'Global Level', 'Local Level','Request Number','Office','Practice'], var_name='Workweek', value_name='Planning Hours')
STAFFIT_report['Month'] = [datetime.strptime(STAFFIT_report['Workweek'][i][0:11], '%d-%b-%Y').strftime("%B") for i in range(len(STAFFIT_report))]

import pandas as pd
from datetime import datetime, timedelta
import collections

#Creating the table for Staffit report file for planning
Jupiter_pipeline = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Jupiter pipeline 20210317.xlsx")
Jupiter_pipeline = Jupiter_pipeline[Jupiter_pipeline['Stage'].isin(['Orals','Contacted','Identified','Proposal Submitted', 'Qualified', 'Request to Propose', 'Verbal Commit'])]
Jupiter_pipeline = Jupiter_pipeline[Jupiter_pipeline['Engagement Duration (Days)'].notnull()]
Jupiter_pipeline = Jupiter_pipeline.reset_index(drop=True)

#First, we want to get all the months that are relevant within our dataset, this is stored in the counter dict which will be used later to track the number of days in each month.
result = []
for i in range(len(Jupiter_pipeline)):
    result += pd.date_range(Jupiter_pipeline['Project Start Date'][i],Jupiter_pipeline['Project End Date'][i], 
                  freq='MS').strftime("%B-%y").tolist()
relevent_months = list(set(result))
counter = dict.fromkeys(relevent_months, [])

def number_of_days(start_date, end_date):
    """
    This function is used to get the dictionary with key being the month and value being the number of days that are in that month.
    The function takes in the starting date and ending date and returns a dictionary which shows the number of days for each month within that time range.
    """
    month_dict = collections.defaultdict(int)
    date = start_date
    
    while date <= end_date:
        key = '{}-{}'.format(date.strftime('%B'),str(date.year)[-2:])

        month_dict[key] += 1
        date += timedelta(days=1)

    return month_dict

for i in range(len(Jupiter_pipeline)):
    start_date = Jupiter_pipeline['Project Start Date'][i]
    end_date = Jupiter_pipeline['Project End Date'][i]
    #Calling the number_of_days function to get the number of days for relevant months for each row in the dataset
    relevant_months_dict = number_of_days(start_date, end_date)
    
    for k,v in counter.items():
        if k in relevant_months_dict.keys():
            counter[k] = counter[k] + [relevant_months_dict[k]]
        else:
            counter[k] = counter[k] + [0]
            
temp = pd.DataFrame.from_dict(counter)
for i in range(len(temp)):
    temp.iloc[i] = temp.iloc[i] * (Jupiter_pipeline['Weighted Split Amount (converted)'][i] / Jupiter_pipeline['Engagement Duration (Days)'][i])
    
relevant_df = Jupiter_pipeline[['Opportunity Leader', 'Split Leader', 'Industry', 'Sector',
       'Account Name', 'Opportunity ID', 'Opportunity Name',
       'Opportunity Split: Opportunity Split ID', 'Stage','Client Service Level 1', 'Client Service L1', 'Client Service Level 2',
       'Client Service Level 3', 'Client Service Level 4', 'Function Level 1']]

#Merging the two dataframes together to get the final dataframe 
final_df = pd.concat([relevant_df, temp], axis=1)
#Unpivot the dataframe so that we can visualize it in PowerBI
final_df = pd.melt(final_df, id_vars=['Opportunity Leader', 'Split Leader', 'Industry', 'Sector',
       'Account Name', 'Opportunity ID', 'Opportunity Name',
       'Opportunity Split: Opportunity Split ID', 'Stage','Client Service Level 1', 'Client Service L1', 'Client Service Level 2',
       'Client Service Level 3', 'Client Service Level 4', 'Function Level 1'], var_name='Month', value_name='Expected Revenue')

#Creating the MIndex column so that we can sort the month column in PowerBI
ranked_months = sorted(final_df['Month'], key = lambda x: (int(x.split('-')[1]), datetime.strptime(x.split('-')[0], "%B").month))
ranked_months = list(dict.fromkeys(ranked_months))

MIndex_dict = {}
acc = 1
for i in ranked_months:
    MIndex_dict[i] = acc
    acc += 1
    
final_df['Month_MIndex'] = [MIndex_dict[final_df['Month'][i]] for i in range(len(final_df))]
final_df['Dates'] = [datetime.strptime(final_df['Month'][i], "%B-%y").strftime("%m-%d-%Y") for i in range(len(final_df))]

import numpy as np
import pandas as pd
from datetime import datetime
from os import listdir

data_folder_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/"
list_of_files = listdir(data_folder_directory)

#Grabbing the month number from the file name
month = [f for f in list_of_files if "GM analyse QRT004 - UHC Profitability-MTD" in f][0][-11:-9]
#Converting the month number to month name
month_name = datetime.strptime(month, "%m").strftime("%B")

#Loading in the EGM_MTD table
EGM_MTD = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/GM analyse QRT004 - UHC Profitability-MTD29032021.xlsx",sheet_name='QRT004')
col_start = np.where(EGM_MTD.iloc[5].str.find("GM UHC") == 0)[0][0]
col_end = np.where(EGM_MTD.iloc[5].str.find("Total Chargeable Hours") == 0)[0][0]
EGM_MTD = EGM_MTD.iloc[5:,col_start:col_end+1]
EGM_MTD.columns = EGM_MTD.iloc[0]
EGM_MTD = EGM_MTD[1:]
EGM_MTD['Month'] = [month_name for i in range(len(EGM_MTD))]

#Loading in the EGM_YTD table
EGM_YTD = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/GM analyse QRT004 - UHC Profitability-YTD29032021.xlsx",sheet_name='QRT004')
col_start = np.where(EGM_YTD.iloc[5].str.find("GM UHC") == 0)[0][0]
col_end = np.where(EGM_YTD.iloc[5].str.find("Total Chargeable Hours") == 0)[0][0]
EGM_YTD = EGM_YTD.iloc[5:,col_start:col_end+1]
EGM_YTD.columns = EGM_YTD.iloc[0]
EGM_YTD = EGM_YTD[1:]

import numpy as np
import pandas as pd
from datetime import datetime
from os import listdir

def get_date(file_name):
    """
    Function to get the dates from the missing hours and parked hours filenames for sorting later.
    """
    return datetime.strptime(file_name.split()[1], '%d-%m-%Y')

def load_missing_hours(filename):
    """
    Function to perform loading and preprocessing of the missing hours files. 
    """
    Missing_hours = pd.read_excel(data_folder_directory + filename,sheet_name="Data")
    Year = [str(date.year)[-2:] for date in Missing_hours['To Date']]
    Missing_hours['Month & Year'] = list(map('-'.join, zip(Missing_hours['Month'], Year)))
    Missing_hours['Month & Year'] = Missing_hours['Month & Year'].str.title()
    
    return Missing_hours

data_folder_directory = "C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/"
list_of_files = listdir(data_folder_directory)

missing_hours_files = [f for f in list_of_files if "Missings" in f]

dates_dict = {}

for i in missing_hours_files:
    dates_dict[i] = get_date(i)

sorted_missing_hours = sorted(dates_dict, key = lambda x: x[1])

new_missing_hours = sorted_missing_hours[-1]
old_missing_hours = sorted_missing_hours[-2]

New_missing_hours = load_missing_hours(new_missing_hours)
Old_missing_hours = load_missing_hours(old_missing_hours)

parked_hours_files = [f for f in list_of_files if "Parked" in f]

dates_dict = {}

for i in parked_hours_files:
    dates_dict[i] = get_date(i)

sorted_parked_hours = sorted(dates_dict, key = lambda x: x[1])

new_parked_hours = sorted_parked_hours[-1]
old_parked_hours = sorted_parked_hours[-2]

New_parked_hours = pd.read_excel(data_folder_directory + new_parked_hours,sheet_name="Data incl. comments on hours")
Old_parked_hours = pd.read_excel(data_folder_directory + old_parked_hours,sheet_name="Data incl. comments on hours")
