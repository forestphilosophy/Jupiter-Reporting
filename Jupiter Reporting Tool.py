import pandas as pd

Missing_hours = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Missings 22-02-2021.xlsx",sheet_name="Data")

Parked_hours = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Parked 22-02-2021.xlsx",sheet_name="Data incl. comments on hours")

Employee_levels = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Table payscale group vs level.xlsx")
Employee_levels = Employee_levels.iloc[1:,1:]
Employee_levels.columns = Employee_levels.iloc[0]
Employee_levels = Employee_levels[1:]

Cyber_EGM = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Cyber GM analyse.xlsx",sheet_name="QRT004")
col_start = int(Cyber_EGM.iloc[5][Cyber_EGM.iloc[5].str.find("GM UHC") == 0].index.to_list()[0].split()[-1])
col_end = int(Cyber_EGM.iloc[5][Cyber_EGM.iloc[5].str.find("Total Chargeable Hours") == 0].index.to_list()[0].split()[-1])
Cyber_EGM = Cyber_EGM.iloc[5:,col_start:col_end+1]
Cyber_EGM.columns = Cyber_EGM.iloc[0]
Cyber_EGM = Cyber_EGM[1:]

Details = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Opport per KRA and INDUSTRY Final.xlsx",sheet_name="data")
People = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Opport per KRA and INDUSTRY Final.xlsx",sheet_name="People")
People = People.dropna(how='all')
KRA_map = dict(People[['Naam schoon','Service Line']].values)
Details['KRA'] = [KRA_map[x] if x in KRA_map else 'not found' for x in Details['Opportunity Leader'] ]

Net_revenue = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/GM analyse Risk Advisory NL (Maart).xlsx",sheet_name="QRT004")
col_start = int(Net_revenue.iloc[5][Net_revenue.iloc[5].str.find("GM UHC") == 0].index.to_list()[0].split()[-1])
col_end = int(Net_revenue.iloc[5][Net_revenue.iloc[5].str.find("Total Chargeable Hours") == 0].index.to_list()[0].split()[-1])
Net_revenue = Net_revenue.iloc[5:,col_start:col_end+1]
Net_revenue.columns = Net_revenue.iloc[0]
Net_revenue = Net_revenue[1:]

KRA = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Bestand KRA's (Test).xlsx")

Targets_KRA = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/FY21 hours targets from 6+6.xlsx",skiprows=2)
Targets_KRA = Targets_KRA.iloc[:,1:]
Targets_KRA.columns = Targets_KRA.iloc[0]
Targets_KRA = Targets_KRA[1:]

Employee = pd.read_excel("C:/Users/jimmlin/OneDrive - Deloitte (O365D)/Desktop/Jupiter Report/Employee  (Juist medewerker).xlsx",sheet_name="QRT006")

col_start = int(Employee.iloc[6][Employee.iloc[6].str.find("Profit Centre Hierarchy Level 4") == 0].index.to_list()[0].split()[-1])
col_end = int(Employee.iloc[6][Employee.iloc[6].str.find("YTD FTE") == 0].index.to_list()[0].split()[-1])
Employee = Employee.iloc[6:,col_start:col_end+1]
Employee.columns = Employee.iloc[0]
Employee = Employee[1:]

Table = {'Identified':1, 'Contacted':2,'Qualified':3,'Request to Propose':4,'Proposal Submitted':5,'Orals':6,'Verbal Commit':7}
Table = pd.DataFrame.from_dict(list(Table.items()))
Table.columns = ['Stage','Order']
