from datetime import date
import pandas as pd
import dateutil
import warnings
from python_calamine import CalamineWorkbook
import time
import stat
import os
import re
import xlsxwriter

'''
DATA EXPORT LINKS:

Floyd:
https://floyd.verticalbridge.com/AdvancedSearch/Search#results=site&siteStatus=Active&siteCategories=Tower&siteClass=Owned&country=Canada&country=United%20States

    * Ensure Acquisition Name is selected in fields for export

https://floyd.verticalbridge.com/AdvancedSearch/Search#results=tenant&siteStatus=Active&siteCategories=Tower&siteClass=Owned&country=Canada&country=United%20States

SiteTracker 1:
https://sitetracker-verticalbridge.lightning.force.com/lightning/r/sitetracker__StGridView__c/a0q5G000007kuXJQAY/view

SiteTracker 2:
https://sitetracker-verticalbridge.lightning.force.com/lightning/r/sitetracker__StGridView__c/a0qVJ000006exh3YAA/view

Power BI:
https://app.powerbi.com/groups/me/apps/1a4ce392-ccfa-4b07-8b8b-c377f27387e3/reports/a3d3cf67-10eb-474a-a7ed-fa9c1da03230/ReportSectionfd506659f9d2800e7012?ctid=128848de-2354-4443-84fa-e54a418f8c65&experience=power-bi

Waterfall:
https://verticalbridge-my.sharepoint.com/:x:/r/personal/jordan_tuck_verticalbridge_com/_layouts/15/Doc.aspx?sourcedoc=%7BFABCF3F3-24A3-4A92-AB4D-4900694E3AEF%7D&file=Pre-Screening%20Waterfall%20List%204.24.2024%201.xlsx&action=default&mobileredirect=true&DefaultItemOpen=1&ct=1722281733297&wdOrigin=OFFICECOM-WEB.START.REC&cid=de66c401-3f12-4847-9f97-b9962640bc2c&wdPreviousSessionSrc=HarmonyWeb&wdPreviousSession=686d458d-5a5c-47d7-a2c5-ed25e805e549

'''

#   ---- FUNCTIONS TO RUN SCRIPT ON ANY MACHINE ----


# Calculates the age at which the user's documents folder was created
def accountAge(filepath):
    return (time.time() - os.stat(filepath)[stat.ST_CTIME])/2592000

''' This function allows any user to run the script on their own machine
    The function will identify the system username and then identify the Vertical Bridge username by determining the age of the account
    An account from before the cloud migration follows a different naming pattern - first initial then last name
'''
# Determines the VB username
def determineAge(filepath):
    cutoffDateDays = pd.to_datetime('08-01-2021') - pd.to_datetime(time.time())
    months = int(str(cutoffDateDays).split(' ')[0]) / 30
    accountAgeMonths = float(accountAge(filepath))

    user = osUser
    names = re.findall('[A-Z][^A-Z]*', user)
    i = 1
    if accountAgeMonths >= months:
        user = names[0][0]
        while i < len(names):
            user += names[i]
            i += 1
    else:
        user = names[0]
        while i < len(names):
            user += '.' + names[i]
            i += 1

    user = user.lower()
    return user

# If new lists are not downloaded, this one is and can be used as an example
testDate = '20240913'

osUser = os.getlogin()
directoryPath = r'/Users/' + osUser + '/Downloads/'
user = determineAge(directoryPath)



#   ---- GET DATA FROM SOURCES - FLOYD, SITETRACKER, POWER BI REPORTS, WATERFALL ----


# Dates referenced:
currDate = date.today().strftime('%Y%m%d')
siteTrackerDate = date.today().strftime('%m_%d_%Y') #add Time
# Date of 18 months ago - Leases beginning soon etc. reference this date
date = (date.today() - dateutil.relativedelta.relativedelta(months = 18))


# Floyd Data (Site List and Tenant Leases)
''' Floyd data is downloaded as {Site or Tenant}_{VB username}_{date}.xlsx
    The VB username was determined by the accountAge and determineAge functions above.
    This searches for the files corresponding to the current date in your system downloads folder.
    If no current files are found, it will use default data.
'''
siteFiles = [filename for filename in os.listdir(directoryPath) if filename.startswith('Site_' + user + '_' + currDate) and filename.endswith('.xlsx')]
siteFilePath = ('Site_' + user + '_' + testDate + '.xlsx') if len(siteFiles) == 0 else siteFiles[0]
fullSiteList = pd.read_excel(directoryPath + siteFilePath, engine = 'calamine')

tenantFiles = [filename for filename in os.listdir(directoryPath) if filename.startswith('Tenant_' + user + '_' + currDate) and filename.endswith('.xlsx')]
tenantFilePath = ('Tenant_' + user + '_' + testDate + '.xlsx') if len(tenantFiles) == 0 else tenantFiles[0]
tenantLeases = pd.read_excel(directoryPath + tenantFilePath, engine = 'calamine')
# ex: r'/Users/TiaraMathur/Downloads/Tenant_tiara.mathur_20240805.xlsx'


# SiteTracker Colocation Leasing Data
''' SiteTracker files download with the time in them.
    Instead of trying to download at a specific time, down to the exact millisecond - 
    This searches in your downloads folder for a downloaded Colocation Leasing data file beginning with the current date and takes that file.
    If no current files are found, it will use default data.
'''
files = [filename for filename in os.listdir(directoryPath) if filename.startswith('Leasing Project' + date.today().strftime('%#m_%#d_%Y'))]
filePath = 'Leasing Project9_11_2024, 9_24_16 AM.csv' if len(files) == 0 else files[0]
allColoLeasing = pd.read_csv(directoryPath + filePath, low_memory = False)


# Waterfall List Data
waterfallList = pd.read_csv(directoryPath + 'Pre-Screening Waterfall List 4.24 (1).csv', header = 0, encoding = 'unicode_escape', low_memory = False)


# SiteTracker Decomm Projects Data
''' SiteTracker files download with the time in them.
    Instead of trying to download at a specific time, down to the exact millisecond - 
    This searches in your downloads folder for a downloaded Decomm Project data file beginning with the current date and takes that file.
    If no current files are found, it will use default data.
'''
files = [filename for filename in os.listdir(directoryPath) if filename.startswith('Project' + date.today().strftime('%#m_%#d_%Y'))]
# Default data
filePath = 'Project9_16_2024, 2_14_46 PM.csv' if len(files) == 0 else files[0]
decommList = pd.read_csv(directoryPath + filePath, low_memory = False)

decommList['Created Date'] = pd.to_datetime(decommList['Created Date'], format = 'mixed')
# For sites with multiple decom projects - select the one with the most recent "created date" and remove duplicates
decommList = decommList.sort_values(by = 'Created Date', ascending = False).drop_duplicates(subset = 'Site Number')


# Power BI Financial Data
'''The Power BI report is downloaded as data.xlsx
    There is no date attached - so this will select the latest one downloaded (data(1), data(2), etc..)
    Make sure there are no other (or at least newer) files in the downloads folder named data + __ + .xlsx
'''
# Find data.xlsx files
files = [filename for filename in os.listdir(directoryPath) if filename.startswith('data') and filename.endswith('xlsx')]
# Extract the actual file name to order - with the extension, data.xlsx will come after data (1).xlsx
# This is the only case that would cause error, but important to catch all potential issues
fileName = [x.split('.')[0] for x in files]
# Properly sort
fileName.sort()

# Reattach the extension
dataFile = fileName[len(fileName) - 1] + '.xlsx'

#Suppressed warning for this because openpyxl feels the need to produce a warning saying it will use its formatting defaults
with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    financialInfo = pd.read_excel(directoryPath + dataFile, engine = 'calamine')



#   ---- GET NAKED SITE LIST AND FLOYD INFO ----


NASiteList = fullSiteList.loc[fullSiteList['Status Of Progress'].isna()]
BuiltSiteList = fullSiteList.loc[fullSiteList['Status Of Progress'] == 'Built']
allSiteList = pd.concat([NASiteList, BuiltSiteList], axis=0)

# allSiteList = Built + NA
allSiteList = allSiteList.loc[allSiteList['SiteClusterName'] == allSiteList['Site No']]

# Active and Inactive leases
# Checked, there are no others
activeTenantLeases = tenantLeases.loc[tenantLeases['Tenant Lease Is Active'] == 'Yes']
inactiveTenantLeases = tenantLeases.loc[tenantLeases['Tenant Lease Is Active'] == 'No']

# Leases that have a commencement date within the last 18 months
soonLeases = tenantLeases.loc[tenantLeases['Tenant Lease Is Active'] == 'No']
soonLeases = soonLeases.loc[soonLeases['Tenant Termination Date'].isnull()]
soonLeases = soonLeases.loc[pd.to_datetime(soonLeases['Tenant Commencement Date']) > pd.to_datetime(date)]

# Sites that have a start date within the last 18 months
newSites = allSiteList.loc[pd.to_datetime(allSiteList['Date Start']) > pd.to_datetime(date)]
newSitesHavingLease = newSites[newSites['Site No'].isin(tenantLeases['Site No'])]

# Sites that had tower stacked notice sent out within the last 18 months
othVR = allSiteList.loc[pd.to_datetime(allSiteList['OTHVRDate']) > pd.to_datetime(date)]

# Non-marketable sites
unmarketable = allSiteList.loc[(allSiteList['Display On Web']) != 'Yes']

# Exclude sites with active leases, sites with new leases, and non-marketable sites from the Naked Site List
nakedSiteList = allSiteList[~allSiteList['Site No'].isin(activeTenantLeases['Site No'])]
nakedSiteList = nakedSiteList[~nakedSiteList['Site No'].isin(soonLeases['Site No'])]
nakedSiteList = nakedSiteList[~nakedSiteList['Site No'].isin(newSitesHavingLease['Site No'])]
nakedSiteList = nakedSiteList[~nakedSiteList['Site No'].isin(unmarketable['Site No'])]
nakedSiteList = nakedSiteList[~nakedSiteList['Site No'].isin(othVR['Site No'])]



#   ---- ADD NAKED INFO AND DAYS/MONTHS/YEARS ----


# Separate sites into those with terminated leases and those that have always been naked
terminated = nakedSiteList[nakedSiteList['Site No'].isin(tenantLeases['Site No'])]
alwaysNaked = nakedSiteList[~nakedSiteList['Site No'].isin(tenantLeases['Site No'])]

# Currently assuming all sites in "terminated" are actually terminated. But some of these are new leases that just have not executed yet
terminatedTL = tenantLeases[tenantLeases['Site No'].isin(terminated['Site No'])]
terminatedNotReallyTL = terminatedTL.loc[terminatedTL['Tenant Termination Date'].isna()]

# Remove new leases from our naked site list and our terminated sites/leases lists
nakedSiteList = nakedSiteList[~nakedSiteList['Site No'].isin(terminatedNotReallyTL['Site No'])]
terminated = terminated[~terminated['Site No'].isin(terminatedNotReallyTL['Site No'])]
terminatedTL = terminatedTL[~terminatedTL['Site No'].isin(terminatedNotReallyTL['Site No'])]

# Find most recently terminated leases
terminatedTL['Tenant Termination Date'] = pd.to_datetime(terminatedTL['Tenant Termination Date'])
# Find the index of the most recent date for each Site No
mostRecentLeaseIndex = terminatedTL.groupby('Site No')['Tenant Termination Date'].idxmax()
# Get sites with the most recent dates
mostRecentLeases = terminatedTL.loc[mostRecentLeaseIndex]

# Distinguish between newly naked and always naked sites in the Naked Status column
terminated['Naked Status'] = 'Naked'
alwaysNaked['Naked Status'] = 'Always Naked'

# Add most recent lease termination info to the terminated sites list
terminatedWithDate = pd.merge(terminated, terminatedTL[['Site No', 'Tenant Termination Date']], on = 'Site No', how = 'left')

# The naked date of terminated sites will be the most recent lease's tenant termination date
terminatedWithDate['Naked Date'] = terminatedWithDate['Tenant Termination Date']
# The naked date of always naked sites will be the start date (or tower stacked date if start is blank.. or 1/1/2014 as a default value if no start info is found)
alwaysNaked['Naked Date'] = alwaysNaked['Date Start']
alwaysNaked.loc[alwaysNaked['Date Start'].isna(), 'Naked Date'] = alwaysNaked['OTHVRDate']
alwaysNaked.loc[alwaysNaked['Naked Date'].isna(), 'Naked Date'] = '1/1/2014'

# Combine the terminated and always naked sites and correctly format naked date
nakedSiteList = pd.concat([terminatedWithDate, alwaysNaked], ignore_index = True, axis = 0)
nakedSiteList['Tenant Termination Date'] = nakedSiteList['Tenant Termination Date'].dt.strftime('%m/%d/%Y')
nakedSiteList['Naked Date'] = pd.to_datetime(nakedSiteList['Naked Date'])
nakedSiteList['Naked Date'] = nakedSiteList['Naked Date'].dt.strftime('%m/%d/%Y')

# Calculate naked days, months, and years
nakedSiteList['Naked Days'] = (pd.to_datetime(currDate) - pd.to_datetime(nakedSiteList['Naked Date'])).dt.days
nakedSiteList['Naked Months'] = (nakedSiteList['Naked Days'] / 30).astype(int)
nakedSiteList['Naked Years'] = (nakedSiteList['Naked Days'] / 365).astype(int)



#   ---- ADD DEALS IN PROGRESS INFORMATION ----


# Remove deals that are dead or fully executed
allColoLeasing = allColoLeasing.drop(allColoLeasing[allColoLeasing['Deal Status'] == 'Fully Executed'].index)
allColoLeasing = allColoLeasing.drop(allColoLeasing[allColoLeasing['Deal Status'] == 'Dead Deal'].index)

# Lambda function to calculate number of deals in-progress and if AT&T, Verizon, etc have deals in progress
dealAggregation = allColoLeasing.groupby('Site Number').apply(lambda x: pd.Series({
    'Deals In-Progress': len(x),
    'Verizon In-Progress Deal': 'Verizon Wireless' in x['Reporting Relationship'].values,
    'T-Mobile In-Progress Deal': any(carrier in x['Reporting Relationship'].values for carrier in ['T-Mobile', 'Sprint']),
    'AT&T In-Progress Deal': 'AT&T' in x['Reporting Relationship'].values,
    'Dish In-Progress Deal': 'Dish' in x['Reporting Relationship'].values,
    'Other In-Progress Deal': 'Other' in x['Reporting Relationship'].values
}), include_groups = False).reset_index()

# Rename column for merging
dealAggregation = dealAggregation.rename(columns = {'Site Number': 'Site No'})
# Combine deals-in-progress information with naked site list for naked sites 
nakedSiteList = pd.merge(nakedSiteList, dealAggregation[['Site No', 'Deals In-Progress', 'AT&T In-Progress Deal', 'Dish In-Progress Deal', 'T-Mobile In-Progress Deal', 'Verizon In-Progress Deal', 'Other In-Progress Deal']], on = 'Site No', how = 'left')

# To avoid warnings with fillna
with pd.option_context("future.no_silent_downcasting", True):
    # For naked sites that did not have any in-progress deals, put 0
    nakedSiteList['Deals In-Progress'] = nakedSiteList['Deals In-Progress'].fillna(0).infer_objects(copy=False)
    # For naked sites that did not have any in-progress deals, put false for deals in progress by each carrier 
    nakedSiteList['AT&T In-Progress Deal'] = nakedSiteList['AT&T In-Progress Deal'].fillna(False).infer_objects(copy=False)
    nakedSiteList['T-Mobile In-Progress Deal'] = nakedSiteList['T-Mobile In-Progress Deal'].fillna(False).infer_objects(copy=False)
    nakedSiteList['Verizon In-Progress Deal'] = nakedSiteList['Verizon In-Progress Deal'].fillna(False).infer_objects(copy=False)
    nakedSiteList['Dish In-Progress Deal'] = nakedSiteList['Dish In-Progress Deal'].fillna(False).infer_objects(copy=False)
    nakedSiteList['Other In-Progress Deal'] = nakedSiteList['Other In-Progress Deal'].fillna(False).infer_objects(copy=False)

# Reformat deals to be integer values (no need to say 1.0 deals for ex)
nakedSiteList['Deals In-Progress'] = nakedSiteList['Deals In-Progress'].astype(int)



#   ---- ADD WATERFALL INFORMATION ----


# Combine relevant columns from waterfall
nakedSiteList = pd.merge(nakedSiteList, waterfallList[['Site No', 'AT&T TLUP', 'AT&T Overall Need', 'TMO TLUP', 'TMO Overall Need', 'VZW TLUP', 'VZW Overall Need']], on = 'Site No', how = 'left')

# Lambda function to calculate number of carriers with need
nakedSiteList['Number of Carriers with Need'] = nakedSiteList.apply(
    lambda x: sum(any(need in str(x[columnName]) for need in ['High Need', 'New Coverage']) 
                  for columnName in ['AT&T Overall Need', 'VZW Overall Need', 'TMO Overall Need']), axis = 1)


#   ---- ADD DECOMM INFORMATION ----


# Rename columns to clarify they refer to decomm
decommList = decommList.rename(columns={'Site Number': 'Site No', 'Project Status': 'Decom Project Status', 'Legacy Decom Status': 'Decom Status', 'TTD Complete Property Restored (A)': 'Decom Date'})
# Merge decomm columns with naked site list
nakedSiteList = pd.merge(nakedSiteList, decommList[['Site No', 'Decom Status', 'Decom Project Status', 'Decom Date']], on = 'Site No', how = 'left')

# Obtain necessary columns from financial information
financialInfo = financialInfo[['Tower Number', 'TCF Status', 'Gross Monthly Rent', 'Site Monthly Percentage Rev Share', 'Monthly Fixed Rev Share', 'Net Monthly Rent', 'Monthly CAM', 'Net Monthly Rent Incl. CAM', 'Monthly Ground Rent', 'Monitoring Expense', 'Insurance Expense', 'Property Tax Expense', 'Utilities Expense', 'Maintenance Expense', 'Monthly Site Operating Expenses', 'Monthly Site TCF', 'Monthly Site TCF Incl. CAM']]
# Rename Site Number column
financialEdited = financialInfo.rename(columns = {'Tower Number': 'Site No'})
# Merge financial data columns with naked site list
nakedSiteList = pd.merge(nakedSiteList, financialEdited, on = 'Site No', how = 'left')

nakedSiteList = nakedSiteList.drop_duplicates(['Site No'])
nakedSiteList = nakedSiteList.drop(nakedSiteList[nakedSiteList['Date Start'].isna()].index)
nakedSiteListCleaned = nakedSiteList.drop(['Site Name', 'Address Line 2', 'Ground Elevation (feet)', 'Owner Site ID', 'Tower Owner', 'FCC Registration Number', 'Legal', 'Leasing Project Manager', 'Real Estate POC', 'Fiber', 'Status Of Progress', 'FAA Height', 'LockBox Address', 'FAA Ht_AGL', 'FAA Study No', 'BTA Name', 'MTA Name', 'MSA/RSA Name', 'BEA Name', 'Mortgage Recorded', 'Display On Web', 'Drone Inspection Date', 'SIR Inspection Date', 'OTHVRDate', 'ProjectType', 'SiteClusterName', 'Naked Days'], axis = 1)



#   ---- ADD HRR INFORMATION ----


# Creating the template needed for GIS Next with the appropriate formatting and setting the radius to 1/2 mile
gisNextTemplate = nakedSiteList[['Site No', 'Site Name', 'Latitude', 'Longitude']]
gisNextTemplate = gisNextTemplate.rename(columns = {'Site No': 'SiteNumber', 'Site Name': 'SiteName'})
gisNextTemplate['Radius']= 0.50

# Saving the template with updated naked site info for GIS Next

'''
Unfortunately this needs to be saved in this way - with column names separately saved, headers completely removed, and then columns reinserted one at a time in the header.
There is no better way to remove weird bold xlsx header formatting with pandas, and GIS Next will give an error.
'''

# Save the columns
columns = gisNextTemplate.columns
# Output to excel
writer = pd.ExcelWriter(directoryPath + 'gisNextTemplate' + currDate + '.xlsx', engine = 'xlsxwriter')
# Write with no header
gisNextTemplate.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)
# Get workbook and worksheet
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add the column names back in
for i, val in enumerate(columns):
    worksheet.write(0, i, val)

writer.close()

# Add the results of CMA Analysis to the Naked Site List
nakedSiteListFull = nakedSiteListCleaned

'''
The CMA Result report is downloaded as CMAResult.xlsx
There is no date attached - so this will select the latest one downloaded (CMAResult(1), CMAResult(2), etc..)
Make sure there are no other (or at least newer) files in the downloads folder named CMAResult + __ + .xlsx
'''

# Find CMAResult.xlsx files
files = [filename for filename in os.listdir(directoryPath) if filename.startswith('CMAResult') and filename.endswith('xlsx')]
# Extract the actual file name to order - with the extension, CMAResult.xlsx will come after CMAResult (1).xlsx
# This is the only case that would cause error, but important to catch all potential issues
fileName = [x.split('.')[0] for x in files]
# Properly sort
fileName.sort()

# Reattach the extension
dataFile = fileName[len(fileName) - 1] + '.xlsx'

gisNextCMA = pd.read_excel(directoryPath + dataFile, 'Raw Data Report', engine = 'calamine')
gisNextCMA = gisNextCMA.drop('Site No', axis = 1)
gisNextCMA = gisNextCMA.rename(columns = {'Site Search ID': 'Site No'})

# There are multiple HRR/HRP Databases
# Combine Verizon 2016, Verizon 2023 etc by taking the first "word"
gisNextCMA['Site Owner'] = gisNextCMA['Site Owner'].apply(lambda x: x.split(' ')[0])

# Filter and find the row with the smallest 'Distance' for each 'Site No' and 'Site Owner'
minDistance = gisNextCMA.loc[gisNextCMA.groupby(['Site No', 'Site Owner'])['Distance'].idxmin()]
# Combine by site number
pivotMerged = minDistance.pivot_table(index = 'Site No', columns = 'Site Owner', values = 'Site Name', aggfunc = 'first')

# Rename columns to match 'ATT HRR', 'T-Mobile HRR', 'Verizon HRR'
pivotMerged.columns = [f'{owner} HRR under 1/2 mi' for owner in pivotMerged.columns]

# Merge the naked site list with these new columns
nakedSiteListFull = pd.merge(nakedSiteListFull, pivotMerged[['ATT HRR under 1/2 mi', 'T-Mobile HRR under 1/2 mi', 'Verizon HRR under 1/2 mi']], on = 'Site No', how = 'left')



#   ---- SAVE AND EXPORT NAKED SITE LIST ----

# Sort by site number ascending
nakedSiteListFull = nakedSiteListFull.sort_values('Site No')

# Path to save the naked site list 
vbPath = r'/Users/' + osUser + '/OneDrive - Vertical Bridge/'
# Write and save
nakedSiteListFull.to_excel(vbPath + 'Naked Site Report ' + date.today().strftime('%m-%d-%y') + '.xlsx', index = False)

print('Naked Site List generated and saved to ' + vbPath + 'Naked Site Report ' + date.today().strftime('%m-%d-%y') + '.xlsx')