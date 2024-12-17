

def determineNakedSites(fullSiteList, tenantLeases):
    import dateutil
    import pandas as pd

    date = (date.today() - dateutil.relativedelta.relativedelta(months = 18))

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

    return nakedSiteList


def addNakedTime(nakedSiteList, tenantLeases):
    from datetime import date
    import pandas as pd

    currDate = date.today().strftime('%Y%m%d')

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

    return nakedSiteList


def getNakedSiteListUpdates(currNakedSiteList, nakedSiteList):
    import pandas as pd

    additions = nakedSiteList[~nakedSiteList['Site No'].isin(currNakedSiteList['Site No'])]
    additions = additions.assign(Change = 'New naked site')
    removals = currNakedSiteList[~currNakedSiteList['Site No'].isin(nakedSiteList['Site No'])]
    removals = removals.assign(Change = 'No longer naked')

    nakedSiteListChanges = []
    nakedSiteListChanges = pd.concat([additions, removals], axis = 0)
    nakedSiteListChanges = nakedSiteListChanges[['Site No', 'Change']]
    return nakedSiteListChanges