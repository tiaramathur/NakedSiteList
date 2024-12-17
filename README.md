# Naked Site List
This script creates a full report of "Naked Sites", towers with no active tenants. Naked sites should be leased out or decommissioned to avoid unnecessary operating expenses.

This report is fully automated and aggregates:
- A full list of naked site names and geographical information
- The duration that these sites have been naked, whether since the termination of a previous lease or since the tower was built/acquired
- Potential deals in progress relating to these sites, for major carriers
- Likelihood of leasing the site based on internal waterfall model
- Current decommission projects relating to each site
- Financial data and operating expenses, to provide a better understanding of costs that will be saved by taking action to lease or remove each tower
- High rent responsibility sites nearby, for potential colocation opportunities

## PyPi Package

The package [nakedSiteList-package](https://pypi.org/project/nakedSiteList-package/) defines functions that aggregate a naked site list, calculate naked duration, and contrast a naked site list report with a previous report to track updates - newly naked towers and towers that are no longer active and "naked".

Installation:

````python
pip install nakedSiteList-package
````
