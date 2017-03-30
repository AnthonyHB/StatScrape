"""
Clear Sky Capital - Car Wash Projections via StatWatch
by AnthonyHB

Input: Dynamic data from StatWatch
Process: ( Monthly Sales (and Car count) / Day of Month ) * Days in Month
Output: Projected Montly Sales & Show Projected Monthly Car Count
"""

from bs4 import BeautifulSoup as bsoup
import calendar as cd
import csv
import datetime as dt
import requests as rq

password = input("Please enter password:")

# Login Form for StatWatch
payload = {
    'username': 'ant',
    'password': password,
    'chain': 'scttaz'
}

# Today's date
now = dt.datetime.now()
today = now.day 
date = dt.datetime.today().strftime("%m/%d/%Y")
daysIM = cd.monthrange(now.year, now.month)[1]


# SiteWatch Session - Use 'with' to ensure the session context is closed after use.
with rq.Session() as s:
    p = s.post('https://www.statwatch.com/login', data=payload)

    # Authorized Request
    pt1 = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=total&start="
    pt2 = "&granularity=6&todate=1&difftype=percentup&unittype=abs"
    url = pt1 + date + pt2
    r = s.get(url)
    
# HTML data 
# Pass all sites into one array
soup = bsoup(r.content, "lxml") 
sites = []

divSites = soup.find_all("div", {"data-trayid": "site"})
for div in divSites:
    liSites = div.find_all("li")
    for li in liSites:
        sites.append(li.get_text())

# Find Sales and Car Count by Site
data = [
    ['Site', 'Monthly Sales', 'Projected Sales', 'Monthly Cars', 'Projected Cars']
]

num = 1
for site in sites:
    # Monthly Sales Total
    divSales = soup.find_all("div", {"data-trayid": "totalsales"})
    for div in divSales:
        liSales = div.find_all("li", attrs={"data-site": site})
        for li in liSales:
            salesMonth = int(li.get_text().replace('$', '').replace(',','')) # Convert to int
            salesProj = ( salesMonth / today ) * daysIM
            data.append([])
            data[num].append(site)
            data[num].append("${:,}".format(round(salesMonth))) 
            data[num].append("${:,}".format(round(salesProj)))

    # Monthly Cars Total
    divCount = soup.find_all("div", {"data-trayid": "total"})
    for div in divCount:
        liCount = div.find_all("li", attrs={"data-site": site})
        for li in liCount:
            carsMonth = int(li.get_text().replace(',', '')) # Convert to int
            carsProj = ( carsMonth / today ) * daysIM
            data[num].append("{:,}".format(carsMonth))
            data[num].append("{:,}".format(round(carsProj)))
            num = num + 1 

# Save Data to CSV
with open("output.csv", 'w') as resultFile:
    wr = csv.writer(resultFile, dialect='excel')
    wr.writerows(data)