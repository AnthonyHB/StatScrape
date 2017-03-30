"""
Clear Sky Capital - Car Wash Projections via StatWatch
by AnthonyHB

Input: Dynamic data from StatWatch
Process: ( Monthly Sales (and Car count) / Day of Month ) * Days in Month
Output: Projected Montly Sales & Show Projected Monthly Car Count
"""

from bs4 import BeautifulSoup as bsoup
import calendar as cd
import datetime as dt
import requests as rq



password = input("Please enter the password:")

# Login Form for StatWatch
payload = {
    'username': 'ant',
    'password': password,
    'chain': 'scttaz'
}

# Today's date
now = dt.datetime.now()
mon = now.month
days = now.day 
daysIM = cd.monthrange(now.year, now.month)[1]

# Monthly Sales Session - Use 'with' to ensure the session context is closed after use.
with rq.Session() as s:
    p = s.post('https://www.statwatch.com/login', data=payload)

    # Authorized Request
    pt1 = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=total&start="
    pt2 = "/2017&granularity=6&todate=1&difftype=percentup&unittype=abs"
    url = pt1 + str(mon) + "/" + str(days) + pt2
    r = s.get(url)

# Parse HTML data 
soup = bsoup(r.content, "lxml")

# Turn all sites into an array
sites = []
divSites = soup.find_all("div", {"data-trayid": "site"})
for div in divSites:
    liSites = div.find_all("li")
    for li in liSites:
        sites.append(li.get_text())

for site in sites:

    # Monthly Sales Total
    divSales = soup.find_all("div", {"data-trayid": "totalsales"})
    for div in divSales:
        liSales = div.find_all("li", attrs={"data-site": site})
        for li in liSales:
            # Remove $ and comma from monSales, and convert to int
            monSales = int(li.get_text().replace('$', '').replace(',',''))
            projSales = ( monSales / days ) * daysIM
            print(site + " Sales to Date: ${:,}".format(round(monSales))) 
            print(site + " Projected Sales: ${:,}".format(round(projSales)))

    # Monthly Cars Total
    divCount = soup.find_all("div", {"data-trayid": "total"})
    for div in divCount:
        liCount = div.find_all("li", attrs={"data-site": site})
        for li in liCount:
            # Remove comma from monCars, and convert to int
            monCars = int(li.get_text().replace(',', ''))
            projCars = ( monCars / days ) * daysIM
            print(site + " Car Count to Date: {:,}".format(monCars))
            print(site + " Projected Car Count: {:,}".format(round(projCars))) 
        
