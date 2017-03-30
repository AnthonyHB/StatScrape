"""
Clear Sky Capital - Car Wash Projections via StatWatch
by AnthonyHB

Input: Dynamic data from StatWatch
Process: ( Monthly Sales (or Car count) / Day of Month ) * Days in Month
Output: Projected Montly Sales & Show Projected Monthly Car Count
"""

from bs4 import BeautifulSoup as bsoup
import requests as rq

password = input("Please enter the password:")

# Login Form for StatWatch
payload = {
    'username': 'ant',
    'password': password,
    'chain': 'scttaz'
}

# Today's date
mon = 3
days = 29
daysIM = 31

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
divTag = soup.find_all("div", {"data-trayid": "totalsales"})

# Monthly Sales totals
for tag in divTag:
    liTags = tag.find_all("li", attrs={"data-site": "01A"})
    for tag in liTags:
        # Remove commas from monSales, and convert to int
        num = tag.get_text().replace(',', '')
        monSales = int(num.replace('$',''))
        salesProj = ( monSales / days ) * daysIM
        print("01A Projected Sales: ${:,}".format(round(salesProj)) + ". Monthly Sales to Date: ${:,}".format(round(monSales)) + ".") 

divTag = soup.find_all("div", {"data-trayid": "total"})

# Monthly Cars totals
for tag in divTag:
    liTags = tag.find_all("li", attrs={"data-site": "01A"})
    for tag in liTags:
        # Remove commas, and convert to int
        monCars = int(tag.get_text().replace(',', ''))
        carsProj = ( monCars / days ) * daysIM
        print("01A Projected Car Count: {:,}".format(round(carsProj)) + ". Monthly Car Count to Date: {:,}".format(monCars) + ".") 
        
