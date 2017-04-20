from bs4 import BeautifulSoup as bsoup
import calendar as cd
import datetime as dt
import requests as rq
import xlwt

history = []
titles = [
            'Date',
            'Location',
            'Month Sales', 
            'Month Cars'
        ]
allDates = [titles]

for x in range(2013, 2018):
    year = x
    for x in range(12):
        month = x + 1
        day = cd.monthrange(year, month)[1]
        if month < 10:
            month = "0" + str(month)
        date = str(month) + "/" + str(day) + "/" + str(year)
        history.append(date)

for date in history:
    # Real function
    payload = {
        'username': 'ant',
        'password': 'lp14',
        'chain': 'scttaz'
    }

    # SiteWatch Session - Use 'with' to ensure the session context is closed after use.
    with rq.Session() as s:
        p = s.post('https://www.statwatch.com/login', data=payload)

        # Authorized Request
        pt1 = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=total&start="
        pt2 = "&granularity=6&todate=1&difftype=percentup&unittype=abs"
        url = pt1 + date + pt2
        r = s.get(url)

    # HTML scrape
    soup = bsoup(r.content, "lxml") 

    sites = {}

    divSites = soup.find_all("div", {"data-trayid": "site"})
    for div in divSites:
        liSites = div.find_all("li")
        for li in liSites:
            siteID = li.get_text()
            sites.update({siteID: {'name': siteID, 'date': date}})

    # Find Sales and Car Count by Site
    for site in sites:

        # Site Location Lookup
        divLocations = soup.find_all("div", {"data-trayid": "site"})
        for div in divLocations:
            liLocations = div.find_all("li", attrs={"data-site": site})
            for li in liLocations:
                location = str(' '.join(li.find('a', {'rel': True}).get('rel')))
                sites[site].update({'location': location})

        # Monthly Sales Lookup
        divSales = soup.find_all("div", {"data-trayid": "totalsales"})
        for div in divSales:
            liSales = div.find_all("li", attrs={"data-site": site})
            for li in liSales:
                # Convert to int for calculation
                x = li.get_text()
                if x != '-' and x != '$-':
                    x = int(li.get_text().replace('$', '').replace(',', ''))
                else:
                    x = 0
                sites[site].update({
                                        'sales': x
                                    })

        # Monthly Cars Lookup
        divCount = soup.find_all("div", {"data-trayid": "total"})
        for div in divCount:
            liCount = div.find_all("li", attrs={"data-site": site})
            for li in liCount:
                # Convert to int for calculation
                x = li.get_text()
                if x != '-' and x != '$-':
                    x = int(li.get_text().replace(',', ''))
                else:
                    x = 0
                sites[site].update({
                                        'cars': x
                                    })

    # separate into regions and companies
    for site in sites:

        dataSheet = [
                        sites[site]['date'],
                        sites[site]['location'],
                        sites[site]['sales'],
                        sites[site]['cars']
                    ]

        allDates.append(dataSheet)

# Save Data to XLS
book = xlwt.Workbook()
sheetAllDates = book.add_sheet("All Dates")

filename = "Historical.xls"

sheetAllDates.col(1).width = 256 * 42

for i, l in enumerate(allDates):
    for j, col in enumerate(l):
        sheetAllDates.write(i, j, col)

book.save(filename)
