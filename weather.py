from bs4 import BeautifulSoup as bsoup
from sitedata import data
import calendar as cd
import datetime as dt
import requests as rq
import xlwt
import simplejson, urllib, requests, webbrowser

history = []
titles = [
            'Date',
            'Location',
            'Precipitation',
            'Open Days'
        ]
allDates = [titles]

for x in range(2013, 2018):
    year = x
    for x in range(12):
        month = x + 1
        day = cd.monthrange(year, month)[1]
        month = str(month)
        date = str(month) + "/" + str(day) + "/" + str(year)
        history.append([month, year])

num = 0

payload = {
            'username': 'ant',
            'password': 'lp14',
            'chain': 'scttaz'
        }

for date in history:

    month = str(history[num][0])
    year = str(history[num][1])
    date = str(month) + "/" + str(year)

    for site in data:
        # SiteWatch Session - Use 'with' to ensure the session context is closed after use.
        with rq.Session() as s:
            p = s.post('https://www.statwatch.com/login', data=payload)
            # Authorized Request
            url = "https://www.statwatch.com/api/2.3/org/SCTTAZ/site/{}/pc/1/almanac/month-summary?year={}&month={}&cb=0.9416995570087567".format(site, year, month)
            r = s.get(url)

        json = r.json() 

        data[site].update({'WeatherPrecip': json['MonthStats']['WeatherPrecip']})
        data[site].update({'OpenDays': json['MonthStats']['OpenDays']})
        data[site].update({'date': date})

        dataSheet = [

                        data[site]['date'],
                        site, # location ID
                        data[site]['WeatherPrecip'],
                        data[site]['OpenDays']
                    ]

        allDates.append(dataSheet)
        print(dataSheet)

    num = num + 1

# Save Data to XLS
book = xlwt.Workbook()
sheetAllDates = book.add_sheet("All Dates")
filename = "Weather.xls"

for i, l in enumerate(allDates):
    for j, col in enumerate(l):
        sheetAllDates.write(i, j, col)

book.save(filename)