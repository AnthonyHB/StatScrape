from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup as bsoup
import calendar as cd
import datetime as dt
import requests as rq
import smtplib
import xlwt

# Today's date
now = dt.datetime.now()
today = now.day 
date = dt.datetime.today().strftime("%m/%d/%Y")
dateName = dt.datetime.today().strftime("%m-%d-%Y")
daysIM = cd.monthrange(now.year, now.month)[1]

# Login Form for StatWatch
password = input('password: ')

payload = {
    'username': 'ant',
    'password': password,
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

titles = [
        'Location', 
        'Month Sales', 
        'Projected', 
        'Goal', 
        'Percent', 
        'Month Cars', 
        'Projected', 
        'Goal', 
        'Percent',
        'Region'
    ]

allSites = [titles]
lp14 = [titles]
lp21 = [titles]
regionAZ = [titles]
regionCA = [titles]
regionNV = [titles]

data = {
    '01A': {
                'salesGoal': 79355, 
                'carsGoal': 12776, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '01C': {
                'salesGoal': 57564, 
                'carsGoal': 9292, 
                'region': 'CA',
                'company': 'LP14'
            },
    '01N': {
                'salesGoal': 69573, 
                'carsGoal': 10818, 
                'region': 'NV',
                'company': 'LP14'
            },
    '02A': {
                'salesGoal': 93575, 
                'carsGoal': 16203, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '02C': {
                'salesGoal': 121795,
                'carsGoal': 14142, 
                'region': 'CA',
                'company': 'LP14'
            },
    '02N': {
                'salesGoal': 84369, 
                'carsGoal': 8553, 
                'region': 'NV',
                'company': 'LP14'
            },
    '03A': {
                'salesGoal': 112866,
                'carsGoal': 20714, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '03C': {
                'salesGoal': 153398,
                'carsGoal': 7896, 
                'region': 'CA',
                'company': 'LP21'
            },
    '03N': {
                'salesGoal': 66684, 
                'carsGoal': 10530, 
                'region': 'NV',
                'company': 'LP14'
            },
    '04A': {
                'salesGoal': 116386,
                'carsGoal': 20035, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '04C': {
                'salesGoal': 32604, 
                'carsGoal': 4903, 
                'region': 'CA',
                'company': 'LP21'
            },
    '04N': {
                'salesGoal': 32616, 
                'carsGoal': 6155, 
                'region': 'NV',
                'company': 'LP14'
            },
    '05A': {
                'salesGoal': 114318,
                'carsGoal': 7063, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '05C': {
                'salesGoal': 75416, 
                'carsGoal': 11257, 
                'region': 'CA',
                'company': 'LP21'
            },
    '05N': {
                'salesGoal': 26094, 
                'carsGoal': 5799, 
                'region': 'NV',
                'company': 'LP14'
            },
    '06A': {
                'salesGoal': 114517,
                'carsGoal': 20935, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '06C': {
                'salesGoal': 62598, 
                'carsGoal': 9791, 
                'region': 'CA',
                'company': 'LP21'
            },
    '07A': {
                'salesGoal': 83399, 
                'carsGoal': 14372, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '07N': {
                'salesGoal': 35989, 
                'carsGoal': 4484, 
                'region': 'NV',
                'company': 'LP14'
            },
    '08A': {
                'salesGoal': 60670, 
                'carsGoal': 15297, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '09A': {
                'salesGoal': 97203, 
                'carsGoal': 16925, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '10A': {
                'salesGoal': 54332, 
                'carsGoal': 9654, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '11A': {
                'salesGoal': 30007, 
                'carsGoal': 4763, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '12A': {
                'salesGoal': 93056, 
                'carsGoal': 19192, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '14A': {
                'salesGoal': 53854, 
                'carsGoal': 7906, 
                'region': 'AZ',
                'company': 'LP14'
            },
    '15A': {
                'salesGoal': 52429, 
                'carsGoal': 5110, 
                'region': 'AZ',
                'company': 'LP14'
            },
     '16A': {
                 'salesGoal': 0, 
                 'carsGoal': 0, 
                 'region': 'AZ',
                 'company': 'LP14'
             }
}

salesTotalAll = salesProjectAll = salesGoalAll = carsTotalAll = carsProjectAll = carsGoalAll = 0
lp14SalesTotal = lp14SalesProject = lp14SalesGoal = lp14SalesGoalPercent = lp14CarsTotal = lp14CarsProject = lp14CarsGoal = lp14CarsGoalPercent = 0
lp21SalesTotal = lp21SalesProject = lp21SalesGoal = lp21SalesGoalPercent = lp21CarsTotal = lp21CarsProject = lp21CarsGoal = lp21CarsGoalPercent = 0

sites = []
siteLocation = ''

divSites = soup.find_all("div", {"data-trayid": "site"})
for div in divSites:
    liSites = div.find_all("li")
    for li in liSites:
        sites.append(li.get_text())

# Find Sales and Car Count by Site
for site in sites:

    if site in data:
        counter = 0
    else:
        data.update({site: {'salesGoal': 0, 'carsGoal': 0, 'region': '','company': ''}})

    # Site Location Lookup
    divLocations = soup.find_all("div", {"data-trayid": "site"})
    for div in divLocations:
        liLocations = div.find_all("li", attrs={"data-site": site})
        for li in liLocations:
            siteLocation = str(' '.join(li.find('a', {'rel': True}).get('rel')))

    # Monthly Sales Lookup
    divSales = soup.find_all("div", {"data-trayid": "totalsales"})
    for div in divSales:
        liSales = div.find_all("li", attrs={"data-site": site})
        for li in liSales:
            # Convert to int for calculation
            x = int(li.get_text().replace('$', '').replace(',',''))
            y = ( x / today ) * daysIM
            z = data[site]['salesGoal']
            if y != 0 and z != 0:
                p = ( y / z ) * 100
            else:
                p = 0
            salesTotal = round(x)
            salesProject = round(y)
            salesGoal = round(z)
            salesGoalPercent= "{:,}".format(round(p))
            salesTotalAll = salesTotalAll + x
            salesProjectAll = salesProjectAll + y
            salesGoalAll = salesGoalAll + z
            data[site].update({
                                    'siteLocation': siteLocation,
                                    'salesTotal': salesTotal,
                                    'salesProject': salesProject,
                                    'salesGoalPercent': salesGoalPercent
                                })

    # Monthly Cars Lookup
    divCount = soup.find_all("div", {"data-trayid": "total"})
    for div in divCount:
        liCount = div.find_all("li", attrs={"data-site": site})
        for li in liCount:
            # Convert to int for calculation
            x = int(li.get_text().replace(',', ''))
            y = ( x / today ) * daysIM
            z = data[site]['carsGoal']
            if y != 0 and z != 0:
                p = ( y / z ) * 100
            else:
                p = 0
            carsTotal = round(x)
            carsProject = round(y)
            carsGoal = round(z)
            carsGoalPercent = "{:,}".format(round(p))
            carsTotalAll = carsTotalAll + x
            carsProjectAll = carsProjectAll + y
            carsGoalAll = carsGoalAll + z
            data[site].update({
                                    'carsTotal': carsTotal,
                                    'carsProject': carsProject,
                                    'carsGoalPercent': carsGoalPercent
                                })

# separate into regions and companies
for site in sites:

    dataSheet = [
                    data[site]['siteLocation'],
                    data[site]['salesTotal'],
                    data[site]['salesProject'],
                    data[site]['salesGoal'],
                    data[site]['salesGoalPercent'],
                    data[site]['carsTotal'],
                    data[site]['carsProject'],
                    data[site]['carsGoal'],
                    data[site]['carsGoalPercent'],
                    data[site]['region']
                ]

    allSites.append(dataSheet)

    if data[site]['region'] == 'AZ':
        regionAZ.append(dataSheet)

    elif data[site]['region'] == 'CA':
        regionCA.append(dataSheet)

    elif data[site]['region'] == 'NV':
        regionNV.append(dataSheet)
    else:
        print(str(data[site]) + " wasn't added to a region.")

    if data[site]['company'] == 'LP14':
        lp14.append(dataSheet)

        lp14SalesTotal = lp14SalesTotal + data[site]['salesTotal']
        lp14SalesProject = lp14SalesProject + data[site]['salesProject']
        lp14SalesGoal = lp14SalesGoal + data[site]['salesGoal']
        lp14CarsTotal = lp14CarsTotal + data[site]['carsTotal']
        lp14CarsProject = lp14CarsProject + data[site]['carsProject']
        lp14CarsGoal = lp14CarsGoal + data[site]['carsGoal']

    elif data[site]['company'] == 'LP21':
        lp21.append(dataSheet)

        lp21SalesTotal = lp21SalesTotal + data[site]['salesTotal']
        lp21SalesProject = lp21SalesProject + data[site]['salesProject']
        lp21SalesGoal = lp21SalesGoal + data[site]['salesGoal']
        lp21CarsTotal = lp21CarsTotal + data[site]['carsTotal']
        lp21CarsProject = lp21CarsProject + data[site]['carsProject']
        lp21CarsGoal = lp21CarsGoal + data[site]['carsGoal']

    else:
        print(str(data[site]) + " wasn't added to a company.")

# Goal Percentages for all stores
if salesProjectAll != 0 and salesGoalAll != 0:
    salesGoalPercentAll = ( salesProjectAll / salesGoalAll ) * 100
else:
    salesGoalPercentAll = 0

if carsProjectAll != 0 and carsGoalAll != 0:
    carsGoalPercentAll = ( carsProjectAll / carsGoalAll ) * 100
else:
    carsGoalPercentAll = 0

if lp14SalesProject != 0 and lp14SalesGoal != 0:
    lp14SalesGoalPercent = ( lp14SalesProject / lp14SalesGoal ) * 100
else:
    lp14SalesGoalPercent = 0

if lp21SalesProject != 0 and lp21SalesGoal != 0:
    lp21SalesGoalPercent = ( lp21SalesProject / lp21SalesGoal ) * 100
else:
    lp21SalesGoalPercent = 0

if lp14CarsProject != 0 and lp14CarsGoal != 0:
    lp14CarsGoalPercent = ( lp14CarsProject / lp14CarsGoal ) * 100
else:
    lp14CarsGoalPercent = 0

if lp21CarsProject != 0 and lp21CarsGoal != 0:
    lp21CarsGoalPercent = ( lp21CarsProject / lp21CarsGoal ) * 100
else:
    lp21CarsGoalPercent = 0

# String formatting for All Store variables
salesTotalAll = "${:,}".format(round(salesTotalAll))
salesProjectAll = "${:,}".format(round(salesProjectAll))
salesGoalAll = "${:,}".format(round(salesGoalAll))
salesGoalPercentAll = "{:,}".format(round(salesGoalPercentAll))

carsTotalAll = "{:,}".format(round(carsTotalAll))
carsProjectAll = "{:,}".format(round(carsProjectAll))
carsGoalAll = "{:,}".format(round(carsGoalAll))
carsGoalPercentAll = "{:,}".format(round(carsGoalPercentAll))

# String formatting for company variables
lp14SalesTotal = "${:,}".format(round(lp14SalesTotal))
lp14SalesProject = "${:,}".format(round(lp14SalesProject))
lp14SalesGoal = "${:,}".format(round(lp14SalesGoal))
lp14SalesGoalPercent = "{:,}".format(round(lp14SalesGoalPercent))

lp14CarsTotal = "{:,}".format(round(lp14CarsTotal))
lp14CarsProject = "{:,}".format(round(lp14CarsProject))
lp14CarsGoal = "{:,}".format(round(lp14CarsGoal))
lp14CarsGoalPercent = "{:,}".format(round(lp14CarsGoalPercent))

lp21SalesTotal = "${:,}".format(round(lp21SalesTotal))
lp21SalesProject = "${:,}".format(round(lp21SalesProject))
lp21SalesGoal = "${:,}".format(round(lp21SalesGoal))
lp21SalesGoalPercent = "{:,}".format(round(lp21SalesGoalPercent))

lp21CarsTotal = "{:,}".format(round(lp21CarsTotal))
lp21CarsProject = "{:,}".format(round(lp21CarsProject))
lp21CarsGoal = "{:,}".format(round(lp21CarsGoal))
lp21CarsGoalPercent = "{:,}".format(round(lp21CarsGoalPercent))


# Save Data to XLS
book = xlwt.Workbook()
sheetAllSites = book.add_sheet("All Sites")
sheetLP14 = book.add_sheet("LP14")
sheetLP21 = book.add_sheet("LP21")
sheetAZ = book.add_sheet("AZ")
sheetCA = book.add_sheet("CA")
sheetNV = book.add_sheet("NV")
filename = "Projections " + dateName + ".xls"

sheetAllSites.col(0).width = sheetLP14.col(0).width = sheetLP21.col(0).width = sheetAZ.col(0).width = sheetCA.col(0).width = sheetNV.col(0).width = 256 * 42

for i, l in enumerate(allSites):
    for j, col in enumerate(l):
        sheetAllSites.write(i, j, col)

for i, l in enumerate(lp14):
    for j, col in enumerate(l):
        sheetLP14.write(i, j, col)

for i, l in enumerate(lp21):
    for j, col in enumerate(l):
        sheetLP21.write(i, j, col)

for i, l in enumerate(regionAZ):
    for j, col in enumerate(l):
        sheetAZ.write(i, j, col)

for i, l in enumerate(regionCA):
    for j, col in enumerate(l):
        sheetCA.write(i, j, col)

for i, l in enumerate(regionNV):
    for j, col in enumerate(l):
        sheetNV.write(i, j, col)

book.save(filename)

def send_email():
    # Email details
    text = """Hello,
    
Here are the projections and goals for all locations as of {0}. Please see attached spreadhseet for a breakdown of these totals. 

"""
    html = """\
<html>
    <head></head>
    <body>
        <table style="border-collapse: collapse; width: 50%">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black"></th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Sales Total</th>
                <td style="border: 1px solid black" align="center">{0}</td>
                <td style="border: 1px solid black" align="center">{1}</td>
                <td style="border: 1px solid black" align="center">{2}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{3}</td>
                <td style="border: 1px solid black" align="center">{4}</td>
                <td style="border: 1px solid black" align="center">{5}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{6}</td>
                <td style="border: 1px solid black" align="center">{7}</td>
                <td style="border: 1px solid black" align="center">{8}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{9}%</td>
                <td style="border: 1px solid black" align="center">{10}%</td>
                <td style="border: 1px solid black" align="center">{11}%</td>
            </tr>
        </table>
        <p></p>
        <table style="border-collapse: collapse; width: 50%">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black"></th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Car Count Total</th>
                <td style="border: 1px solid black" align="center">{12}</td>
                <td style="border: 1px solid black" align="center">{13}</td>
                <td style="border: 1px solid black" align="center">{14}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{15}</td>
                <td style="border: 1px solid black" align="center">{16}</td>
                <td style="border: 1px solid black" align="center">{17}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{18}</td>
                <td style="border: 1px solid black" align="center">{19}</td>
                <td style="border: 1px solid black" align="center">{20}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{21}%</td>
                <td style="border: 1px solid black" align="center">{22}%</td>
                <td style="border: 1px solid black" align="center">{23}%</td>
            </tr>
        </table>
    </body>
</html>
    """
    myEmail = 'anthony@clearskycapitalinc.com'
    myGmail = 'anthony.benites17@gmail.com'
    myPassword = input("password: ")

    tBarrett = 'tim@clearskycapitalinc.com'
    mWong = 'michael@clearskycapitalinc.com'
    aShell = 'andrew@clearskycapitalinc.com'
    mGrimes = 'mike@quickncleanaz.com'
    recipients = [myEmail, myGmail]
    
    msg = MIMEMultipart()

    text = text.format(date)
    html = html.format(salesTotalAll, lp14SalesTotal, lp21SalesTotal,  salesProjectAll, lp14SalesProject, lp21SalesProject,  salesGoalAll,  lp14SalesGoal, lp21SalesGoal, salesGoalPercentAll,  lp14SalesGoalPercent, lp21SalesGoalPercent, carsTotalAll,  lp14CarsTotal, lp21CarsTotal, carsProjectAll, lp14CarsProject, lp21CarsProject, carsGoalAll, lp14CarsGoal, lp21CarsGoal, carsGoalPercentAll, lp14CarsGoalPercent, lp21CarsGoalPercent)

    text = MIMEText(text)
    html = MIMEText(html, 'html')
    msg.attach(text)
    msg.attach(html)    

    # XLS attachment
    fp = open(filename, 'rb')
    att = MIMEApplication(fp.read(),_subtype="xls")
    fp.close()
    att.add_header('Content-Disposition','attachment',filename=filename)
    msg.attach(att)

    # Send email
    msg['Subject'] = "Projections: " + date
    msg['To'] = ", ".join(recipients)
    msg['From'] = myEmail

    mail = smtplib.SMTP(host='smtp.office365.com', port=587)
    mail.starttls()
    mail.login(myEmail, myPassword)
    mail.sendmail(myEmail, recipients, msg.as_string())
    mail.quit()

if __name__=='__main__':    
    send_email()