from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup as bsoup
import calendar as cd
import datetime as dt
import requests as rq
import smtplib
import xlwt

password = input("Please enter password: ")

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
dateName = dt.datetime.today().strftime("%m-%d-%Y")
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
# Set up arrays for sites and data
soup = bsoup(r.content, "lxml") 
sites = []
siteName = ''
data = [
    ['Location', 'Month Sales', 'Projected', 'Goal', 'Percent', 'Month Cars', 'Projected', 'Goal', 'Percent']
]

salesTA = 0
salesPA = 0
salesGA = 0
carsTA = 0
carsPA = 0
carsGA = 0
num = 1

divSites = soup.find_all("div", {"data-trayid": "site"})
for div in divSites:
    liSites = div.find_all("li")
    for li in liSites:
        sites.append(li.get_text())

# Find Sales and Car Count by Site
for site in sites:

    # Site Name Lookup
    divNames = soup.find_all("div", {"data-trayid": "site"})
    for div in divNames:
        liNames = div.find_all("li", attrs={"data-site": site})
        for li in liNames:
            siteName = str(' '.join(li.find('a', {'rel': True}).get('rel')))

    # Monthly Sales Total
    divSales = soup.find_all("div", {"data-trayid": "totalsales"})
    for div in divSales:
        liSales = div.find_all("li", attrs={"data-site": site})
        for li in liSales:
            # Convert to int for calculation
            x = int(li.get_text().replace('$', '').replace(',',''))
            y = ( x / today ) * daysIM
            z = y # Goal for month from Mike's spreadhseet
            if y != 0 and z != 0:
                p = ( y / z ) * 100
            else:
                p = 0
            salesT = round(x)
            salesP = round(y)
            salesG = round(z)
            salesGP = "{:,}".format(round(p))
            salesTA = salesTA + x
            salesPA = salesPA + y
            salesGA = salesGA + z
            data.append([])
            data[num].append(siteName)
            data[num].append(salesT) 
            data[num].append(salesP)
            data[num].append(salesG)
            data[num].append(salesGP)

    # Monthly Cars Total
    divCount = soup.find_all("div", {"data-trayid": "total"})

    for div in divCount:
        liCount = div.find_all("li", attrs={"data-site": site})
        for li in liCount:
            # Convert to int for calculation
            x = int(li.get_text().replace(',', ''))
            y = ( x / today ) * daysIM
            z = y # Goal for month from Mike's spreadhseet
            if y != 0 and z != 0:
                p = ( y / z ) * 100
            else:
                p = 0
            carsT = round(x)
            carsP = round(y)
            carsG = round(z)
            carsGP = "{:,}".format(round(p))
            carsTA = carsTA + x
            carsPA = carsPA + y
            carsGA = carsGA + z
            data[num].append(carsT)
            data[num].append(carsP) 
            data[num].append(carsG) 
            data[num].append(carsGP) 
            num = num + 1

# Goal Percentages for all stores
if salesPA != 0 and salesGA != 0:
    salesGPA = ( salesPA / salesGA ) * 100
else:
    salesGPA = 0

if carsPA != 0 and carsGA != 0:
    carsGPA = ( carsPA / carsGA ) * 100
else:
    carsGPA = 0

# String formatting for AllStore variables
salesTA = "${:,}".format(round(salesTA))
salesPA = "${:,}".format(round(salesPA))
salesGA = "${:,}".format(round(salesGA))
salesGPA = "{:,}".format(round(salesGPA))
carsTA = "{:,}".format(round(carsTA))
carsPA = "{:,}".format(round(carsPA))
carsGA = "{:,}".format(round(carsGA))
carsGPA = "{:,}".format(round(carsGPA))

# Save Data to XLS
book = xlwt.Workbook()
sheet = book.add_sheet("Data")
filename = "Projections " + dateName + ".xls"

sheet.col(0).width = 256 * 42

for i, l in enumerate(data):
    for j, col in enumerate(l):
        sheet.write(i, j, col)

book.save(filename)

# Send Email
def send_email():
    myEmail = 'anthony@clearskycapitalinc.com'
    myPassword = input('Email password: ')
    tim = 'tim@clearskycapitalinc.com'

    # Email details
    msgTo = myEmail
    msgFrom = myEmail
    msg = MIMEMultipart()
    bodyText = """Hello,

Here are the projections and goals for all locations as of {0}. Please see attached spreadhseet for a breakdown of these totals. 

"""
    bodyTable = """

Sales Total:    {1}
Projection:     {2}
Goal:               {3}
Goal Projected: {4}%

Car Count Total: {5}
Projection:     {6}
Goal:               {7}
Goal Projected: {8}%

"""
    body = MIMEText(bodyText.format(date, salesTA, salesPA, salesGA, salesGPA, carsTA, carsPA, carsGA, carsGPA))
    msg.attach(body)

    # XLS attachment
    fp = open(filename, 'rb')
    att = MIMEApplication(fp.read(),_subtype="xls")
    fp.close()
    att.add_header('Content-Disposition','attachment',filename=filename)
    msg.attach(att)

    # Send email
    msg['Subject'] = "Projections: " + date
    msg['To'] = msgTo
    msg['From'] = msgFrom

    mail = smtplib.SMTP(host='smtp.office365.com', port=587)
    mail.starttls()
    mail.login(myEmail, myPassword)
    mail.sendmail(msgFrom, msgTo, msg.as_string())
    mail.quit()

if __name__=='__main__':
    send_email()
