"""
Clear Sky Capital - Car Wash Projections via StatWatch
by AnthonyHB

Input: Dynamic data from StatWatch
Process: ( Monthly Sales (and Car count) / Day of Month ) * Days in Month
Output: Projected Montly Sales & Show Projected Monthly Car Count
"""

from bs4 import BeautifulSoup as bsoup
from email.mime.text import MIMEText
import calendar as cd
import csv
import datetime as dt
import requests as rq
import xlwt
import smtplib

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
data = [
    ['Location', 'Month Sales', 'Projected', 'Month Cars', 'Projected']
]
sites = []
siteName = ''
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
            salesM = int(li.get_text().replace('$', '').replace(',','')) # Convert to int
            salesP = ( salesM / today ) * daysIM
            data.append([])
            data[num].append(siteName)
            data[num].append("${:,}".format(round(salesM))) 
            data[num].append("${:,}".format(round(salesP)))

    # Monthly Cars Total
    divCount = soup.find_all("div", {"data-trayid": "total"})

    for div in divCount:
        liCount = div.find_all("li", attrs={"data-site": site})
        for li in liCount:
            carsM = int(li.get_text().replace(',', '')) # Convert to int
            carsP = ( carsM / today ) * daysIM
            data[num].append("{:,}".format(carsM))
            data[num].append("{:,}".format(round(carsP))) 
            num = num + 1

# Save Data to XLS
book = xlwt.Workbook()
sheet = book.add_sheet("Data")

sheet.col(0).width = 256 * 42

for i, l in enumerate(data):
    for j, col in enumerate(l):
        sheet.write(i, j, col)

book.save(str("Projections " + dateName + ".xls"))

"""
# Under Construction / If Needed
# Save Data to CSV (If needed)
with open("output.csv", 'w') as resultFile:
    wr = csv.writer(resultFile, dialect='excel')
    wr.writerows(data)

# Send Email
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 443
SMTP_USERNAME = emailusername
SMTP_PASSWORD = emailpassword

EMAIL_TO = [emailusername]
EMAIL_FROM = emailusername
EMAIL_SUBJECT = "Demo Email : "

DATE_FORMAT = "%d/%m/%Y"
EMAIL_SPACE = ", "

DATA='This is the content of the email.'

def send_email():
    msg = MIMEText(DATA)
    msg['Subject'] = EMAIL_SUBJECT + " %s" % (date)
    msg['To'] = EMAIL_SPACE.join(EMAIL_TO)
    msg['From'] = EMAIL_FROM
    mail = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    mail.starttls()
    mail.login(SMTP_USERNAME, SMTP_PASSWORD)
    mail.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
    mail.quit()

if __name__=='__main__':
    send_email()
"""
