"""

Input Q2 information

Fix projections for Raceway

Make Region Email for Mike

Make Region Email for Managers

"""

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup as bsoup
import calendar as cd
import datetime as dt
import requests as rq
from sitedata import data
from htmlBody import cskyHTML
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

soup = bsoup(r.content, "lxml") 

class CarWash(object):
    def __init__(self, name):
        self.name = name

    def scrape_data(self):
        divAddresses = soup.find_all("div", {"data-trayid": "site"})
        for div in divAddresses:
            liAddresses = div.find_all("li", attrs={"data-site": self.name})
            for li in liAddresses:
                self.address = str(' '.join(li.find('a', {'rel': True}).get('rel')))
                self.region = data[self.name]['region'],
                self.company = data[self.name]['company']

        divSales = soup.find_all("div", {"data-trayid": "totalsales"})
        for div in divSales:
            liSales = div.find_all("li", attrs={"data-site": self.name})
            for li in liSales:
                # Convert to int for calculation
                x = int(li.get_text().replace('$', '').replace(',',''))
                y = ( x / today ) * daysIM
                z = data[self.name]['salesGoal']
                if y != 0 and z != 0:
                    p = ( y / z ) * 100
                else:
                    p = 0
                self.salesTotal = round(x)
                self.salesProject = round(y)
                self.salesGoal = round(z)
                self.salesPercent= "{:,}".format(round(p))

        divCount = soup.find_all("div", {"data-trayid": "total"})
        for div in divCount:
            liCount = div.find_all("li", attrs={"data-site": self.name})
            for li in liCount:
                # Convert to int for calculation
                x = int(li.get_text().replace(',', ''))
                y = ( x / today ) * daysIM
                z = data[self.name]['carsGoal']
                if y != 0 and z != 0:
                    p = ( y / z ) * 100
                else:
                    p = 0
                self.carsTotal = round(x)
                self.carsProject = round(y)
                self.carsGoal = round(z)
                self.carsPercent = "{:,}".format(round(p))

class Summary(object):
    def __init__(self, name):
        self.name = name,
        self.carwashes = {}
        self.sheet = [
                        [
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
                     ]

    def addSite(self, site, row):
        self.carwashes.update({site: sites[site]})
        self.sheet.append(row)

    def getTotals(self):
        self.salesTotal = 0
        self.salesProject = 0
        self.salesGoal = 0
        self.salesPercent = 0

        self.carsTotal = 0
        self.carsProject = 0
        self.carsGoal = 0
        self.carsPercent = 0

        for site in self.carwashes:
            self.salesTotal = self.salesTotal + self.carwashes[site].salesTotal
            self.salesProject = self.salesProject + self.carwashes[site].salesProject
            self.salesGoal = self.salesGoal + self.carwashes[site].salesGoal

            self.carsTotal = self.carsTotal + self.carwashes[site].carsTotal
            self.carsProject = self.carsProject + self.carwashes[site].carsProject
            self.carsGoal = self.carsGoal + self.carwashes[site].carsGoal

        if self.salesProject != 0 and self.salesGoal != 0:
            self.salesPercent = ( self.salesProject / self.salesGoal ) * 100

        if self.carsProject != 0 and self.carsGoal != 0:
            self.carsPercent = ( self.carsProject / self.carsGoal ) * 100

        self.salesTotal = "${:,}".format(round(self.salesTotal))
        self.salesProject = "${:,}".format(round(self.salesProject))
        self.salesGoal = "${:,}".format(round(self.salesGoal))
        self.salesPercent = "{:,}".format(round(self.salesPercent))

        self.carsTotal = "{:,}".format(round(self.carsTotal))
        self.carsProject = "{:,}".format(round(self.carsProject))
        self.carsGoal = "{:,}".format(round(self.carsGoal))
        self.carsPercent = "{:,}".format(round(self.carsPercent))

def start_scraping(groups):
    divSites = soup.find_all("div", {"data-trayid": "site"})
    for div in divSites:
        liSites = div.find_all("li")
        for li in liSites:
            siteID = li.get_text()
            sites.update({siteID: CarWash(siteID)})
            sites[siteID].scrape_data()

    for group in groups:
        summaries.update({group: Summary(group)})

    for site in sites:
        dataRow = [
                        sites[site].address,
                        sites[site].salesTotal,
                        sites[site].salesProject,
                        sites[site].salesGoal,
                        sites[site].salesPercent,
                        sites[site].carsTotal,
                        sites[site].carsProject,
                        sites[site].carsGoal,
                        sites[site].carsPercent,
                        sites[site].region
                    ]

        for summary in summaries:
            if summary == data[site]['region'] or summary == data[site]['company'] or summary == 'All':
                summaries[summary].addSite(site, dataRow)

    for summary in summaries:
        summaries[summary].getTotals()

def get_filename():
    filename = "Projections " + dateName + ".xls"
    return filename

def create_workbook(filename):
    # Save Data to XLS
    book = xlwt.Workbook()
    sheets = {}

    for data in dataSheets:
        sheets.update({data: book.add_sheet(data)})
        sheets[data].col(0).width = 256 * 42

        for i, l in enumerate(summaries[data].sheet):
            for j, col in enumerate(l):
                sheets[data].write(i, j, col)

    book.save(filename)

def send_email(filename):
    # Email details
    text = """Hello,
    
Here are the projections and goals for all locations as of {0}. Please see attached spreadhseet for a breakdown of these totals. 

"""
    html = cskyHTML
    myEmail = 'anthony@clearskycapitalinc.com'
    myGmail = 'anthony.benites17@gmail.com'
    myPassword = input("password: ")

    tBarrett = 'tim@clearskycapitalinc.com'
    mWong = 'michael@clearskycapitalinc.com'
    aShell = 'andrew@clearskycapitalinc.com'
    mGrimes = 'mike@quickncleanaz.com'
    mCollins = 'matthew@clearskycapitalinc.com'
    recipients = [myEmail, myGmail]
    
    msg = MIMEMultipart()

    text = text.format(date)
    html = html.format(
                        summaries['All'].salesTotal, 
                        summaries['LP14'].salesTotal, 
                        summaries['LP21'].salesTotal,  
                        summaries['All'].salesProject, 
                        summaries['LP14'].salesProject, 
                        summaries['LP21'].salesProject,  
                        summaries['All'].salesGoal,  
                        summaries['LP14'].salesGoal, 
                        summaries['LP21'].salesGoal, 
                        summaries['All'].salesPercent,  
                        summaries['LP14'].salesPercent, 
                        summaries['LP21'].salesPercent, 
                        summaries['All'].carsTotal,  
                        summaries['LP14'].carsTotal, 
                        summaries['LP21'].carsTotal, 
                        summaries['All'].carsProject, 
                        summaries['LP14'].carsProject, 
                        summaries['LP21'].carsProject, 
                        summaries['All'].carsGoal, 
                        summaries['LP14'].carsGoal, 
                        summaries['LP21'].carsGoal, 
                        summaries['All'].carsPercent, 
                        summaries['LP14'].carsPercent, 
                        summaries['LP21'].carsPercent
                    )

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

"""
This is the actual running program.
"""

sites = {}
summaries = {}
dataSheets = ['All', 'LP14', 'LP21', 'AZ', 'CA', 'NV']

if __name__=='__main__': 
    start_scraping(dataSheets)
    create_workbook(get_filename())   
    send_email(get_filename())
