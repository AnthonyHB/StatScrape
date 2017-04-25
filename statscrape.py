"""
Make Region Email for Mike and Managers
"""
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from bs4 import BeautifulSoup as bsoup
from sitedata import qDict
from emailhtml import cskyHTML
import calendar as cd
import datetime as dt
import requests as rq
import pandas as pd
import smtplib
import io

class CarWash(object):
    def __init__(self, name):
        self.name = name
        self.region = info.ix[self.name]['Region']
        self.company = info.ix[self.name]['Company']
        self.address = info.ix[self.name]['Location']

    def get_total(self, x):
        if x != '-' and x != '$-':
            total = float(x.replace('$', '').replace(',', '').replace(' %',''))
        else:
            total = 0
        return total

    def get_project(self, total, day, daysIM):
        project = ( total / day ) * daysIM
        return project

    def get_percent(self, project, goal):
        if project != 0 and goal != 0:
            percent = project / goal
        else:
            percent = 0
        return percent

    def scrape_init(self):
        # Pull M1
        with rq.Session() as s:
            p = s.post('https://www.statwatch.com/login', data=payload)
            url = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=total&start={}&granularity=6&todate=1&difftype=percentup&unittype=abs".format(date)
            r = s.get(url)
        soup = bsoup(r.content, "lxml") 

        # Sales
        divSales = soup.find_all("div", {"data-trayid": "totalsales"})
        for div in divSales:
            liSales = div.find_all("li", attrs={"data-site": self.name})
            for li in liSales:
                # Get text from li, and use it to calculate totals
                self.salesTotal = self.get_total(li.get_text())
                self.salesProject = self.get_project(self.salesTotal, today, daysIM)
                self.salesGoal = goals.ix[self.name].ix['Sales'][month]
                self.salesPercent= self.get_percent(self.salesProject, self.salesGoal)

        # Car Count
        divCount = soup.find_all("div", {"data-trayid": "total"})
        for div in divCount:
            liCount = div.find_all("li", attrs={"data-site": self.name})
            for li in liCount:
                self.carsTotal = self.get_total(li.get_text())
                self.carsProject = self.get_project(self.carsTotal, today, daysIM)
                self.carsGoal = goals.ix[self.name].ix['Cars'][month]
                self.carsPercent = self.get_percent(self.carsProject, self.carsGoal)

        # Labor 
        divLabor = soup.find_all("div", {"data-trayid": "laborpercent"})
        for div in divLabor:
            liLabor = div.find_all("li", attrs={"data-site": self.name})
            for li in liLabor:
                # Returns a percentage of revenue, calculates to a dollar amount.
                self.laborTotal = ( self.get_total(li.get_text()) / 100 ) * self.salesTotal
                self.laborProject = self.get_project(self.laborTotal, today, daysIM)
                self.laborGoal = goals.ix[self.name].ix['Labor'][month]
                self.laborPercent = self.get_percent(self.laborProject, self.laborGoal)

        # Pull M2 
        with rq.Session() as s:
            p = s.post('https://www.statwatch.com/login', data=payload)
            newDate = str(month2) + "/" + str(cd.monthrange(now.year, month2)[1]) + "/" + str(now.year)
            url = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=totalsales&start={}&granularity=1&todate=1&difftype=percentup&unittype=abs".format(newDate)
            r = s.get(url)
        soup = bsoup(r.content, "lxml") 

        # Sales
        divSales = soup.find_all("div", {"data-trayid": "totalsales"})
        for div in divSales:
            liSales = div.find_all("li", attrs={"data-site": self.name})
            for li in liSales:
                self.salesTotalM2 = self.get_total(li.get_text())
                self.salesGoalM2 = goals.ix[self.name].ix['Sales'][str(month2)]

        # Car count
        divCount = soup.find_all("div", {"data-trayid": "total"})
        for div in divCount:
            liCount = div.find_all("li", attrs={"data-site": self.name})
            for li in liCount:
                self.carsTotalM2 = self.get_total(li.get_text())
                self.carsGoalM2 = goals.ix[self.name].ix['Cars'][str(month2)]

        # Labor 
        divLabor = soup.find_all("div", {"data-trayid": "laborpercent"})
        for div in divLabor:
            liLabor = div.find_all("li", attrs={"data-site": self.name})
            for li in liLabor:
                self.laborTotalM2 = ( self.get_total(li.get_text()) / 100 ) * self.salesTotalM2
                self.laborGoalM2 = goals.ix[self.name].ix['Labor'][str(month2)]

        # Pull M3 
        with rq.Session() as s:
            p = s.post('https://www.statwatch.com/login', data=payload)
            newDate = str(month3) + "/" + str(cd.monthrange(now.year, month3)[1]) + "/" + str(now.year)
            url = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=totalsales&start={}&granularity=1&todate=1&difftype=percentup&unittype=abs".format(newDate)
            r = s.get(url)
        soup = bsoup(r.content, "lxml") 

        # Sales
        divSales = soup.find_all("div", {"data-trayid": "totalsales"})
        for div in divSales:
            liSales = div.find_all("li", attrs={"data-site": self.name})
            for li in liSales:
                self.salesTotalM3 = self.get_total(li.get_text())
                self.salesGoalM3 = goals.ix[self.name].ix['Sales'][str(month3)]

        # Car count
        divCount = soup.find_all("div", {"data-trayid": "total"})
        for div in divCount:
            liCount = div.find_all("li", attrs={"data-site": self.name})
            for li in liCount:
                self.carsTotalM3 = self.get_total(li.get_text())
                self.carsGoalM3 = goals.ix[self.name].ix['Cars'][str(month3)]

        # Labor
        divLabor = soup.find_all("div", {"data-trayid": "laborpercent"})
        for div in divLabor:
            liLabor = div.find_all("li", attrs={"data-site": self.name})
            for li in liLabor:
                self.laborTotalM3 = ( self.get_total(li.get_text()) / 100 ) * self.salesTotalM3
                self.laborGoalM3 = goals.ix[self.name].ix['Labor'][str(month3)]

        self.salesTotalQ = self.salesTotal + self.salesTotalM2 + self.salesTotalM3
        self.salesProjectQ = ( ( self.salesTotalQ ) / daysPIQ ) * daysIQ
        self.salesGoalQ = self.salesGoal + self.salesGoalM2 + self.salesGoalM3

        self.carsTotalQ = self.carsTotal + self.carsTotalM2 + self.carsTotalM3
        self.carsProjectQ = ( ( self.carsTotalQ ) / daysPIQ ) * daysIQ
        self.carsGoalQ = self.carsGoal + self.carsGoalM2 + self.carsGoalM3

        self.laborTotalQ = self.laborTotal + self.laborTotalM2 + self.laborTotalM3
        self.laborProjectQ = ( ( self.laborTotalQ ) / daysPIQ ) * daysIQ
        self.laborGoalQ = self.laborGoal + self.laborGoalM2 + self.laborGoalM3

    def dilute_sales(self, day):
        # M1
        with rq.Session() as s:
            p = s.post('https://www.statwatch.com/login', data=payload)
            newDate = str(now.month) + "/" + str(day) + "/" + str(now.year)
            url = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=totalsales&start={}&granularity=1&todate=1&difftype=percentup&unittype=abs".format(newDate)
            r = s.get(url)
        soup = bsoup(r.content, "lxml") 

        divSales = soup.find_all("div", {"data-trayid": "totalsales"})
        for div in divSales:
            liSales = div.find_all("li", attrs={"data-site": self.name})
            for li in liSales:
                x = self.get_total(li.get_text())
                salesDilute = ( x / daysIM ) * today
                salesTotalNew = self.salesTotal - x + salesDilute
                salesTotalNewQ = salesTotalNew + self.salesTotalM2 + self.salesTotalM3
                self.salesProject = self.get_project(salesTotalNew, today, daysIM)
                self.salesPercent = self.get_percent(self.salesProject, self.salesGoal)
                self.salesProjectQ = ( salesTotalNewQ / daysPIQ ) * daysIQ

class Summary(object):
    def __init__(self, name):
        self.name = name,
        self.carwashes = {}
        self.titles = [
                        'Location', 
                        'Sales MTD', 
                        'Projected', 
                        'Plan', 
                        'Percent', 
                        'Cars MTD', 
                        'Projected', 
                        'Plan', 
                        'Percent',
                        'Labor MTD',
                        'Projected',
                        'Plan',
                        'Percent'
                    ]
        self.indexRow = []
        self.sheet = []

    def add_site(self, site, row):
        self.carwashes.update({site: sites[site]})
        self.indexRow.append(site)
        self.sheet.append(row)

    def int_to_string(self):
        self.salesTotal = "${:,}".format(round(int(self.salesTotal)))
        self.salesProject = "${:,}".format(round(int(self.salesProject)))
        self.salesGoal = "${:,}".format(round(int(self.salesGoal)))
        self.salesPercent = "{:,}".format(round(int(self.salesPercent)))

        self.carsTotal = "{:,}".format(round(int(self.carsTotal)))
        self.carsProject = "{:,}".format(round(int(self.carsProject)))
        self.carsGoal = "{:,}".format(round(int(self.carsGoal)))
        self.carsPercent = "{:,}".format(round(int(self.carsPercent)))

        self.laborTotal = "{:,}".format(round(int(self.laborTotal)))
        self.laborProject = "{:,}".format(round(int(self.laborProject)))
        self.laborGoal = "{:,}".format(round(int(self.laborGoal)))
        self.laborPercent = "{:,}".format(round(int(self.laborPercent)))

        self.salesTotalQ = "${:,}".format(round(int(self.salesTotalQ)))
        self.salesProjectQ = "${:,}".format(round(int(self.salesProjectQ)))
        self.salesGoalQ = "${:,}".format(round(int(self.salesGoalQ)))
        self.salesPercentQ = "{:,}".format(round(int(self.salesPercentQ)))

        self.carsTotalQ = "{:,}".format(round(int(self.carsTotalQ)))
        self.carsProjectQ = "{:,}".format(round(int(self.carsProjectQ)))
        self.carsGoalQ = "{:,}".format(round(int(self.carsGoalQ)))
        self.carsPercentQ = "{:,}".format(round(int(self.carsPercentQ)))

        self.laborTotalQ = "{:,}".format(round(int(self.laborTotalQ)))
        self.laborProjectQ = "{:,}".format(round(int(self.laborProjectQ)))
        self.laborGoalQ = "{:,}".format(round(int(self.laborGoalQ)))
        self.laborPercentQ = "{:,}".format(round(int(self.laborPercentQ)))

    def get_percent(self, project, goal):
        if project != 0 and goal != 0:
            percent = ( project / goal ) * 100
        else:
            percent = 0
        return percent

    def get_totals(self):
        self.salesTotal = self.salesProject = self.salesGoal = self.salesPercent = 0
        self.carsTotal = self.carsProject = self.carsGoal = self.carsPercent = 0
        self.laborTotal = self.laborProject = self.laborGoal = self.laborPercent = 0
        self.salesTotalQ = self.salesProjectQ = self.salesGoalQ = self.salesPercentQ = 0
        self.carsTotalQ = self.carsProjectQ = self.carsGoalQ = self.carsPercentQ = 0
        self.laborTotalQ = self.laborProjectQ = self.laborGoalQ = self.laborPercentQ = 0

        for site in self.carwashes:
            self.salesTotal = self.salesTotal + self.carwashes[site].salesTotal
            self.salesProject = self.salesProject + self.carwashes[site].salesProject
            self.salesGoal = self.salesGoal + self.carwashes[site].salesGoal

            self.carsTotal = self.carsTotal + self.carwashes[site].carsTotal
            self.carsProject = self.carsProject + self.carwashes[site].carsProject
            self.carsGoal = self.carsGoal + self.carwashes[site].carsGoal

            self.laborTotal = self.laborTotal + self.carwashes[site].laborTotal
            self.laborProject = self.laborProject + self.carwashes[site].laborProject
            self.laborGoal = self.laborGoal + self.carwashes[site].laborGoal

            self.salesTotalQ = self.salesTotalQ + self.carwashes[site].salesTotalQ
            self.salesProjectQ = self.salesProjectQ + self.carwashes[site].salesProjectQ
            self.salesGoalQ = self.salesGoalQ + self.carwashes[site].salesGoalQ

            self.carsTotalQ = self.carsTotalQ + self.carwashes[site].carsTotalQ
            self.carsProjectQ = self.carsProjectQ + self.carwashes[site].carsProjectQ
            self.carsGoalQ = self.carsGoalQ + self.carwashes[site].carsGoalQ

            self.laborTotalQ = self.laborTotalQ + self.carwashes[site].laborTotalQ
            self.laborProjectQ = self.laborProjectQ + self.carwashes[site].laborProjectQ
            self.laborGoalQ = self.laborGoalQ + self.carwashes[site].laborGoalQ

        self.salesPercent = self.get_percent(self.salesProject, self.salesGoal)
        self.carsPercent = self.get_percent(self.carsProject, self.carsGoal)
        self.laborPercent = self.get_percent(self.laborProject, self.laborGoal)
        self.salesPercentQ = self.get_percent(self.salesProjectQ, self.salesGoalQ)
        self.carsPercentQ = self.get_percent(self.carsProjectQ, self.carsGoalQ)
        self.laborPercentQ = self.get_percent(self.laborProjectQ, self.laborGoalQ)

        self.int_to_string()

def get_daysPIQ(month):
    # Figure out if this is a first, second or third month in quarter
    if month == 1 or month == 4 or month == 7 or month == 10:
        daysPIQ = today
    elif month == 2 or month == 5 or month == 8 or month == 11:
        daysPIQ = today + cd.monthrange(now.year, month2)[1]
    elif month == 3 or month == 6 or month == 9 or month == 12:
        daysPIQ = today + cd.monthrange(now.year, month2)[1] + cd.monthrange(now.year, month3)[1]
    return daysPIQ

def start_scraping(groups):
    with rq.Session() as s:
        p = s.post('https://www.statwatch.com/login', data=payload)
        url = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=total&start={}&granularity=6&todate=1&difftype=percentup&unittype=abs".format(date)
        r = s.get(url)
    soup = bsoup(r.content, "lxml") 
        
    divSites = soup.find_all("div", {"data-trayid": "site"})
    for div in divSites:
        liSites = div.find_all("li")
        for li in liSites:
            siteID = li.get_text()
            sites.update({siteID: CarWash(siteID)})
            sites[siteID].scrape_init()
            print(sites[siteID].name + " is scraped.")

    sites['02C'].dilute_sales(1)

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
                        sites[site].laborTotal,
                        sites[site].laborProject,
                        sites[site].laborGoal,
                        sites[site].laborPercent
                    ]

        for summary in summaries:
            if summary == sites[site].region or summary == sites[site].company or summary == 'All':
                summaries[summary].add_site(site, dataRow)

    for summary in summaries:
        summaries[summary].get_totals()

def get_filename():
    filename = "Projections " + dateName + ".xlsx"
    return filename

def create_workbook():
    # Save Data to XLS
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    format1 = workbook.add_format({'num_format': '#,##0'})
    format2 = workbook.add_format({'num_format': '0%'})
    # Red fill with dark red text.
    formatR = workbook.add_format({'bg_color':   '#FFC7CE',
                                    'font_color': '#9C0006'})
    # Light yellow fill with dark yellow text.
    formatY = workbook.add_format({'bg_color':   '#FFEB9C',
                                    'font_color': '#9C6500'})
    # Green fill with dark green text.
    formatG = workbook.add_format({'bg_color':   '#C6EFCE',
                                    'font_color': '#006100'})

    for data in dataSheets:
        df = pd.DataFrame(summaries[data].sheet, columns=summaries[data].titles, index=summaries[data].indexRow)
        df.to_excel(writer, sheet_name=data)
        worksheet = writer.sheets[data]
        worksheet.set_column('B:B', 40)
        worksheet.set_column('C:E', 10, format1)
        worksheet.set_column('G:I', 10, format1)
        worksheet.set_column('K:M', 10, format1)
        
        num_rows = len(df.index) + 1
        worksheet.set_column('F:F', 10, format2)
        color_range = "F2:F{}".format(num_rows)
        worksheet.conditional_format(color_range, {'type': 'cell',
                                    'criteria': 'less than',
                                    'value':  1,
                                    'format':   formatR,
                                    })
        worksheet.conditional_format(color_range, {'type': 'cell',
                                    'criteria': 'greater than',
                                    'value':  1,
                                    'format':   formatG,
                                    })
        worksheet.set_column('J:J', 10, format2)
        color_range = "J2:J{}".format(num_rows)
        worksheet.conditional_format(color_range, {'type': 'cell',
                                    'criteria': 'less than',
                                    'value':  1,
                                    'format':   formatR,
                                    })
        worksheet.conditional_format(color_range, {'type': 'cell',
                                    'criteria': 'greater than',
                                    'value':  1,
                                    'format':   formatG,
                                    })
        worksheet.set_column('N:N', 10, format2)
        color_range = "N2:N{}".format(num_rows)
        worksheet.conditional_format(color_range, {'type': 'cell',
                                    'criteria': 'greater than',
                                    'value':  1,
                                    'format':   formatR,
                                    })
        worksheet.conditional_format(color_range, {'type': 'cell',
                                    'criteria': 'less than',
                                    'value':  1,
                                    'format':   formatG,
                                    })

    writer.save()

def send_email(filename):
    # Email details
    text = """Hello,
    
Here are the projections and goals for all locations as of {}. Please see attached spreadhseet for a breakdown of these totals. 

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

    test = [myEmail, myGmail]
    real = [tBarrett, aShell, mWong, mGrimes, mCollins, myEmail]

    recipients = test
    
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
                        summaries['All'].salesTotalQ, 
                        summaries['LP14'].salesTotalQ, 
                        summaries['LP21'].salesTotalQ,  
                        summaries['All'].salesProjectQ, 
                        summaries['LP14'].salesProjectQ, 
                        summaries['LP21'].salesProjectQ,  
                        summaries['All'].salesGoalQ,  
                        summaries['LP14'].salesGoalQ, 
                        summaries['LP21'].salesGoalQ, 
                        summaries['All'].salesPercentQ,
                        summaries['LP14'].salesPercentQ,   
                        summaries['LP21'].salesPercentQ,  
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
                        summaries['LP21'].carsPercent,
                        summaries['All'].carsTotalQ,  
                        summaries['LP14'].carsTotalQ, 
                        summaries['LP21'].carsTotalQ, 
                        summaries['All'].carsProjectQ, 
                        summaries['LP14'].carsProjectQ, 
                        summaries['LP21'].carsProjectQ, 
                        summaries['All'].carsGoalQ, 
                        summaries['LP14'].carsGoalQ, 
                        summaries['LP21'].carsGoalQ, 
                        summaries['All'].carsPercentQ, 
                        summaries['LP14'].carsPercentQ, 
                        summaries['LP21'].carsPercentQ,
                        summaries['All'].laborTotal,  
                        summaries['LP14'].laborTotal, 
                        summaries['LP21'].laborTotal, 
                        summaries['All'].laborProject, 
                        summaries['LP14'].laborProject, 
                        summaries['LP21'].laborProject, 
                        summaries['All'].laborGoal, 
                        summaries['LP14'].laborGoal, 
                        summaries['LP21'].laborGoal, 
                        summaries['All'].laborPercent, 
                        summaries['LP14'].laborPercent, 
                        summaries['LP21'].laborPercent,
                        summaries['All'].laborTotalQ,  
                        summaries['LP14'].laborTotalQ, 
                        summaries['LP21'].laborTotalQ, 
                        summaries['All'].laborProjectQ, 
                        summaries['LP14'].laborProjectQ, 
                        summaries['LP21'].laborProjectQ, 
                        summaries['All'].laborGoalQ, 
                        summaries['LP14'].laborGoalQ, 
                        summaries['LP21'].laborGoalQ, 
                        summaries['All'].laborPercentQ, 
                        summaries['LP14'].laborPercentQ, 
                        summaries['LP21'].laborPercentQ
                    )

    text = MIMEText(text)
    html = MIMEText(html, 'html')
    msg.attach(text)
    msg.attach(html)    

    # XLS attachment
    att = MIMEApplication(output.getvalue(),_subtype="xlsx")
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

# Date variables
now = dt.datetime.now()
today = now.day 
date = dt.datetime.today().strftime("%m/%d/%Y")
dateName = dt.datetime.today().strftime("%m-%d-%Y")
daysIM = cd.monthrange(now.year, now.month)[1]

# Quarter Variables
month = str(now.month)
month2 = qDict[month]['others'][0]
month3 = qDict[month]['others'][1]
quarter = qDict[month]['quarter']
daysIQ = cd.monthrange(now.year, now.month)[1] + cd.monthrange(now.year, month2)[1] + cd.monthrange(now.year, month3)[1]
daysPIQ = get_daysPIQ(now.month)

goals = pd.read_csv('goals.csv')
goals.set_index(['SiteID', 'Metric'], inplace=True)

info = pd.read_csv('info.csv')
info.set_index(['SiteID'], inplace=True)

# Statwatch Login Form
payload = {
    'username': 'ant',
    'password': input('Statwatch password: '),
    'chain': 'scttaz'
}

sites = {}
summaries = {}
dataSheets = ['All', 'LP14', 'LP21', 'AZ', 'CA', 'NV']
output = io.BytesIO()

if __name__=='__main__': 
    start_scraping(dataSheets)
    create_workbook()   
    send_email(get_filename())
