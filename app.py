from bs4 import BeautifulSoup as bsoup
import requests as rq

# Fill in your details here to be posted to the login form.
payload = {
    'username': 'ant',
    'password': 'lp14',
    'chain': 'scttaz'
}

# Use 'with' to ensure the session context is closed after use.
with rq.Session() as s:
    p = s.post('https://www.statwatch.com/login', data=payload)
    # print the html returned or something more intelligent to see if it's a successful login page.

    # The authorized request
    cMonth = "3"
    url1 = "https://www.statwatch.com/ajax/multi-site-table?groupID=all&compareSiteID=PREV&sort=site&sortd=asc&pc=0&activeStat=total&start="
    url2 = "/1/2017&granularity=6&todate=1&difftype=percentup&unittype=abs"
    r = s.get(url1 + cMonth + url2)
        
soup = bsoup(r.content, "lxml")
divs = soup.find_all("div", { "data-trayid" : "total" })
for div in divs:
    print(div.get_text())