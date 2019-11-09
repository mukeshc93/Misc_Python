#from bs4 import BeautifulSoup
import bs4
import pandas as pd
from urllib.request import urlopen
#pass the URL
url = urlopen("http://www.bseindia.com/markets/equity/EQReports/bulk_deals.aspx?expandable=3")
#read the source from the URL
readHtml = url.read()
#close the url
url.close()
#passing HTML to scrape it
soup = bs4.BeautifulSoup(readHtml, 'html.parser')
all_tables=soup.find_all('table')
right_table=soup.find('table', id="ctl00_ContentPlaceHolder1_gvbulk_deals")

Deal_Date = []
Security_Code = []
Security_Name = []
Client_Name = []
Deal_Type = []
Quantity = []
Price = []

for tr in right_table.find_all('tr')[3:]:
    tds = tr.find_all('td')

    column_1 = tds[0].string.strip()
    Deal_Date.append(column_1)

    column_2 = tds[1].string.strip()
    Security_Code.append(column_2)

    column_3 = tds[2].string.strip()
    Security_Name.append(column_3)

    column_4 = tds[3].string.strip()
    Client_Name.append(column_4)

    column_5 = tds[4].string.strip()
    Deal_Type.append(column_5)

    column_6 = tds[5].string.strip()
    Quantity.append(column_6)

    column_7 = tds[6].string.strip()
    Price.append(column_7)

columns = {'Deal_Date': Deal_Date, 'Security_Code': Security_Code, 'Security_Name': Security_Name, 'Client_Name': Client_Name, 'Deal_Type': Deal_Type, 'Quantity': Quantity, 'Price': Price}
df=pd.DataFrame(columns)
print(df)

df.to_csv('current.csv', sep=',',index = False)
df = pd.read_csv('current.csv',sep=',')
df1 = pd.read_csv('Bulk_17Jun2016.csv',sep=',')
df1.columns = ['Deal_Date', 'Security_Code', 'Security_Name', 'Client_Name', 'Deal_Type', 'Quantity', 'Price']
final = df.append(df1)
final.to_csv('current.csv', sep=',',index = False)


    

