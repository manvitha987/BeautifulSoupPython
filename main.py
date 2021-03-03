from bs4 import BeautifulSoup
import json
import pyodbc

soup = BeautifulSoup(open("C:/Users/Admin/Downloads/EP.html"), 'html.parser')

tabledata = soup.find("tbody")

rows = tabledata.findChildren("tr", recursive=False)

tList = []

for row in rows:
    col = row.findChildren("td")
    onerow = dict()
    onerow["Date"] = col[1].text
    onerow["DocumentType"] = col[2].find("a").text
    onerow["Procedure"] = col[3].text.replace('\xa0', '')
    onerow["NumberOfPages"] = col[4].text

    tList.append(onerow)

# with open('Manvitha.json', 'w') as outfile:
# json.dump(tList, outfile)

jsontable = json.dumps(tList)
# print(jsontable)


conn = pyodbc.connect('driver={SQL Server};'
                      'Server=DESKTOP-R6EOG0D;'
                      'Database=Documents;'
                      'trusted_Connection=yes;')
cursor = conn.cursor()
count = 0
for i in tList:
    date = i['Date']
    type = i['DocumentType']
    procedure = i['Procedure']
    pages = i['NumberOfPages']

    try:
        cmd = '''Insert into DocumentTable values(?,?,?,?)'''
        cursor.execute(cmd, date, type, procedure, pages)
        # print('record inserted')
        count = count + 1

    except Exception as e:
        print(e)

conn.commit()
print(count, ' records inserted')
