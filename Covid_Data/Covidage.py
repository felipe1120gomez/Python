import urllib.request
import urllib.parse
import urllib.error
import sqlite3
import json
import time
import ssl

print('This program reads data from the US National Center for Health Statistics and creates a visualization of recorded deaths by gender and age.')
time.sleep(1)
print('The program will use the API: https://data.cdc.gov/api/views/3apk-4u4f/rows.json?accessType=DOWNLOAD')
time.sleep(1)

api_json = input('Press "Enter" to start or type "exit" to quit the program: ')

if (len(api_json)) < 1:
    api_json = 'https://data.cdc.gov/api/views/3apk-4u4f/rows.json?accessType=DOWNLOAD'
else:
    quit()

conn = sqlite3.connect('covid_data.sqlite')
cur = conn.cursor()

cur.execute('''DROP TABLE IF EXISTS DeathsMale ''')
cur.execute('''DROP TABLE IF EXISTS DeathsFemale ''')

cur.executescript('''
CREATE TABLE IF NOT EXISTS DeathsMale (
    id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
    Age TEXT,
    Deaths INTEGER
);
CREATE TABLE IF NOT EXISTS DeathsFemale (
    id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE,
    Age TEXT,
    Deaths INTEGER
)
''')

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

connection = urllib.request.urlopen(api_json, context=ctx)
data = connection.read().decode()

try:
    js = json.loads(data)
except:
    print('==== Failure To Retrieve ====')
    print(data)  # We print in case unicode causes an error
    quit()


if 'data' not in js :
    print('==== The JSON file does not meet the requirements ====')
    print(data)
    quit()

for line in js['data']:
    age = line[12]
    gender = line[11]
    deaths = int(line[14])

    if gender == 'Male' :
        cur.execute('''INSERT INTO DeathsMale (Age, Deaths)
            VALUES ( ?, ? )''', ( age, deaths) )

    elif gender == 'Female' :
        cur.execute('''INSERT INTO DeathsFemale (Age, Deaths)
            VALUES ( ?, ? )''', ( age, deaths) )

conn.commit()

cur.close()

print('Creating JavaScript file')
time.sleep(2)

conn_1 = sqlite3.connect('file:covid_data.sqlite?mode=ro', uri=True)
cur_1 = conn_1.cursor()

cur_1.execute('SELECT id, Age, Deaths FROM DeathsFemale')
dictfemale = dict()
for row in cur_1:
    dictfemale[row[0]] = (row[1],row[2])

cur_1.execute('SELECT id, Age, Deaths FROM DeathsMale')
dictmale = dict()
for row in cur_1:
    dictmale[row[0]] = (row[1],row[2])

deathsfemale = list()
ages = list()
for (age_id, item) in list(dictfemale.items()):
    age = item[0]
    if age not in ages:
        ages.append(age)
    deathF = item[1]
    deathsfemale.append(deathF)

deathsmale = list()
for (age_id, item) in list(dictmale.items()):
    deathM = item[1]
    deathsmale.append(deathM)

gender = ('Female','Male')

fhand = open('gbar.js','w')
fhand.write("gline = [ ['Ages'")
for age in ages:
    fhand.write(",'"+age+"'")
fhand.write("]")

for gen in gender:
    fhand.write(",\n['"+gen+"'")
    if gen == 'Female':
        for i in deathsfemale:
            fhand.write(","+str(i))
    elif gen == 'Male':
        for j in deathsmale:
            fhand.write(","+str(j))
    fhand.write("]")

fhand.write("\n];\n")
fhand.close()

print('Output written to covid_data.sqlite')
time.sleep(1)
print("Output written to gbar.js")
time.sleep(1)
print("Open gbar.htm to visualize the data")
