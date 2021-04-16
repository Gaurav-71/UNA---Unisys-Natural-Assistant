import mysql.connector
#from datetime import date
from datetime import datetime
import speech_recognition as sr
import pyttsx3
#import schedule
#import threading
from user_info import mysql_username, mysql_password, mysql_host, mysql_dbname
mydb = mysql.connector.connect(host=mysql_host,user=mysql_username,passwd=mysql_password, database=mysql_dbname)

if(mydb):
    print("Connection Successful!")
else:
    print("Unsuccessful.")

mycursor=mydb.cursor()
def AddNewReminder(text,dt):
	sql = ("INSERT INTO REM(TEXT,DATEANDTIME) VALUES(%s,%s)")
	reminder = (text,dt)
	mycursor.execute(sql,reminder)
	mydb.commit()
    #schedule.every(30).seconds.do(Remind)

# mycursor.execute("SELECT * FROM REM")
# myres=mycursor.fetchall()
# for row in myres:
#     print(row)