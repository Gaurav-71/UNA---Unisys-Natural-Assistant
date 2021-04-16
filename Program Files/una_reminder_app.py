import mysql.connector
import time
#from datetime import date
from datetime import datetime
import speech_recognition as sr
import pyttsx3
import schedule
import threading
# mydb = mysql.connector.connect(host="localhost",user="root",passwd="", database="reminderappdb")

# if(mydb):
#     print("Connection Successful!")
# else:
#     print("Unsuccessful.")


# mycursor = mydb.cursor()
#mycursor.execute("DELETE FROM REMINDER")
#mycursor.execute("CREATE DATABASE reminderappdb")
# mycursor.execute("CREATE TABLE REM(TEXT varchar(150), DATEANDTIME datetime)")

def ShowReminders():
    mycursor.execute("SELECT * FROM REM")
    myres=mycursor.fetchall()
    return myres

def ClearReminders(mycursor):
    print("Entered CR")
    mycursor.execute("SELECT * FROM REM")
    myres=mycursor.fetchall()
    for row in myres:
        if row[1] <= datetime.now():
            rem = row[0]
            #print(rem)
            for r in myres:
                mycursor.execute("DELETE FROM REM WHERE TEXT ='" +rem+"'")
                mydb.commit()


def Remind(mycursor):
    print("Entered Remind")
    engine = pyttsx3.init()
    mycursor.execute("SELECT * FROM REM")
    res = mycursor.fetchall()
    #print(res)
    for r in res:
        #print(datetime.datetime.strptime(r[1],"%Y/%m/%d %H:%M:%S"))
        #print(datetime.datetime.now())
        #if datetime.strptime("%Y/%m/%d %H:%M:%S",r[1]) <= datetime.now():
        if r[1] <= datetime.now():
            remind = r[0]
            print(remind)
            # engine.say(remind)
            # #print(remind)
            # engine.runAndWait()
            # engine.stop()
    # ClearReminders()
       
#t1=threading.Thread(target=AddNewReminder,args=("Client Meeting 555", "2021/02/21 06:44:00",))
#t2=threading.Thread(target=AddNewReminder,args=("Client Meeting 666", "2021/02/22 06:52:00",))

# ShowReminders()
#ClearReminders()
#schedule.every(1).minutes.do(Remind)
#t1.start()
#t2.start()
#t1.join()
# while(1):
#     #schedule.run_pending()
#     Remind()
#     time.sleep(60)


def runReminder():
	while(1):
	    #schedule.run_pending()
	    mydb = mysql.connector.connect(host="localhost",user="root",passwd="", database="reminderappdb")
	    Remind(mydb.cursor())
	    mydb.close()
	    time.sleep(30)