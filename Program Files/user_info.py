#type in your mysql username
mysql_username = ""

#type in your mysql passowrd
mysql_password = ""

#type in your mysql host
mysql_host = "localhost"

#type in your mysql dbname
mysql_dbname = "reminderappdb"


#type in your email id
user_email = ""

#type in your email password
user_password = ""

#type in the name:email_id
mail_id_dict =  { 
			#"name1": "email_id1",
			#"name2": "email_id2"
			}

#type in your outlook cliend id
outlook_client_id = ""

#type in your outlook secret id
outlook_secret_id = ""




####dont change this:

#create the database if not present already
import mysql.connector
try:
	mydb = mysql.connector.connect(host=mysql_host,user=mysql_username,passwd=mysql_password)
	mycursor = mydb.cursor()
	mycursor.execute(f"CREATE DATABASE {mysql_dbname}")
	mycursor.close()
	mydb.close()
except Exception as e:
	pass

#create the table in the database if not present already
try:
	mydb = mysql.connector.connect(host=mysql_host,user=mysql_username,passwd=mysql_password, database=mysql_dbname)
	mycursor = mydb.cursor()
	mycursor.execute("CREATE TABLE REM(TEXT varchar(150), DATEANDTIME datetime)")
	mycursor.close()
	mydb.close()
except Exception as e:
	pass
####