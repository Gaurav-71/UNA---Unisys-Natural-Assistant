import datetime
import webbrowser
import wikipedia

import speech_recognition as sr
from pyttsx3.engine import Engine
import pyttsx3

import os
import sys
import time

from kivy.uix.widget import Widget
from kivy.properties import ObjectProperty
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.relativelayout import RelativeLayout
from kivy.uix.textinput import TextInput 
from kivy.uix.button import Button
from kivy.properties import ObjectProperty
from kivy.uix.popup import Popup 
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle
from kivy.lang import Builder
import kivy

from kivy.properties import ColorProperty
from automated_mail import *
# from una_reminder_app import *
from datetime import datetime
from una_add_reminder import *
import mysql.connector
from una_outlook import *

import win32com.client
from collections import namedtuple

#
from user_info import outlook_client_id, outlook_secret_id
from user_info import mysql_username, mysql_password, mysql_host, mysql_dbname
#

from O365 import Account, MSGraphProtocol
# from pygame import *
# import pygame
# KIVYIMAGE = pygame
import threading

import subprocess

import datefinder as dtf

kivy.require('2.0.0')  # replace with your current kivy version !
# Builder.load_file('una.kv')
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voices', voices[0].id)

isAuth = False
account = ''

class MyLayout(Widget):
	todos = []
	startedReminder = False
	def startReminder(self):
		threading.Thread(target=self.runReminder).start()


	def runReminder(self):
		while(1):
			mydb = mysql.connector.connect(host=mysql_host,user=mysql_username,passwd=mysql_password, database=mysql_dbname)
			self.remind(mydb)
			self.clearReminders(mydb)
			time.sleep(60)
			mydb.close()

	def clearReminders(self, mydb):
		print("Entered CR")

		mycursor = mydb.cursor()

		mycursor.execute("SELECT * FROM REM")
		myres=mycursor.fetchall()

		mycursor = mydb.cursor()
		for row in myres:
			if type(row) != None and row[1] <= datetime.datetime.now():
				rem = row[0]
				print(f"DELETE FROM REM WHERE TEXT = '{rem}'")
				deleteStatement = f"DELETE FROM REM WHERE TEXT = '{rem}'"
				mycursor.execute(deleteStatement)
				mydb.commit()

	def remind(self, mydb):
		reminder_text = ''
		mycursor = mydb.cursor()
		try:
			print("Entered Remind")

			mycursor.execute("SELECT * FROM REM")
			res = mycursor.fetchall()

			reminders = []
			for r in res:
				if r[1] <= datetime.datetime.now():
					remind = r[0]
					reminders.append(remind)
					reminder_text += f"\n{remind}"

			if len(reminder_text) > 0:
				top_grid = GridLayout(
					row_force_default = True,
					row_default_height = 40,
					col_force_default = True,
					col_default_width = 200
							)
				top_grid.cols = 1
				
				label = Label(text=reminder_text)
				
				# label.pos_hint = top_grid.width/2, top_grid.height/2
				
				top_grid.add_widget(label)
				cancelButton = Button(text="Close")
				top_grid.add_widget(cancelButton)

				# Instantiate the modal popup and display 
				popup = Popup(title ='Reminder', 
							  content = top_grid, 
							  size_hint =(None, None), size =(900, 450),
							  auto_dismiss = False
							  )   
				
				def onOpen(miss):
					self.speak("Reminder")
					for r in reminders:
						self.speak(r)

				popup.bind(on_open = onOpen)
				popup.open()

				cancelButton.bind(on_press = popup.dismiss)


		except Exception as e:
			print(e)

	def listenThread(self):
		threading.Thread(target=self.listen).start()

	def listen(self):
		r = sr.Recognizer()
		with sr.Microphone() as source:
			print('listening')
			audio = r.listen(source)
			try:
				command = r.recognize_google(audio)
				command = command.lower()
				print(command)
				self.analyzeText(command)
			except Exception as e:
				print(e)
				print('Error')
				self.speak("Sorry I didn't understand that, please try again.")

	def speechInput(self):
		r = sr.Recognizer()
		with sr.Microphone() as source:
			print('listening')
			audio = r.listen(source)
			try:
				command = r.recognize_google(audio)
				command = command.lower()
				print(command)
				return command
			except Exception as e:
				print(e)
				print('Error')
				self.speak("Sorry I didn't understand that, please try again.")
				return self.speechInput()

	def testInput(self):
		com = "wikipedia lion"
		self.analyzeText(com)

	def speak(self, message):
		print(message)
		engine.say(message)
		engine.runAndWait()
		engine.stop()

	def analyzeText(self, command):
		if not self.startedReminder:
			self.startReminder()
			self.startedReminder = True

		if "wikipedia" in command:

			# top_grid = GridLayout(
			# # row_force_default = True,
			# # row_default_height = 40,
			# # col_force_default = True,   
			# # col_default_width = 500
			#         )
			# top_grid.cols = 1

			# self.speak("Searching wikipedia")

			top_grid = GridLayout()
			top_grid.cols = 1

			command = command.replace("wikipedia ", "")

			results = wikipedia.summary(command, sentences=2)
			label = Label(text=results)
			label.size_hint = (1, None)
			label.v_align = "middle"
			label.h_align = "middle"
			label.font_size = '18dp'
			# label.pos_hint = top_grid.width/2, top_grid.height/2
			label.bind(size=label.setter('text_size'))
			top_grid.add_widget(label)

			button_grid = RelativeLayout()
			button_grid.cols = 1
			cancelButton = Button(text="Close", size_hint =(.1, .075),pos_hint ={'center_x':.95, 'center_y':.95})
			button_grid.add_widget(cancelButton)

			top_grid.add_widget(button_grid)
			# Instantiate the modal popup and display 
			popup = Popup(title =f'Summary of "{command}"',
						  content = top_grid,
						  auto_dismiss = False)   
			
			def onOpen(miss):
				# threading.Thread(target=self.speak, args=(results,)).start()
				self.speak(results)

			popup.bind(on_open = onOpen)

			cancelButton.bind(on_press=popup.dismiss)
			popup.open()

			
			cancelButton.bind(on_press=popup.dismiss)
		# elif "open youtube" in command:
		#     webbrowser.open("www.youtube.com")
		elif "what can you do" in command:
			self.speak("Hello! I am Unisys Natural Assistant. Let me take you through what I am capable of!")
			self.speak("Want me to send your emails? Done!")
			self.speak("Want me to search wikipedia? Done!")
			self.speak("What about summarize your meetings and make your life easier? Done!")
			self.speak("Want me to open your favorite website? Done!")
			self.speak("Want me to take you through your day? Done!")
			self.speak("Want me to remind you that I am the best? Done!")
			self.speak("Okay! See you soon! sending you back to my creators!")

		elif "open" in command:
			command = command.replace("open", "").strip()

			self.speak(f"opening {command}")
			print(command)
			webbrowser.open(f"www.{command}.com")

		# elif "add" in command and ("todo" in command or "to do" in command):
		#     r = sr.Recognizer()
		#     with sr.Microphone() as source:
		#         print('In Add Todo')
		#         audio = r.listen(source)
		#         time.sleep(1)
		#         try:
		#             todo = r.recognize_google(audio)
		#             todos.append(todo.lower())
		#             print(todo)
		#         except:
		#             print("Error")
		#             self.speak("Sorry I didn't understand that, please try again.")

		# elif "show" in command and "todo" in command:
		#     for todo in todos:
		#         self.speak(todo)

		elif "mail" in command:
			s = get_subject()
			m = get_mail()
			get_receivers()
			attatch = attachments()
			receivers_string = ",".join(str(x) for x in receivers)
			# message = MIMEMultipart("alternative")
			# message["Subject"] = subject
			# message["From"] = sender
			# message["To"] = receivers_string
			# text = MIMEText(mail, "plain")
			# message.attach(text)
			# attachments()
			# try:
			#     with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
			#         smtp.login(sender, EMAIL_PASSWORD)
			#         smtp.sendmail(sender, receivers, message.as_string())
			#         print("Done")
			# except Exception as e:
			#     print(e)


			top_grid = GridLayout(
			row_force_default = False,
			row_default_height = 20,
			col_force_default = False,
			col_default_width = 200
					)
			top_grid.cols = 2

			top_grid.add_widget(Label(text="Subject: "))
			subject = TextInput(text = s, multiline=True)
			top_grid.add_widget(subject)

			top_grid.add_widget(Label(text="Body: "))
			mail = TextInput(text = m, multiline=True)
			top_grid.add_widget(mail)

			top_grid.add_widget(Label(text="Receiver(s): "))
			receiver= TextInput(text = receivers_string, multiline=True)
			top_grid.add_widget(receiver)

			if attatch != None:
				top_grid.add_widget(Label(text="Attatchment(if any) : "))
				attatch = TextInput(text = attatch, multiline=True)
				top_grid.add_widget(attatch)

			cancelButton = Button(text="Cancel")
			top_grid.add_widget(cancelButton)
			closeButton = Button(text="Send")
			top_grid.add_widget(closeButton)

			# Instantiate the modal popup and display 
			popup = Popup(title ='Email', 
						  content = top_grid, 
						  size_hint =(None, None), size =(900, 450),
						  auto_dismiss = False)   
			popup.open()

			def onClose(miss):
				popup.dismiss()
				message = MIMEMultipart("alternative")
				
				message["Subject"] = subject.text
				message["From"] = sender
				message["To"] = receiver.text
				text = MIMEText(mail.text, "plain")
				message.attach(text)

				if attatch != None:
					try:
						files = os.listdir()
						filename = attatch.text
						print(filename)
						
						with open(os.getcwd()+ "\\" + filename , "rb") as attachment:
							part = MIMEBase("application", "octet-stream")
							part.set_payload(attachment.read())

						encoders.encode_base64(part)

						part.add_header(
							"Content-Disposition",
							f"attachment; filename= {filename}",
						)

						message.attach(part)
					except Exception as e:
						print(e)

				try:
					with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
						smtp.login(sender, EMAIL_PASSWORD)
						smtp.sendmail(sender, receivers, message.as_string())
						print("Done")
				except Exception as e:
					print(e)

			closeButton.bind(on_press = onClose)
			cancelButton.bind(on_press = popup.dismiss)


			

		elif "directory" in command:
			print(os.getcwd())

		# elif "reminder" in command:
		#     MyPopup().open()

		# elif "show" in command and "reminder" in command:
		#     myres = ShowReminders()
		#     for res in myres:
		#         self.speak(res)

		elif "add" in command and "reminder" in command:
			self.speak("Name of reminder")
			s = self.speechInput()

			top_grid = GridLayout(
				row_force_default = True,
				row_default_height = 40,
				col_force_default = True,
				col_default_width = 200
						)
			top_grid.cols = 2

			while True:
				try:
					self.speak('Date')
					d = f'{next(dtf.find_dates(self.speechInput()))}'.split(' ')[0]
					print(d)
					break
				except:
					pass
			while True:
				try:
					self.speak('Time')
					t = f'{next(dtf.find_dates(self.speechInput()))}'.split(' ')[1]
					print(t)
					break
				except:
					pass

			top_grid.add_widget(Label(text="Subject: "))
			subject = TextInput(text = s, multiline=True)
			top_grid.add_widget(subject)

			top_grid.add_widget(Label(text="Date(YYYY/MM/DD): "))
			date = TextInput(text = d, multiline=False)
			top_grid.add_widget(date)

			top_grid.add_widget(Label(text="Start Time(hh:mm:ss): "))
			t = TextInput(text = t, multiline=False)
			top_grid.add_widget(t)

			cancelButton = Button(text="Cancel")
			top_grid.add_widget(cancelButton)
			closeButton = Button(text="Set")
			top_grid.add_widget(closeButton)

			# Instantiate the modal popup and display 
			popup = Popup(title ='Set Reminder', 
						  content = top_grid, 
						  size_hint =(None, None), size =(900, 450),
						  auto_dismiss = False)   
			popup.open()

			def onClose(miss):
				popup.dismiss()
				AddNewReminder(subject.text, f'{date.text} {t.text}')

			closeButton.bind(on_press = onClose)
			cancelButton.bind(on_press = popup.dismiss)

		elif "calendar" in command:
			global isAuth
			global account
			if not isAuth:
				event = namedtuple("event", "Start Subject Duration")
				# #from here


				CLIENT_ID = outlook_client_id
				SECRET_ID = outlook_secret_id
				credentials = (CLIENT_ID, SECRET_ID)



				# protocol = MSGraphProtocol() 
				# scopes = ['Calendars.Read', 'Calendars.ReadWrite']
				# account = Account(credentials, protocol=protocol)

				protocol = MSGraphProtocol() 
				scopes = ['Calendars.Read', 'Calendars.ReadWrite']
				account = Account(credentials, protocol=protocol)
				account.con.scopes = account.protocol.get_scopes_for(scopes)
				print(account.con.get_authorization_url(scopes = scopes)[0])
				webbrowser.open(account.con.get_authorization_url(scopes = scopes)[0])

				top_grid = GridLayout(
					row_force_default = True,
					row_default_height = 100,
					col_force_default = True,
					col_default_width = 800
							)
				top_grid.cols = 1
				
				# label = Label(text=reminder_text)

				top_grid.add_widget(Label(text="URL: "))
				url = TextInput()
				top_grid.add_widget(url)
				# label.pos_hint = top_grid.width/2, top_grid.height/2
				

				bottom_grid = GridLayout(
					row_force_default = True,
					row_default_height = 40,
					col_force_default = True,
					col_default_width = 100
							)
				bottom_grid.cols = 2


				cancelButton = Button(text="Cancel")
				bottom_grid.add_widget(cancelButton)

				submitButton = Button(text="Submit")
				bottom_grid.add_widget(submitButton)

				top_grid.add_widget(bottom_grid)

				# submitButton = Button(text="Submit")
				# top_grid.add_widget(submitButton)

				# Instantiate the modal popup and display 
				popup = Popup(title ='Paste Authenticated URL', 
							  content = top_grid, 
							  size_hint =(None, None), size =(900, 450),
							  auto_dismiss = False
							  )   

				popup.open()

				def onSubmit(miss):
					print(url.text)
					if account.con.request_token(url.text, scopes = scopes):
					   print('Authenticated!')
					   global isAuth 
					   isAuth = True

					popup.dismiss()

				submitButton.bind(on_press = onSubmit)
				cancelButton.bind(on_press = popup.dismiss)

				

				# if account.authenticate(scopes=scopes):
				#     print('Authenticated!')
				#     isAuth = True

			if "show" in command and isAuth:

				top_grid = GridLayout(
					row_force_default = True,
					row_default_height = 40,
					col_force_default = True,
					col_default_width = 200
							)
				top_grid.cols = 2

				while True:
					try:
						self.speak('End Date')
						d = f'{next(dtf.find_dates(self.speechInput()))}'.split(' ')[0]
						print(d)
						break
					except:
						pass
				# while True:
				#     try:
				#         self.speak('End Time')
				#         t = f'{next(dtf.find_dates(self.speechInput()))}'.split(' ')[1]
				#         print(t)
				#         break
				#     except:
				#         pass


				top_grid.add_widget(Label(text="End Date(YYYY/MM/DD): "))
				endDate = TextInput(text = d, multiline=False)
				top_grid.add_widget(endDate)

				# top_grid.add_widget(Label(text="End Time(hh:mm:ss): "))
				# endTime = TextInput(text = t, multiline=False)
				# top_grid.add_widget(endTime)

				cancelButton = Button(text="Cancel")
				top_grid.add_widget(cancelButton)
				closeButton = Button(text="Search")
				top_grid.add_widget(closeButton)

				# Instantiate the modal popup and display 
				popup = Popup(title ='Calendar', 
							  content = top_grid, 
							  size_hint =(None, None), size =(900, 450),
							  auto_dismiss = False)   
				popup.open()

				def onClose(miss):
					popup.dismiss()
					schedule = account.schedule()
					calendar = schedule.get_default_calendar()
					print(endDate)
					end = list(map(int, str(endDate.text).split('-')))
					# eTime = list(map(int, str(endTime.text).split(':')))

					q = calendar.new_query('start').greater_equal(datetime.datetime.today())
					q.chain('and').on_attribute('end').less_equal(datetime.datetime(end[0], end[1], end[2],  5, 30))
					events = calendar.get_events(query = q, include_recurring=False)
					
					allevents = []
					allevents_text = ""
					for event in events:
						print(event)
						new_event = parse_event(event)
						print(new_event["date"], new_event["subject"], new_event["from"], new_event["to"])
						allevents.append(f'{new_event["subject"]} on {new_event["date"]} from {new_event["from"]} to {new_event["to"]}')
						allevents_text = f'{new_event["subject"]} on {new_event["date"]} from {new_event["from"]} to {new_event["to"]}' + allevents_text
						allevents_text += "\n"

					print(allevents_text)
					print(q)

					if len(allevents_text) > 0:
						top_grid1 = GridLayout(
					row_force_default = True,
					row_default_height = 60,
					col_force_default = True,
					col_default_width = 900
							)

						top_grid1.cols = 1
						
						for e in allevents[::-1]:
							label1 = Label(text=e)
							top_grid1.add_widget(label1)

						# label.pos_hint = top_grid.width/2, top_grid.height/2
						   
						cancelButton1 = Button(text="Close")
						top_grid1.add_widget(cancelButton1)

						# Instantiate the modal popup and display 
						popup1 = Popup(title ='All Events', 
									  content = top_grid1, 
									  size_hint =(None, None), size =(900, 450),
									  auto_dismiss = False
									  )   
						
						def onOpen1(miss):
							self.speak("Your Schedule Looks Like This")
							for e in allevents[::-1]:
								self.speak(e)

						popup1.bind(on_open = onOpen1)
						popup1.open()

						cancelButton1.bind(on_press = popup1.dismiss)




				closeButton.bind(on_press = onClose)
				cancelButton.bind(on_press = popup.dismiss)

			elif "add" in command and isAuth:
				# add_event(account)

				top_grid = GridLayout(
				row_force_default = True,
				row_default_height = 40,
				col_force_default = True,
				col_default_width = 200
						)
				top_grid.cols = 2

				# global startDate
				# global endDate, startTime, endTime
				self.speak("Subject of the event")
				s = self.speechInput()

				while True:
					try:
						self.speak('Date')
						d = f'{next(dtf.find_dates(self.speechInput()))}'.split(' ')[0]
						print(d)
						break
					except:
						pass
				while True:
					try:
						self.speak('Start Time')
						t = f'{next(dtf.find_dates(self.speechInput()))}'.split(' ')[1]
						print(t)
						break
					except:
						pass

				while True:
					try:
						self.speak('End Time')
						t1 = f'{next(dtf.find_dates(self.speechInput()))}'.split(' ')[1]
						print(t)
						break
					except:
						pass

				top_grid.add_widget(Label(text="Subject: "))
				subject = TextInput(text = s, multiline=True)
				top_grid.add_widget(subject)

				top_grid.add_widget(Label(text="Start Date(DD-MM-YYYY): "))
				startDate = TextInput(text = d, multiline=False)
				top_grid.add_widget(startDate)

				top_grid.add_widget(Label(text="Start Time(hh:mm): "))
				startTime= TextInput(text = t, multiline=False)
				top_grid.add_widget(startTime)

				top_grid.add_widget(Label(text="End Date(DD-MM-YYYY): "))
				endDate = TextInput(text = d, multiline=False)
				top_grid.add_widget(endDate)

				top_grid.add_widget(Label(text="End Time(hh:mm): "))
				endTime = TextInput(text = t1, multiline=False)
				top_grid.add_widget(endTime)

				cancelButton = Button(text="Cancel")
				top_grid.add_widget(cancelButton)
				closeButton = Button(text="Set")
				top_grid.add_widget(closeButton)

				# Instantiate the modal popup and display 
				popup = Popup(title ='Outlook Calendar', 
							  content = top_grid, 
							  size_hint =(None, None), size =(900, 450),
							  auto_dismiss = False)   
				popup.open()

				def onClose(miss):
					try:
						popup.dismiss()
						schedule = account.schedule()
						calendar = schedule.get_default_calendar()
						new_event = calendar.new_event()
						new_event.subject = subject.text
						start = list(map(int, str(startDate.text).split('-')))
						end = list(map(int, str(endDate.text).split('-')))
						sTime = list(map(int, str(startTime.text).split(':')))
						eTime = list(map(int, str(endTime.text).split(':')))

						print(start, end, sTime, eTime)
						# new_event.start = datetime.datetime(2021, 3, 7, 14, 30)
						# new_event.end = datetime.datetime(2021, 3, 7, 15, 0)
						new_event.start = datetime.datetime(start[0], start[1], start[2], sTime[0], sTime[1])
						new_event.end = datetime.datetime(end[0], end[1], end[2],  eTime[0], eTime[1])
						new_event.save()

					except Exception as e:
						print(e)

				closeButton.bind(on_press = onClose)
				cancelButton.bind(on_press = popup.dismiss)



		elif "mom" in command or (("minutes " in command or "summarize" in command or "summarise" in command) and "meeting" in command ) :
			top_grid = GridLayout(
				row_force_default = True,
				row_default_height = 100,
				col_force_default = True,
				col_default_width = 800
						)
			top_grid.cols = 1
			
			# label = Label(text=reminder_text)

			top_grid.add_widget(Label(text="File Name: "))
			fileName = TextInput()
			top_grid.add_widget(fileName)
			# label.pos_hint = top_grid.width/2, top_grid.height/2
			


			bottom_grid = GridLayout(
				row_force_default = True,
				row_default_height = 40,
				col_force_default = True,
				col_default_width = 100
						)
			bottom_grid.cols = 2


			cancelButton = Button(text="Cancel")
			bottom_grid.add_widget(cancelButton)

			submitButton = Button(text="Submit")
			bottom_grid.add_widget(submitButton)

			top_grid.add_widget(bottom_grid)

			# Instantiate the modal popup and display 
			popup = Popup(title ='Minutes Of Meeting', 
						  content = top_grid, 
						  size_hint =(None, None), size =(900, 450),
						  auto_dismiss = False
						  )   

			popup.open()

			def onSubmit(miss):
				subprocess.Popen(f"python una_mom.py {fileName.text}")
				popup.dismiss()

			submitButton.bind(on_press = onSubmit)
			cancelButton.bind(on_press = popup.dismiss)

		else:
			self.speak('Error')

	def wish(self):
		hour = int(datetime.datetime.now().hour)
		if hour >= 0 and hour <= 12:
			self.speak("Good morning !")
		elif hour > 12 and hour < 18:
			self.speak("Good afternoon !")
		else:
			self.speak("Good Evening")
		self.speak("I'm here to help ! Ask me anything")


	def hello(self):
		pass
		# self.add_widget(Label(text=f"Name:bye"))
		#self.speak("Hello")

class unaApp(App):
	def build(self):
		myLayout = MyLayout()
		# myLayout.startReminder()
		return myLayout


if __name__ == '__main__':
	# reminderThread = threading.Thread(target=runReminder)
	# reminderThread.start()
	unaApp().run()