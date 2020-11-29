import sys
import time
import pickle
import os.path
import datetime
import collections
import platform
import pymsgbox
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# Class for performing all actions in the browser
class Browser():
	def __init__(self, payload):
		self.payload = payload
		options = Options()
		#options.headless = True
		self.browser = webdriver.Chrome(ChromeDriverManager().install(), options=options)
		self.cookies = []
		self.platformOS = platform.system()

	# # Gets all the data for a row
	# def getRow(self, algorithms):
	# 	row = []
	# 	for person in algorithms:
	# 		row.append(person['liveTradingReturns'])
	# 		row.append(person['liveTradingPnL'])
	#
	# 	return row
	#
	# # Drive the actions to perform
	# def runOld(self):
	# 	self.login()
	# 	algorithms = self.getAlgorithms()
	# 	for i, alg in enumerate(algorithms):
	# 		person = self.getAlgorithmData(alg)
	# 		person = self.getLiveAlgorithmData(person)
	# 		algorithms[i] = person
	# 		person = {}
	#
	# 	return self.getRow(algorithms)

	# Login
	def login(self):
		url = 'https://ts2.clubinterconnect.com/carleton/home/login.do'
		self.go(url)

		username = self.browser.find_element_by_id("userid") # Username form field
		password = self.browser.find_element_by_id("password") # Password form field

		username.send_keys(self.payload[EMAIL_KEY])
		password.send_keys(self.payload[PASSWORD_KEY])

		submitButton = self.browser.find_element_by_id("submit")
		submitButton.click()

		time.sleep(1) # Need to sleep to allow for the redirection to posts

	def run(self):
		self.login()
		courtLinks = self.getDayLinks()

		# Gets the last day
		lastCourtDate = courtLinks[next(reversed(courtLinks))]
		self.getCalendarTable(lastCourtDate)
		# Gets the courts by times
		times = self.getByTime(lastCourtDate)
		self.openLinksByTimesAndCourts(times)

		#for date, courtDate in courtLinks.items():
		#	self.getCalendarTable(courtDate)
		#	courtDate.toString()


	def getCalendarTable(self, courtDate):
		print("\n---------\nNavigating to Date: " + courtDate.date + "\n---------\n")
		self.go(courtDate.pageLink)
		table = self.browser.find_elements_by_tag_name('table')

		body = table[0].find_elements_by_tag_name("tbody")
		rows = body[0].find_elements_by_tag_name("tr")

		for row in rows:
			if ("Available" not in row.text):
				continue
			for element in row.find_elements_by_tag_name('a'):
				link = element.get_attribute('href')
				if ("date" in link and "time" in link):
					courtDate.addCourtLink(link)
					# print(link)
			# print(row.find_elements_by_tag_name('a')[0].get_attribute('href'))
			# print(row.text)

	def getDayLinks(self):
		url = 'https://ts2.clubinterconnect.com/carleton/home/calendarDayView.do?id=7'
		self.go(url)
		links = self.browser.find_element_by_id('caldaylink')

		courtDates = collections.OrderedDict()

		currentDay = links.find_element_by_tag_name('span')
		print(currentDay.text)
		courtDates[currentDay.text] = CourtDate(currentDay.text, url)

		for link in links.find_elements_by_tag_name('a'):
			print(link.text)
			courtDates[link.text] = CourtDate(link.text, link.get_attribute('href'))

		return courtDates

	def getByTime(self, courtDate):
		times = collections.OrderedDict()
		format = "%I:%M %p"
		minTimeHour = 16 # 4pm
		minTimeMin = 0
		maxTimeHour = 22 # 10pm
		maxTimeMin = 00
		# Creates a dictionary of times with a list of court links
		# Checks if the time is within our threshold
		for link in courtDate.courtLinks:
			time = self.getTime(link)
			actualTime = datetime.datetime.strptime(time, format).time()
			if actualTime > datetime.time(minTimeHour,minTimeMin) and actualTime < datetime.time(maxTimeHour,maxTimeMin):
				if time in times:
					times[time].append(link)
				else:
					times[time] = [link]
		return times

	def openLinksByTimesAndCourts(self, times):
		maxCourt = None
		alreadyBooked = False
		shouldExit = False
		errorText = "You either timed out or you don't have permission to access this page."
		for time, links in times.items():
			courts = collections.OrderedDict()
			# Creates a dictionary of court links
			for link in links:
				court = int(link[link.find("item="):link.find("&date")].replace("item=",''))
				courts[court] = link

			# Sorts the links by the court number in descending order
			courtKeys = sorted(courts, key=lambda key: courts[key], reverse=True)

			for key in courtKeys:
				# maxCourt = max(courts, key=int)
				self.newTab(courts[key])
				# In-case it was already booked in the time it took to get here
				if errorText in self.browser.page_source:
					continue;

				if not alreadyBooked:
					#self.bookCurrentCourt()
					alreadyBooked = True
					result = self.messageBox('Booked Court', 'Booked Court ' + str(key) + ' at ' + str(time), ["Try Next", "Finish Booking"])
					print(result)

					if result == "Finish Booking":
						shouldExit = True
						break
					else
						alreadyBooked = False
						# TODO : Unbook?

			if shouldExit:
				break

		inp = input()
			# for court, link in courts:
			# 	if court != maxCourt:
			# 		self.newTab(link)

	def bookCurrentCourt(self):

		submitButton = self.browser.find_element_by_id("submit")
		submitButton.click()

		time.sleep(1) # Need to sleep to allow for the redirection to posts


	def getTime(self, link):
		timePart = link[link.find("time="):]
		timePart = timePart.replace('time=', '').replace('%20', ' ')
		#print(timePart)
		return timePart

	# Go to a specific url. Sleep for 1 seconds to ensure page loads. If error occurs, increase sleep time
	def go(self, url):
		self.browser.get(url)
		self.setCookies()
		time.sleep(1) # Need to sleep the allow for the page to load

	# Need to check OS because MacOS uses COMMAND+T but everyone else uses CONTROL+T
	def newTab(self, url):
		if self.platformOS == 'Darwin': # Darwin is MacOS, could also be Linux or Windows
			self.browser.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')
		else:
			self.browser.find_element_by_tag_name('body').sendKeys(Keys.CONTROL +"t");
		self.go(url)

	# Set the cookies for the current page because sometimes they are lost
	def setCookies(self):
		for cookie in self.cookies:
			self.browser.add_cookie(cookie)

	def messageBox(self, title, text, buttons=["OK", "Cancel"]):
		return pymsgbox.confirm(text, title=title, buttons=buttons)

class CourtDate():
	def __init__(self, date, pageLink=""):
		self.date = date
		self.pageLink = pageLink
		self.courtLinks = []

	def setPageLink(self, link):
		self.pageLink = link

	def addCourtLink(self, link):
		if link not in self.courtLinks:
			self.courtLinks.append(link)

	def toString(self):
		for link in self.courtLinks:
			print(self.date + " " + link)

# Class for accessing the Google Drive API
class GoogleDrive():
	def __init__(self):
		self.createAccessToken()

	# Create the access token for the user and save it to the file system for next use
	def createAccessToken(self):
		SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
		creds = None
		# The file token.pickle stores the user's access and refresh tokens, and is
		# created automatically when the authorization flow completes for the first
		# time.
		if os.path.exists('token.pickle'):
			with open('token.pickle', 'rb') as token:
				creds = pickle.load(token)
		# If there are no (valid) credentials available, let the user log in.
		if not creds or not creds.valid:
			if creds and creds.expired and creds.refresh_token:
				creds.refresh(Request())
			else:
				flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
				creds = flow.run_local_server()
			# Save the credentials for the next run
			with open('token.pickle', 'wb') as token:
				pickle.dump(creds, token)

		self.service = build('sheets', 'v4', credentials=creds)

	# Insert a row into a sheet by appending it.
	def insertData(self, row, sheetsID):
		# The ID and range of a sample spreadsheet.
		SPREADSHEET_ID = sheetsID
		RANGE_NAME = 'Data!A:K'
		VALUE_INPUT_OPTION = 'USER_ENTERED'
		INSERT_DATA_OPTION = 'INSERT_ROWS'

		row.append(str(datetime.date.today()))
		values = [row[::-1]]
		body = {'values': values}
		result = self.service.spreadsheets().values().append(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME, valueInputOption=VALUE_INPUT_OPTION, insertDataOption=INSERT_DATA_OPTION, body=body).execute()
		print('{0} cells updated.'.format(result['updates']['updatedCells']))

EMAIL = 1
PASSWORD = 2
SHEETS_ID = 3

EMAIL_KEY = 'user[email]'
PASSWORD_KEY = 'user[password]'
SHEETS_KEY = 'sheetsID'

def main():
	payload = {
		EMAIL_KEY: "<User Name>",
		PASSWORD_KEY: "<Password>",
		SHEETS_KEY: "<Sheets ID>"
	}

	payload[EMAIL_KEY] = sys.argv[EMAIL]
	payload[PASSWORD_KEY] = sys.argv[PASSWORD]
	payload[SHEETS_KEY] = sys.argv[SHEETS_ID]

	# if '/' in payload[SHEETS_KEY]:
	# 	payload[SHEETS_KEY] = payload[SHEETS_KEY].split('/')[-2]

	bowser = Browser(payload)
	row = bowser.run()

	# gdrive = GoogleDrive()
	# gdrive.insertData(row, payload[SHEETS_KEY])

if __name__ == '__main__':
	main()
