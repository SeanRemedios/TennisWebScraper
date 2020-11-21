import sys
import time
import pickle
import os.path
import datetime
import collections
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Class for performing all actions in the browser
class Browser():
	def __init__(self, payload):
		self.payload = payload
		options = Options()
		options.headless = True
		self.browser = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
		self.cookies = []

	# Gets all the data for a row
	def getRow(self, algorithms):
		row = []
		for person in algorithms:
			row.append(person['liveTradingReturns'])
			row.append(person['liveTradingPnL'])

		return row

	# Drive the actions to perform
	def runOld(self):
		self.login()
		algorithms = self.getAlgorithms()
		for i, alg in enumerate(algorithms):
			person = self.getAlgorithmData(alg)
			person = self.getLiveAlgorithmData(person)
			algorithms[i] = person
			person = {}

		return self.getRow(algorithms)

	# Login to quantopian
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

		for date, courtDate in courtLinks.items():
			self.getCalendarTable(courtDate)
			courtDate.toString()


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

	# Find all the algorithms in the table
	def getAlgorithms(self):
		url = 'https://www.quantopian.com/algorithms'
		self.go(url)

		table = self.browser.find_elements_by_id("algorithms-table")
		rows = table[0].find_elements_by_tag_name("tr") # get all of the rows in the table
		rows = rows[1:]
		algorithms = []
		person = {}
		# Get the name of the algorithm and the url to go to
		for row in rows:
			person['name'] = ' '.join(row.text.split(' ')[0:2])
			print(person['name'])

			person['algURL'] = row.find_elements_by_tag_name('a')[0].get_attribute('href')
			print(person['algURL'])

			algorithms.append(person)
			person = {}

		return algorithms

	# Get the live trading link
	def getAlgorithmData(self, person):
		url = person['algURL']
		self.go(url)

		liveTradingLink = self.browser.find_elements_by_id('live-trading-link')
		person['liveAlgURL'] = [] if liveTradingLink == [] else liveTradingLink[0].find_elements_by_tag_name('a')[0].get_attribute('href')
		# print(person['liveAlgURL'])
		return person

	# Get all the live trading data
	def getLiveAlgorithmData(self, person):
		url = person['liveAlgURL']
		if url != []:
			self.go(url)

			# Add any other pertinent data
			liveTradingReturns = self.browser.find_elements_by_id('livetrading-stats-returns')
			liveTradingPnL = self.browser.find_elements_by_id('livetrading-stats-dollarpnl')

			person['liveTradingReturns'] = '--' if liveTradingReturns == [] else liveTradingReturns[0].text
			person['liveTradingPnL'] = '--' if liveTradingPnL == [] else liveTradingPnL[0].text

			print(person['name'] + ' - % Returns: ' + person['liveTradingReturns'])
			print(person['name'] + ' - $ P/L: ' + person['liveTradingPnL'])

		else: # Algorithm is not live trading so just add blanks
			person['liveTradingReturns'] = '--'
			person['liveTradingPnL'] = '--'

		return person

	# Go to a specific url. Sleep for 1.5 seconds to ensure page loads. If error occurs, increase sleep time
	def go(self, url):
		self.browser.get(url)
		self.setCookies()
		time.sleep(1.5) # Need to sleep the allow for the page to load

	# Set the cookies for the current page because sometimes they are lost
	def setCookies(self):
		for cookie in self.cookies:
			self.browser.add_cookie(cookie)

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
