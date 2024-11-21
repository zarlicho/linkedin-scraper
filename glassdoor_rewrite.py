from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import BatchHttpRequest
import time,json,urllib.parse,random,os,requests
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from datetime import datetime
from seleniumbase import SB
import pandas as pd
from bs4 import BeautifulSoup
from lxml import html

class GSheet:
	def __init__(self,excelFile):
		self.creds = None
		self.LOCAL_FILE_PATH = excelFile
		self.SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
		self.SPREADSHEET_ID = "1inxQ5SMuXUGRfhKibrcfXR6kCA_TcCMTMf5k8mDbK8I" # change this with your google sheet id!
		self.RANGE_NAME = "Sheet1"
		self.setup()
		self.service = build("sheets", "v4", credentials=self.creds)

	def setup(self):
		if os.path.exists("token.json"):
			self.creds = Credentials.from_authorized_user_file("token.json", self.SCOPES)
		if not self.creds or not self.creds.valid:
			if self.creds and self.creds.expired and self.creds.refresh_token:
				self.creds.refresh(Request())
			else:
				flow = InstalledAppFlow.from_client_secrets_file(
					"credentials.json", self.SCOPES
				)
				self.creds = flow.run_local_server(port=0)
			with open("token.json", "w") as token:
				token.write(self.creds.to_json())

	def get_sheet_names(self):
		"""Get the names of all sheets in the spreadsheet."""
		try:
			sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
			sheets = sheet_metadata.get('sheets', '')
			sheet_names = [sheet['properties']['title'] for sheet in sheets]
			return sheet_names
		except HttpError as err:
			print(err)
			return []

	def get_sheet(self):
		excel_file = self.LOCAL_FILE_PATH.split(".xlsx")[0]
		"""Export all sheets to an Excel file locally."""
		writer = pd.ExcelWriter(self.LOCAL_FILE_PATH, engine='xlsxwriter')
		result = self.service.spreadsheets().values().get(
			spreadsheetId=self.SPREADSHEET_ID, range=excel_file
		).execute()
		values = result.get("values", [])
		if not values:
			print(f"No data found in sheet: {excel_file}")
			return
		df = pd.DataFrame(values)
		df.to_excel(writer, sheet_name=excel_file, index=False, header=False)
		writer._save()
		print(f"Exported sheet {excel_file} to {excel_file}.xlsx")

	def update_cell_by_index(self, sheet_name, row, col, new_value):
		"""Update the value of a specific cell by its index."""
		try:
			cell_range = f"{sheet_name}!{chr(65 + col)}{row + 2}"
			values = [[new_value]]
			body = {'values': values}
			result = self.service.spreadsheets().values().update(
				spreadsheetId=self.SPREADSHEET_ID,
				range=cell_range,
				valueInputOption="RAW",
				body=body
			).execute()
			print(f"Updated cell {cell_range} with value '{new_value}'.")
		except HttpError as err:
			print(err)

	def add_row(self, sheet_name, row_data):
		"""Add a new row with the provided data."""
		try:
			values = [row_data]
			body = {'values': values}
			result = self.service.spreadsheets().values().append(
				spreadsheetId=self.SPREADSHEET_ID,
				range=sheet_name,
				valueInputOption="RAW",
				insertDataOption="INSERT_ROWS",
				body=body
			).execute()
			print(f"Added new row to {sheet_name} with data: {row_data}")
		except HttpError as err:
			print(err)

	def update_excel_cell(self, company, updates):
		"""Update the values for a specific company in the local Excel file."""
		try:
			df = pd.read_excel(self.LOCAL_FILE_PATH)
			company_row = df[df['Company Name'] == company].index

			if len(company_row) == 0:
				new_row = pd.DataFrame([updates])
				df = pd.concat([df, new_row], ignore_index=True)
				df.to_excel(self.LOCAL_FILE_PATH, index=False)
				print(f"Added new data for company: '{company}' in local Excel file.")
				return len(df) - 1  # Return the new row index
			else:
				for column, new_value in updates.items():
					if column in df.columns:
						df.loc[company_row, column] = new_value
					else:
						print(f"Column '{column}' not found in the local Excel file.")
				df.to_excel(self.LOCAL_FILE_PATH, index=False)
				print(f"Updated data for company: '{company}' in local Excel file.")
				return company_row[0]  # Return the existing row index
		except Exception as err:
			print(f"Error updating local Excel file: {err}")
			return None

	def update_locally(self, cell_updated):
		company = cell_updated.get('Company Name')
		print("updated company: ",company)
		if not company:
			print("No 'Company Name' specified in the update data.")
			return
		# Update local Excel file first
		row_index = self.update_excel_cell(company, cell_updated)
		if row_index is not None:
			for column, new_value in cell_updated.items():
				df = pd.read_excel(self.LOCAL_FILE_PATH)
				if column in df.columns:
					col = df.columns.get_loc(column)
					self.update_cell_by_index("GlassDoorData", row_index, col, new_value)
		else:
			print(f"Failed to update or add company: '{company}' in Google Sheets.")

class GlassdoorScraper:
	def __init__(self):
		self.sb = None
		self.company_location = None
		self.VIEW_NAME = "ATX Ventures"  # Enter VIEW_NAME here. OR leave it empty.
		self.GLASSDOOR_LOGIN_EMAIL = "czgojueycxqdjnvzjr@tmmbt.net"
		self.GLASSDOOR_LOGIN_PASSWORD = "czgojueycxqdjnvzjr@tmmbt.net"
		self.CRM_BASE_ID = 'appjvhsxUUz6o0dzo'
		self.CRM_BASE_Prospectus_Tabke = 'tblf4Ed9PaDo76QHH'
		self.API_KEY = 'patQIAmVOLuXelY42.df469e641a30f1e69d29195be1c1b1362c9416fffc0ac17fd3e1a0b49be8b961'
		self.WEBSCRAPER_BASE_ID = "appQfs70fHCsFgeUe"
		self.WEBSCRAPER_BASE_GLASSDOOR_TABLE_ID = "tbl2hHNNmdeHYSKMr"
		self.headers = {'Authorization': 'Bearer ' + self.API_KEY}
		self.Post_Header = {'Authorization': 'Bearer ' + self.API_KEY,'Content-Type': 'application/json'}
		self.AllRecordIds = []
		self.getInputCompanyTable()

	def cleanWebsiteURL(self,website):
		return website.replace("https:", "").replace("http:", "").replace("//", "").replace("www.", "").replace("\\", "").split(".com")[0].split(".co.uk")[0].split(".io")[0].split(".ai")[0].split(".tech")[0].split(".app")[0].split(".dev")[0].split(".mobi")[0].split(".cloud")[0].split(".network")[0].split(".digital")[0].split(".software")[0].split("/")[0].lower().strip()

	def random_sleep(self, min_seconds=1, max_seconds=5):
		time.sleep(random.uniform(min_seconds, max_seconds))

	def random_mouse_movements(self, element):
		action = self.sb.actions()
		for _ in range(random.randint(10, 30)):
			x_offset = random.randint(-10, 10)
			y_offset = random.randint(-10, 10)
			action.move_to_element_with_offset(element, x_offset, y_offset).perform()
			self.random_sleep(0.01, 0.1)

	def get_element(self,selector,by):
		try:
			self.sb.wait_for_element_visible(selector,by=by,timeout=10)
			if self.sb.is_element_visible(selector,by):
				try:
					return self.sb.find_element(selector,by)
				except Exception as e:
					return f"error: {e}"
			else:
				return f"error {selector} not found!"
		except Exception as e:
			print(e)
			pass
		return None

	def get_element_bs4(self,selector,pgSource,types):
		tree = html.fromstring(pgSource)
		if len(tree.xpath(selector)) != 0:
			if types=="text":
				return tree.xpath(selector)[0].text_content()
			else:
				return tree.xpath(selector)[0]
		else:
			return None
	
	def login_glassdoor(self):
		global status
		status = False
		for retry in range(3):
			try:
				print(f"Trying login {retry+1} times")
				self.sb.uc_open_with_reconnect("https://www.glassdoor.com/index.htm", 3)
				self.sb.type("#inlineUserEmail", self.GLASSDOOR_LOGIN_EMAIL + "\n", timeout=10, retry=2)
				self.sb.type("/html/body/div[3]/section[1]/div[2]/div/div/div[1]/div[1]/div/div/div/form/div[1]/div[1]/div/div[1]/input", self.GLASSDOOR_LOGIN_PASSWORD + "\n",timeout=10, retry=2)
				if self.sb.is_valid_url("https://www.glassdoor.com/Community/index.htm"):
					status = True
					break
				else:
					status = False
					continue
			except Exception as e:
				print(f"Login failed at {retry+1} times")
		return status
	def get_elements(self, selector, by):
		try:
			self.sb.wait_for_element_visible(selector, by=by, timeout=10)
			if self.sb.is_element_visible(selector, by):
				try:
					return self.sb.find_elements(selector, by, limit=50)
				except Exception as e:
					print(f"Error finding elements {selector}: {e}")
					return None
			else:
				print(f"Element {selector} not found!")
				return None
		except Exception as e:
			print(f"Exception in get_elements for {selector}: {e}")
			return None

	def scrape_company_page(self,GSht):
		for Records in self.AllRecordIds:
			CompanyName = Records["Company Name"]
			print("Company Name: ",CompanyName, Records["GD URL"])
			self.sb.wait_for_text_visible(text="Search", selector="//*[@id='UtilityNav']/div[1]/button",by="xpath",timeout=10)
			self.sb.open(Records["GD URL"])
			TotalReviews = self.get_element_bs4(selector='//*[@class="review-overview_reviewCount__hQpzR"]',pgSource=self.sb.get_page_source(),types="text")
			CompanyRating = self.get_element_bs4(selector='//*[@class="rating-headline-average_rating__J5rIy"]',pgSource=self.sb.get_page_source(),types="text")
			if TotalReviews:
				TotalReviews = TotalReviews.split(" ")[0].split("(")[1]
				print("total reviews: ",TotalReviews)
			if CompanyRating:
				CompanyRating = CompanyRating
				print("company rating: ",CompanyRating)
			EngagedEmployerElement = self.get_element_bs4(selector='//*[@id="__next"]/div/div[1]/div/main/div/div[1]/div[3]/div[2]/div/span/p',pgSource=self.sb.get_page_source(),types="text")
			EngagedEmployer = "Yes" if EngagedEmployerElement and "Engaged" in EngagedEmployerElement else "No"
			# print("engaged status: ",EngagedEmployer)
			BenefitsElement = self.get_element_bs4(selector='//*[@id="benefits"]/a/@href',pgSource=self.sb.get_page_source(),types="href")
			BenefitsRating = None
			BenefitsNoofReviews = None
			if BenefitsElement:
				BenefitsLink = BenefitsElement
				print(BenefitsLink)
				self.sb.open(f"https://www.glassdoor.com{BenefitsLink}?filter.employmentStatus=REGULAR")
				self.random_sleep()
				BenefitsRating = self.get_element_bs4(selector='//*[@class="css-1s4ou26"]',pgSource=self.sb.get_page_source(),types="text")
				BenefitsNoofReviews = self.get_element_bs4(selector='//*[@class="d-flex justify-content-center mb css-1uyte9r"]/span',pgSource=self.sb.get_page_source(),types="text")
				# print("benefits review:	", BenefitsNoofReviews)
				# print("benefits rating:	", BenefitsRating)
				HealthIns = self.sb.find_elements(selector="//*[@data-test='benefit-Health Insurance']",by="xpath",limit=50)
				if HealthIns:
					global HealthSta,HealthRat # Health Insurance 
					HealthSta,HealthRat = None,None
					for HelathInfo in HealthIns:
						HealthSta = self.get_element_bs4(selector='//*[@class="mr-xxsm strong css-1p6dnxi ecvyovn3"]',pgSource=self.sb.get_page_source(),types="text")
						HealthRat = self.get_element_bs4(selector="//*[@class='d-flex align-items-center css-1ffljup ecvyovn1']",pgSource=self.sb.get_page_source(),types="text")
						print("helath info:	",HealthSta,HealthRat)
				else:
					print("Health Insurance data not found!")
					HealthSta,HealthRat = None,None

				if BenefitsRating:
					BenefitsRating = BenefitsRating
				if BenefitsNoofReviews:
					BenefitsNoofReviews = BenefitsNoofReviews.replace("Ratings", "").replace("Rating", "").strip()
				
				self.sb.click_if_visible(selector="//*[@id='2']",by="xpath",timeout=5)
				# d-flex align-items-center css-1ffljup ecvyovn1
				RetirementReview = self.get_element_bs4(selector='//*[@class="d-flex align-items-center css-1ffljup ecvyovn1"]',pgSource=self.sb.get_page_source(),types="text")
				if RetirementReview:
					global RetireSta,RetireRat # Retirement Data
					RetireSta,RetireRat = None,None
					RetireSta = self.get_element_bs4(selector='//*[@class="d-flex align-items-center css-1ffljup ecvyovn1"]',pgSource=self.sb.get_page_source(),types="text").split(" ")[0]
					RetireRat = self.get_element_bs4(selector='//*[@class="d-inline-flex align-items-center css-1cub7fk ecvyovn2"]',pgSource=self.sb.get_page_source(),types="text").split("★")[0]
					print("retirement data: ",RetireRat,RetireSta)
				else:
					print("Retirement data not found!")
					RetireSta,RetireRat = None,None
				GLASSDOOR_ID = Records["GD URL"].split("EI_IE")[1].split(".")[0]
				print(" " * 9, "GD Overall Review:", CompanyRating)
				print(" " * 9, "GD # of Reviews (Overall):", TotalReviews)
				print(" " * 9, "GD Benefits Review:", BenefitsRating)
				print(" " * 9, "GD # of Reviews (Benefits):", int(BenefitsNoofReviews.replace(",", "")) if BenefitsNoofReviews and "-" not in BenefitsNoofReviews else 0)
				print(" " * 9, "GD Engaged Employer", EngagedEmployer)
				print(" " * 9, "Glassdoor ID:", GLASSDOOR_ID)
				print(" " * 9, "GD Retirement Review: ", RetireSta)
				print(" " * 9, "GD # of Reviews (Retirement): ", RetireRat)
				print(" " * 9, "GD Health Insurance Review: ", HealthSta)
				print(" " * 9, "GD # of Reviews (Health Insurance): ", HealthRat)

				json_post_data = {
					"Glassdoor Data Last Modified":str(datetime.now().strftime("%d/%m/%y")),
					"Company Name": CompanyName,
					"Glassdoor URL": Records["GD URL"],
					"Glassdoor ID": GLASSDOOR_ID,
					"GD Overall Review": float(CompanyRating) if CompanyRating and "-" not in CompanyRating else None,
					"GD # of Reviews (Overall)": int(TotalReviews) if TotalReviews else 0,
					"GD Benefits Review": float(BenefitsRating) if BenefitsRating else None,
					"GD # of Reviews (Benefits)": int(BenefitsNoofReviews.replace(",", "")) if BenefitsNoofReviews and "-" not in BenefitsNoofReviews else 0,
					"GD Retirement Review": float(RetireSta) if RetireSta else 0,
					"GD # of Reviews (Retirement)":float(RetireRat) if RetireRat else 0,
					"GD Health Insurance Review":float(HealthSta) if HealthSta else 0,
					"GD # of Reviews (Health Insurance)":float(RetireRat) if RetireRat else 0,
					"Glassdoor Engaged": EngagedEmployer,
				}
				GSht.update_locally(cell_updated=json_post_data)
				response = requests.request("GET", f"https://api.airtable.com/v0/{self.CRM_BASE_ID}/{self.CRM_BASE_Prospectus_Tabke}?filterByFormula=%7BCompany+Name%7D%3D%27{urllib.parse.quote_plus(CompanyName)}%27", headers=self.headers).json()
				CompanyRecordID = response["records"][0]["id"]
				
				data = {"fields": {
					"Glassdoor ID": GLASSDOOR_ID,
					"GD Overall Review": float(CompanyRating) if CompanyRating and "-" not in CompanyRating else None,
					"GD # of Reviews (Overall)": int(TotalReviews) if TotalReviews else 0,
					"GD Benefits Review": float(BenefitsRating) if BenefitsRating else None,
					"GD # of Reviews (Benefits)": int(BenefitsNoofReviews.replace(",", "")) if BenefitsNoofReviews and "-" not in BenefitsNoofReviews else 0,
					"GD Retirement Review": float(RetireSta) if RetireSta else 0,
					"GD # of Reviews (Retirement)":float(RetireRat) if RetireRat else 0,
					"GD Health Insurance Review":float(HealthSta) if HealthSta else 0,
					"GD # of Reviews (Health Insurance)":float(RetireRat) if RetireRat else 0,
					"Glassdoor Engaged": EngagedEmployer,
				}}

				json_data = json.dumps(data)
				response = requests.request("PATCH", f"https://api.airtable.com/v0/{self.CRM_BASE_ID}/{self.CRM_BASE_Prospectus_Tabke}/{CompanyRecordID}", headers=self.Post_Header, data=json_data)
				print(" " * 9, "CRM PATHCING STATUS: ", response.status_code)
				print(" " * 9, f"UPDATING {CompanyName} TO CRM DONE!")
				if response.status_code != 200:
					print("Error patching Airtable record:", response.json())

	def getInputCompanyTable(self):
		offset = ''
		while 1:
			CompanyTableURL = f'https://api.airtable.com/v0/{self.CRM_BASE_ID}/{self.CRM_BASE_Prospectus_Tabke}'
			if len(self.VIEW_NAME) > 1:
				OutputTable = requests.get(CompanyTableURL, headers=self.headers, params={'offset': offset, 'view': self.VIEW_NAME}).json()
			else:
				OutputTable = requests.get(CompanyTableURL, headers=self.headers, params={'offset': offset}).json()
			for Records in OutputTable["records"]:
				for recordsKey, recordsValue in Records.items():
					if recordsKey == "fields":
						SingleRecord = {}
						try:
							if recordsValue["Glassdoor URL"].startswith("http"):
								SingleRecord["Company Name"] = recordsValue["Company Name"]
								SingleRecord["GD URL"] = recordsValue["Glassdoor URL"]
								self.AllRecordIds.append(SingleRecord)
						except KeyError:
							print("Glassdoor URL Not Found!")
							continue
						print("Company: ",recordsValue["Company Name"])
			try:
				nextOffset = OutputTable["offset"]
				offset = nextOffset
			except KeyError:
				break

	def Main(self):
		gsht = GSheet(excelFile="GlassDoorData.xlsx") # matchh the sheet name in your google sheet, 
		gsht.get_sheet()
		with SB(uc=True) as Sb:
			self.sb = Sb
			if self.login_glassdoor():
				print("Login successfully")
				self.scrape_company_page(gsht)
			
if __name__ == "__main__":
	GD = GlassdoorScraper()
	GD.Main()