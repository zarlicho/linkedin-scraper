from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time,json,requests,winsound,os.path

import pandas as pd
from openpyxl import load_workbook
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import BatchHttpRequest

class GSheet:
    def __init__(self,excelFile):
        self.creds = None
        self.LOCAL_FILE_PATH = excelFile
        self.SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        self.SPREADSHEET_ID = "1zczzINj9Ae3qfx4FTdDIVaYwJ9tQFOLnGyAgVkY1M_c" # change this with your google sheet id!
        self.RANGE_NAME = "Sheet1"
        self.setup()

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

    def get_sheet_names(self, service):
        """Get the names of all sheets in the spreadsheet."""
        try:
            sheet_metadata = service.spreadsheets().get(spreadsheetId=self.SPREADSHEET_ID).execute()
            sheets = sheet_metadata.get('sheets', '')
            sheet_names = [sheet['properties']['title'] for sheet in sheets]
            return sheet_names
        except HttpError as err:
            print(err)
            return []

    def get_sheet(self, service):
        excel_file = self.LOCAL_FILE_PATH.split(".xlsx")[0]
        """Export all sheets to an Excel file locally."""
        writer = pd.ExcelWriter(self.LOCAL_FILE_PATH, engine='xlsxwriter')
        result = service.spreadsheets().values().get(
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

    def update_cell_by_index(self, service, sheet_name, row, col, new_value):
        """Update the value of a specific cell by its index."""
        try:
            cell_range = f"{sheet_name}!{chr(65 + col)}{row + 2}"
            values = [[new_value]]
            body = {'values': values}
            result = service.spreadsheets().values().update(
                spreadsheetId=self.SPREADSHEET_ID,
                range=cell_range,
                valueInputOption="RAW",
                body=body
            ).execute()
            print(f"Updated cell {cell_range} with value '{new_value}'.")
        except HttpError as err:
            print(err)

    def add_row(self, service, sheet_name, row_data):
        """Add a new row with the provided data."""
        try:
            values = [row_data]
            body = {'values': values}
            result = service.spreadsheets().values().append(
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
            company_row = df[df['Company'] == company].index

            if len(company_row) == 0:
                new_row = pd.DataFrame([updates])
                df = pd.concat([df, new_row], ignore_index=True)
                df.to_excel(self.LOCAL_FILE_PATH, index=False)
                print(f"Added new data for company '{company}' in local Excel file.")
                return len(df) - 1  # Return the new row index
            else:
                for column, new_value in updates.items():
                    if column in df.columns:
                        df.loc[company_row, column] = new_value
                    else:
                        print(f"Column '{column}' not found in the local Excel file.")
                df.to_excel(self.LOCAL_FILE_PATH, index=False)
                print(f"Updated data for company '{company}' in local Excel file.")
                return company_row[0]  # Return the existing row index
        except Exception as err:
            print(f"Error updating local Excel file: {err}")
            return None

    def update_locally(self, service, cell_updated):
        company = cell_updated.get('Company')
        if not company:
            print("No 'Company' specified in the update data.")
            return
        # Update local Excel file first
        row_index = self.update_excel_cell(company, cell_updated)
        if row_index is not None:
            for column, new_value in cell_updated.items():
                df = pd.read_excel(self.LOCAL_FILE_PATH)
                if column in df.columns:
                    col = df.columns.get_loc(column)
                    self.update_cell_by_index(service, "LinkedInCompanyInfo", row_index, col, new_value)
        else:
            print(f"Failed to update or add company '{company}' in Google Sheets.")

class Linkedin:
	def __init__(self):
		# credential 
		self.VIEW_NAME = "Dwayne View"
		self.LinkedIN_LOGIN_EMAIL = "dg135862@gmail.com"
		self.LinkedIN_LOGIN_PASSWORD = "$ystem@dmin97"
		self.INPUT_BASE_ID = 'appjvhsxUUz6o0dzo'
		self.OUTPUT_BASE_ID = 'appQfs70fHCsFgeUe'
		self.API_KEY = 'patQIAmVOLuXelY42.df469e641a30f1e69d29195be1c1b1362c9416fffc0ac17fd3e1a0b49be8b961'
		self.CompanyTable = 'tbl6d9xMvwRKcTlfY'
		self.Prospectus_Table = 'tblf4Ed9PaDo76QHH'
		self.GeoCitiesTable = 'tbl4PsNMGFGC4BRyE'
		self.OUTPUT_Table = 'tbli5Waff0LBrM5jU'
		self.WebscraperBase_OpenJob = "tblFx6SBmtNRCeOgm"
		self.headers = {'Authorization': 'Bearer '+ self.API_KEY}
		self.Post_Header = {'Authorization': 'Bearer '+ self.API_KEY,'Content-Type': 'application/json'}
		self.geoTableIds = {}    
		self.AllRecordIds = []
		self.social_info = None

	def update_crm(self,json_update_data,record_data):
		json_update_data = json.dumps(json_update_data)
		r = requests.patch("https://api.airtable.com/v0/"+self.INPUT_BASE_ID+"/"+self.Prospectus_Table+"/"+ record_data,data = json_update_data, headers=self.Post_Header)
		return r.text,r.status_code

	def getInputCompanyTable(self):
		offset = ''
		while 1:
			CompanyTableURL = 'https://api.airtable.com/v0/'+self.INPUT_BASE_ID +'/'+ self.Prospectus_Table 
			if len(self.VIEW_NAME) > 1:
				OutputTable = requests.get(CompanyTableURL, headers=self.headers,params={'offset': offset,'view':self.VIEW_NAME}).json()
			else:
				OutputTable = requests.get(CompanyTableURL, headers=self.headers,params={'offset': offset}).json()	
			for Records in OutputTable["records"]:
				for recordsKey,recordsValue in Records.items():
					if recordsKey == "fields":
						SingleRecord = {}
						CityCountry = []
						try:
							SingleRecord["Company"] = recordsValue["Company Name"]
						except:
							continue	
						try:
							int(recordsValue['LinkedIn ID'].replace("5B","").replace("5D","").replace('"',""))
							SingleRecord["CompanyId"] = recordsValue['LinkedIn ID'].replace("5B","").replace("5D","").replace('"',"")
						except:
							print(" "*4,recordsValue["Company Name"]+"[Not Found]")
							continue
						print(" "*4,recordsValue["Company Name"]+"["+SingleRecord["CompanyId"]+"]")	
						try:
							for citytoScrap in recordsValue['HQ Scrape']: 	
								CityCountry.append(citytoScrap+";HQ EEs")
						except:
							()
						try:		
							for citytoScrap in recordsValue['US Scrape']:
								CityCountry.append(citytoScrap+";US EEs")
						except:
							()
						try:
							for citytoScrap in recordsValue['Other US Cities To Scrape']:
								CityCountry.append(citytoScrap+";Other US Cities")
						except:
							()
						try:
							for citytoScrap in recordsValue['Countries to Scape']:
								CityCountry.append(citytoScrap+";Other Countries")
						except:
							()	
						SingleRecord["CityCountryToScrap"] = CityCountry
						self.AllRecordIds.append(SingleRecord)	
			try:
				nextOffset = OutputTable["offset"]
				offset = nextOffset
			except:
				break
		print(len(self.AllRecordIds))

	def GeoLocationIds(self):
		offset = ""
		while 1:
			geoTableUrl = 'https://api.airtable.com/v0/' + self.INPUT_BASE_ID + "/" + self.GeoCitiesTable
			r = requests.get(geoTableUrl, headers=self.headers,params={'offset': offset}).json()
			for Records in r["records"]:
				try:
					Records["fields"]["geoUrn"]
					locationName = Records["fields"]["Name"].replace("\n","").strip() +"|"+ Records["id"].replace("\n","").strip()
					print(" "*4,Records["fields"]["Name"].replace("\n","").strip()+"["+Records["fields"]["geoUrn"].replace("\n","").strip()+"]")
					locationGeoId = Records["fields"]["geoUrn"].replace("\n","").strip()
					self.geoTableIds[locationName] = locationGeoId
				except:
					try:
						locationName = Records["fields"]["Name"] +"|"+ Records["id"]
						self.geoTableIds[locationName] = "NULL"
						print(" "*4,Records["fields"]["Name"]+"[Not Found]")
					except:	
						()
			try:
				nextOffset = r["offset"]
				offset = nextOffset
			except:
				break			


	def Get_ChromeDriver(self):
		chrome_options = webdriver.ChromeOptions()
		chrome_options.add_argument("--start-maximized")
		# chrome_options.add_argument('--headless')
		# chrome_options.add_argument("--disable-gpu")
		chrome_options.add_argument("--no-sandbox")
		chrome_options.add_argument('--log-level=3')
		chrome_options.add_argument("--disable-notifications")
		chrome_options.add_argument('--ignore-certificate-errors-spki-list')
		chrome_options.add_argument('--ignore-ssl-errors')
		chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
		chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
		chrome_options.add_experimental_option('useAutomationExtension', False)	
		#os.environ['WDM_LOG_LEVEL'] = '0'
		prefs = {
					"credentials_enable_service": False, 
					"profile.password_manager_enabled": False , 
					"profile.default_content_setting_values.geolocation": 2,
					#"profile.managed_default_content_settings.images": 2
				}
		chrome_options.add_experimental_option("prefs", prefs)
		driver = webdriver.Chrome(options=chrome_options)
		#driver.delete_all_cookies()
		return driver


	def Login_LinkedIn(self,driver):
		print("load linkedin cookies...")
		with open("linkedin-cookies.json","r") as co:
			cookies = json.load(co)
		driver.get("https://www.linkedin.com/")
		for kuki in cookies:
			driver.add_cookie(kuki)
		driver.refresh()
		print(" "*2,"Login Successfull")
		return driver

	def convalue(self,val):
		if 'k' in val.lower():
			return int(float(val.lower().replace('k', '')) * 1_000)
		elif 'm' in val.lower():
			return int(float(val.lower().replace('m', '')) * 1_000_000)
		else:
			if val.isdigit():
				return int(val)
			else:
				return 0
	def scrapData(self,driver):
		gsht = GSheet(excelFile="LinkedInCompanyInfo.xlsx") # match the sheet name in your google sheet, 
		services = build("sheets", "v4", credentials=gsht.creds)
		gsht.get_sheet(service=services)
		for Records in self.AllRecordIds:
			#------------------------------------------------- Only Company Scrap --------------------------------------------------------------------
			this_CompanyId = Records["CompanyId"].replace('"','')
			driver.get("https://www.linkedin.com/search/results/people/?currentCompany=["+this_CompanyId+"]&origin=COMPANY_PAGE_CANNED_SEARCH&sid=ZJJ")		
			try:
				WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "search-results-container")))
			except:
				continue
			TotalResults = driver.find_element(By.CLASS_NAME,"search-results-container").find_element(By.TAG_NAME,"h2").text.split("result")[0]
			print("\n"+Records["Company"]+"["+this_CompanyId+"]")
			print(" "*3,"Total Employees: ",TotalResults,"Results")
			if "No" in TotalResults:
					TotalResults = "0"
			try:
				int(TotalResults)
			except:
				TotalResults = "0"	

			#-------------------------------------------------------------------------------------------------------------------------------------------------
			TotalEEs = 0 
			USEEs = 0
			HQEEs = 0
			OtherUSCities = ""
			OtherCountries = ""

			TotalEEs = int(TotalResults)
			#print(this_CompanyId)
			driver,openJobCount,linkedinURL,companyDetailsfromFunction = self.scrapOpenJobPage(driver,this_CompanyId)
			#print("_"*30)
			openJobCount = int(openJobCount)

			print(" "*3,"Open Jobs: ",openJobCount)
			print(" "*3,"Followers: ",companyDetailsfromFunction["Followers"])
			print(" "*3,"Website: ",companyDetailsfromFunction["companyWebsite"])
			print(" "*3,"Department: ",companyDetailsfromFunction["Department"])

			for CityCountry in Records["CityCountryToScrap"]:
				this_ProspectGeo = CityCountry.split(";")[1].replace('"',"")
				CityCountry = CityCountry.split(";")[0]
				#----------------------------------------------------------------------------------------
				LocationFound = False
				toPrintLocation = ""
				for location,GeoId_ in self.geoTableIds.items():
					if location.split('|')[0] == CityCountry or location.split('|')[1] == CityCountry:
						this_CompanyLocationID = GeoId_.replace('"','')
						this_CompanyLocationName = location.split('|')[0]
						if GeoId_ == "NULL":
							toPrintLocation = this_CompanyLocationName
							break
						LocationFound = True
						break
				if LocationFound == False:
					print(" "*3,"*GeoId for",toPrintLocation+"["+CityCountry+"] Not Found")
					continue
				#-----------------------------------------------------------------------------------------		
				driver.get("https://www.linkedin.com/search/results/people/?currentCompany=["+this_CompanyId+"]&geoUrn=["+this_CompanyLocationID+"]&origin=FACETED_SEARCH&sid=9vr")		
				try:
					WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "search-results-container")))
				except:
					continue
				TotalResults = driver.find_element(By.CLASS_NAME,"search-results-container").find_element(By.TAG_NAME,"h2").text.split("result")[0]
				if "No" in TotalResults:
					TotalResults = "0"
				try:
					int(TotalResults)
				except:
					TotalResults = "0"		
				print(" "*3,this_CompanyLocationName,"["+this_CompanyLocationID+"] :",str(TotalResults),"Results")
				time.sleep(1)

				if this_ProspectGeo == "HQ EEs":
					HQEEs = int(TotalResults)
				if this_ProspectGeo == "US EEs":
					USEEs = int(TotalResults)
				if this_ProspectGeo == "Other US Cities":
					OtherUSCities = OtherUSCities + "" + this_CompanyLocationName + " (" + str(TotalResults) + "),"
				if this_ProspectGeo == "Other Countries":
					OtherCountries = OtherCountries + " " + this_CompanyLocationName + " (" + str(TotalResults) + "),"

			OtherUSCities = sorted(OtherUSCities.split(","),reverse=True)
			OtherUSCities = ', '.join(OtherUSCities)
			OtherCountries = sorted(OtherCountries.split(","),reverse=True)
			OtherCountries = ', '.join(OtherCountries)
			RecordIdURL = "https://api.airtable.com/v0/"+self.INPUT_BASE_ID+"/"+self.Prospectus_Table+"?filterByFormula={LinkedIn ID}='"+str(this_CompanyId)+"'"
			time.sleep(1)	
			try:
				RecordIDToUpdateData = requests.get(RecordIdURL, headers=self.headers).json()["records"][0]["id"]
			except:
				print(" "*5,"["+str(this_CompanyId)+"]-->Id Not Found in Prospects Table")	
				continue
			
			gsheet_update = {
				"Company":Records["Company"],
				"LinkedIn URL":companyDetailsfromFunction['linkedinUrl'],
				"Logo": f"Logo_.jpg ({companyDetailsfromFunction['Company Logo']})",
				"Year Founded (Scraped)":companyDetailsfromFunction["yearFounded"],
				"LinkedIn Description (Scraped)":companyDetailsfromFunction["Short Description"],
				"LinkedIn Followers (Scraped)":self.convalue(companyDetailsfromFunction["Followers"]),
				"Total EEs (Scraped)":TotalEEs,
				"HQ City (Scraped)":companyDetailsfromFunction["headQuarter"],
				"Website (Scraped)":companyDetailsfromFunction["companyWebsite"],
				"Industry (Scraped)":companyDetailsfromFunction["Department"]
			}
			print(gsheet_update)
			gsht.update_locally(service=services,cell_updated=gsheet_update)

			crm_update_data = { 
									"fields": { 
										"Total EEs (Scraped)" :TotalEEs,
										"US EEs (Scraped)":USEEs,
										"HQ EEs (Scraped)":HQEEs,
										"Other US Cities (Scraped)":OtherUSCities.strip().strip(',').strip(),
										"Other Countries (Scraped)":OtherCountries.strip().strip(',').strip(),
										"Open Jobs (Scraped)":openJobCount,
										"Website (Delete)" : companyDetailsfromFunction["companyWebsite"],
										"LinkedIn Description (Scraped)":companyDetailsfromFunction["Short Description"],
										"Industry (Scraped)":companyDetailsfromFunction["Department"],
										"LinkedIn Followers (Scraped)":self.convalue(companyDetailsfromFunction["Followers"]),
										"Year Founded (Scraped)":int(companyDetailsfromFunction["yearFounded"])
										# "Logo (From Companies": [{"url":companyDetailsfromFunction["Company Logo"],"filename":this_CompanyId+"_Logo_.jpg"}]
										#"Create a Field for HeadQuarter":companyDetailsfromFunction["headQuarter"]
									}
								}
			print(self.update_crm(crm_update_data,record_data=RecordIDToUpdateData))
		return driver

	def scrapOpenJobPage(self,driver,this_CompanyId):
		companyDetails = {}
		companyDetails["Department"] = ""
		companyDetails["headQuarter"] = ""
		companyDetails["Followers"] = "0"
		companyDetails["Company Logo"] = ""
		companyDetails["Short Description"] = ""
		companyDetails["companyWebsite"] = ""	
		companyDetails["yearFounded"] = "0"
		companyDetails["linkedinUrl"] = ""
			
		driver.get("https://www.linkedin.com/company/"+this_CompanyId+"/jobs/")
		print(driver.current_url)
		companyDetails["linkedinUrl"] = driver.current_url

		try:
			jobOpeningtag = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CLASS_NAME, "org-jobs-job-search-form-module__headline"))).text
			# jobOpeningtag = driver.find_element(By.CLASS_NAME,"org-jobs-job-search-form-module__headline").text
			jobOpeningCount = jobOpeningtag.split("has ")[1].split(" ")[0].replace(",","").strip()
		except:
			jobOpeningCount = "0"
		linkedinURL = str(driver.current_url).replace("/jobs/","")
		
		#Section Company Top Details
		try:
			WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.CLASS_NAME, "block.mt2")))
		except:
			return driver,jobOpeningCount,linkedinURL,companyDetails
		
		try:
			shortDescription = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'org-top-card-summary__tagline'))).text
		except:
			shortDescription = ""
		
		try:
			CompaDetaillist = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "org-top-card-summary-info-list__info-item")))
			tempCompa = {}
			for index,detail in enumerate(CompaDetaillist):
				tempCompa["Department"] = 0
				tempCompa["headQuarter"] = 1
				tempCompa["Followers"] = 2
				if "followers" in detail.text:
					companyDetails["Followers"] = detail.text.split(" followers")[0]
				if "employees" in detail.text:
					if companyDetails["Followers"] == "0" or companyDetails["Followers"] ==2:
						companyDetails[list(tempCompa.keys())[list(tempCompa.values()).index(index)]] = ""
				if "followers" not in detail.text and "employees" not in detail.text and index==0:
					companyDetails[list(tempCompa.keys())[list(tempCompa.values()).index(index)]] = detail.text.split(",")[0]
		except Exception as e:
			print(f"error when scrape {this_CompanyId} company detail: ",e)	
		driver.get("https://www.linkedin.com/company/"+this_CompanyId+"/about/")		
		try:
			companyWebsite = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//dt[h3[text()="Website"]]'))).find_element(By.XPATH, 'following-sibling::dd[@class="mb4 t-black--light text-body-medium"]').text
			if "bit.ly" not in companyWebsite:
				companyWebsite = companyWebsite.split("?")[0].replace("//","|").split("/")[0].replace("|","//")
		except Exception as e:
			print(e)
			companyWebsite = ""
		try:
			companyLogoImageUrl = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'evi-image.lazy-image.ember-view.org-top-card-primary-content__logo'))).get_attribute("src")
		except:
			companyLogoImageUrl = ""		
		try:
			year_founded = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//dt[h3[text()="Founded"]]'))).find_element(By.XPATH, 'following-sibling::dd[@class="mb4 t-black--light text-body-medium"]').text.strip()
		except Exception as e:
			print(e)
			year_founded = "0"

		try:
			HQ = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//dt[h3[text()="Headquarters"]]'))).find_element(By.XPATH, 'following-sibling::dd[@class="mb4 t-black--light text-body-medium"]').text
		except Exception as e:
			print(e)
			HQ = ""

		print("data website: ",companyWebsite)
		print("data logo: ",companyLogoImageUrl)
		print("year founded: ",year_founded)
		companyDetails["Company Logo"] = companyLogoImageUrl
		companyDetails["Short Description"] = shortDescription
		companyDetails["companyWebsite"] = companyWebsite	
		companyDetails["yearFounded"] = year_founded	
		companyDetails["headQuarter"] = HQ

		return driver,jobOpeningCount,linkedinURL,companyDetails

if __name__ == "__main__":
	linkedin = Linkedin()
	print("Getting Companies to be Scraped:")
	linkedin.getInputCompanyTable()
	print("Scrapping GeoLocations:")	
	linkedin.GeoLocationIds()
	print("Start Chrome Driver Instance")
	driver = linkedin.Get_ChromeDriver()
	print("Login to LinkedIn")
	driver = linkedin.Login_LinkedIn(driver)
	print("Scrapping Employee Count")
	driver = linkedin.scrapData(driver)
	winsound.Beep(1500, 50)
	driver.quit()
