import re,requests,urllib.parse,json,os
from seleniumbase import SB
from dotenv import load_dotenv
load_dotenv()

class LookUp:
	def __init__(self):
		self.sb = None
		self.VIEW_NAME = "ATX Ventures" # Enter VIEW_NAME here. OR leave it empty.
		self.GLASSDOOR_LOGIN_EMAIL = os.getenv("Glassdoor_Email")
		self.GLASSDOOR_PASSWORD = os.getenv("Glassdoor_Pass")
		self.INPUT_BASE_ID = os.getenv("INPUT_BASE_ID")
		self.API_KEY = os.getenv("API_KEY")
		self.Prospectus_Table = os.getenv("Prospectus_Table")
		self.CRM_BASE_ID = os.getenv("INPUT_BASE_ID")
		self.CRM_BASE_Prospectus_Tabke  = os.getenv("Prospectus_Table")
		self.headers = {'Authorization': 'Bearer '+ self.API_KEY}
		self.Post_Header = {'Authorization': 'Bearer '+ self.API_KEY,'Content-Type': 'application/json'}
		self.AllRecordIds = []
		self.GDrecords = []
		

	def updateCrm(self,company,gdurl):
		print("function calling",gdurl)
		CompanyRecordIDURL = "https://api.airtable.com/v0/" + self.CRM_BASE_ID + "/" + self.CRM_BASE_Prospectus_Tabke + "?filterByFormula=%7BCompany+Name%7D%3D%27" + urllib.parse.quote_plus(company) + "%27"
		response = requests.request("GET", CompanyRecordIDURL, headers=self.headers).json()
		if list(response.keys())[0]!="error":
			print(response)
			CompanyRecordID = response["records"][0]["id"]		
			# data = {"fields": {"Glassdoor URL": gdurl,}}

			json_data = json.dumps({"fields": {"Glassdoor URL": gdurl}})
			response = requests.request("PATCH", f"https://api.airtable.com/v0/{self.CRM_BASE_ID}/{self.CRM_BASE_Prospectus_Tabke}/{CompanyRecordID}", headers=self.Post_Header, data=json_data)
			print(" " * 9, "CRM PATHCING STATUS: ", response.status_code)
			print(" " * 9, f"UPDATING {company} TO CRM DONE!")
			if response.status_code != 200:
				print("Error patching Airtable record:", response.json())
		else:
			print(f"{company} failed to update crm")

	def getInputCompanyTable(self):
		offset = ''
		while 1:
			CompanyTableURL = 'https://api.airtable.com/v0/' + self.INPUT_BASE_ID + '/' + self.Prospectus_Table
			if len(self.VIEW_NAME) > 1:
				OutputTable = requests.get(CompanyTableURL, headers=self.headers, params={'offset': offset, 'view': self.VIEW_NAME}).json()
			else:
				OutputTable = requests.get(CompanyTableURL, headers=self.headers, params={'offset': offset}).json()
			for Records in OutputTable["records"]:
				for recordsKey, recordsValue in Records.items():
					if recordsKey == "fields":
						SingleRecord = {}
						try:
							if "Glassdoor URL" not in list(recordsValue.keys()) and "Website (from Companies)" in list(recordsValue.keys()):
								SingleRecord = {
									"Company Name": recordsValue['Company Name'],
									"Website": recordsValue["Website (from Companies)"],
									"recId": Records['id']
								}
								print(SingleRecord)
								self.AllRecordIds.append(SingleRecord)
						except Exception as e:
							print(e)
							print(f"{recordsValue['Company Name']} already exist!")
			try:
				nextOffset = OutputTable["offset"]
				offset = nextOffset
			except KeyError:
				break

	def login_glassdoor(self):
		self.sb.uc_open_with_reconnect("https://www.glassdoor.com/index.htm", 3)
		for repeat in range(3): #repeat login for 3 times
			print(f"trying login for {repeat} times")
			try:
				self.sb.wait_for_element_visible("#inlineUserEmail",timeout=4)
				if self.sb.is_element_visible("#inlineUserEmail"):
					self.sb.type("#inlineUserEmail", self.GLASSDOOR_LOGIN_EMAIL + "\n", timeout=3, retry=2)
					self.sb.type("/html/body/div[3]/section[1]/div[2]/div/div/div[1]/div[1]/div/div/div/form/div[1]/div[1]/div/div[1]/input", self.GLASSDOOR_PASSWORD + "\n")
					break
			except:
				continue
			else:
				print("Login failed cause element not found!")
				continue
			
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

	def get_element(self,selector,by):
		try:
			self.sb.wait_for_element_visible(selector,by=by,timeout=3)
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

	def filterUrl(self,url,pat):
		match = re.search(pat,url)
		if match:
			return match.group(1)
	
	def search_company(self):
		print(self.AllRecordIds)
		for Records in self.AllRecordIds:

			self.sb.open(f"https://www.glassdoor.com/Reviews/company-reviews.htm?typedKeyword={Records['Company Name']}&context=Review&sc.keyword={Records['Company Name']}")			
			page = self.sb.get_current_url()
			print("searching ", Records['Company Name'])
			print("current url ", page)
			if "Explore" not in page:
				isfound = False
				isCompany = self.get_element('//*[@id="MainCol"]/div/header/h1',by="xpath")			
				if isCompany:
					if f'Showing results for ' in isCompany.text:
						paging = self.get_elements('//*[@class="page "]',by="xpath")
						pageIndex = 0
						if paging:
							for pagenum in range(1,len(paging)+4):
								pageIndex+=1
							print("Page Total: ",pagenum)
						else:
							pageIndex=1
						for pg in range(1,pageIndex+3):
							companyList = self.get_elements('//*[@class="single-company-result module "]',by="xpath")
							if companyList:
								for x in range(1,len(companyList)+2):
									gdurl = self.get_element(f'//*[@id="MainCol"]/div/div[{x}]/div/div[1]/div/div[2]/h2/a',by="xpath")					
									compurl = self.get_element(f'//*[@id="MainCol"]/div/div[{x}]/div/div[1]/div/div[2]/div/p[2]/span/a',by="xpath")					
									if gdurl and compurl:
										print(Records['Website'][0])
										print("company name: ",Records['Company Name'])
										if self.filterUrl(compurl.get_attribute("href"),pat=r'^(?:https?://)?(?:www\.)?([^/]+)') == self.filterUrl(Records['Website'][0],pat=r'^(?:https?://)?(?:www\.)?([^/]+)'):
											print(f"Found Glassdoor {Records['Company Name']} URL: ", gdurl.get_attribute("href"))
											self.updateCrm(company=Records['Company Name'],gdurl=gdurl.get_attribute("href"))
											self.GDrecords.append({"Company Name":Records['Company Name'],"GD URL":gdurl.get_attribute("href")})
											isfound = True
											break
								if isfound:
									break
								self.sb.open(f'{page.split(".htm")[0]}_IP{pg+1}.htm')	
								# print(f"{Records['Company Name']} Not Found!")							
						if isfound == False:
							print(f"{Records['Company Name']} not found")
							self.updateCrm(company=Records['Company Name'],gdurl="\n")
					else:
						print(f"{Records['Company Name']} not found")
						self.updateCrm(company=Records['Company Name'],gdurl="\n")
				else:
					if (self.filterUrl(url=page,pat=r'^(?:https?://)?(?:www\.)?[^/]+/([^/?]+)')) == "Overview":
						compurl = self.get_element(selector='//*[@id="__next"]/div/div[1]/div[2]/main/div/div/div[1]/div/ul/li[1]/a',by="xpath")
						if compurl:
							if self.filterUrl(compurl.get_attribute("href"),pat=r'^(?:https?://)?(?:www\.)?([^/]+)') == self.filterUrl(Records['Website'][0],pat=r'^(?:https?://)?(?:www\.)?([^/]+)'):
								print(f"Found Glassdoor {Records['Company Name']} URL: ", page)
								print("company name: ",Records['Company Name'])
								self.updateCrm(company=Records['Company Name'],gdurl=page)
								self.GDrecords.append({"Company Name":Records['Company Name'],"GD URL":page})
							else:
								print(f"{Records['Company Name']} not found")
								self.updateCrm(company=Records['Company Name'],gdurl="\n")
						else:
							print(f"{Records['Company Name']} not found")
							self.updateCrm(company=Records['Company Name'],gdurl="\n")
			else:
				self.updateCrm(company=Records['Company Name'],gdurl="\n")

	def Main(self):
		with SB(uc=True) as Sb:
			self.sb = Sb
			self.login_glassdoor()
			self.getInputCompanyTable()
			self.search_company()


if __name__ == "__main__":
	lkp = LookUp()
	lkp.Main()
