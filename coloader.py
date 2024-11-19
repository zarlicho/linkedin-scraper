from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
import json

class CookiesLoader:
    def __init__(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--auto-open-devtools-for-tabs")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument('--ignore-certificate-errors-spki-list')
        chrome_options.add_argument('--ignore-ssl-errors')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)	
        self.driver = webdriver.Chrome(options=chrome_options)

    def load_cookies(self,website):
        self.driver.get(website)
        if input("type done for close: ") == "done":
            print(f"geting linkedin cookies")
            cookies = self.driver.get_cookies()
            print(f"load linkedin cookies")
            with open(f"linkedin-cookies.json","w") as coo:
                json.dump(cookies,coo)
            print(f"done to load linkedin cookies!")   
            self.driver.close()         
            self.driver.quit()         

if __name__ == "__main__":
    loader = CookiesLoader()
    loader.load_cookies(website="https://www.linkedin.com/checkpoint/lg/sign-in-another-account?trk=guest_homepage-basic_nav-header-signin")