from seleniumbase import Driver
from PIL import Image
import os, random, time
from loguru import logger
from abc import ABC, abstractmethod
from seleniumbase import page_actions as pa
from selenium.webdriver.support.ui import Select

class CreateFolder():
    @staticmethod
    def create_folder_if_not_exists(folder_path, name):
        cleaned_name = name.replace("/","").replace("\\","").replace("<","").replace(">","").replace("*","").replace("?", "").strip()
        full_folder_path = os.path.join(folder_path, cleaned_name)
        base_folder_path = os.path.basename(folder_path)

        if base_folder_path.lower() == cleaned_name.lower():
            pass
        elif os.path.exists(full_folder_path):
            pass
        elif not os.path.exists(full_folder_path):
            os.makedirs(full_folder_path)

class WebImageCropper(ABC):
    def __init__(self):
        pass

    @abstractmethod
    def crop(self):
        pass

class OFACImageCropper(WebImageCropper):
    def __init__(self, image_name, folder):
        self.image_name = image_name
        self.folder = folder

    def crop(self):
        with Image.open(f"{self.folder}\\{self.image_name}\\{self.image_name}_ofac.png") as img:
            width, height = img.size
            left = int(width * 0.24)  
            right = int(width * 0.752)  
            top = 0  + int(height * 0.05)
            bottom = height - int(height * 0.0445)
            cropped_img = img.crop((left, top, right, bottom))
            cropped_img.save(f"{self.folder}\\{self.image_name}\\{self.image_name}_ofac.png")

class OICImageCropperV1(WebImageCropper):
    def __init__(self, image_name, folder):
        self.image_name = image_name
        self.folder = folder

    def crop(self):
        if self.image_name[-1] == ".":
            self.image_name = self.image_name[:-1]
        with Image.open(f"{self.folder}\\{self.image_name}\\{self.image_name}_oic.png") as img:
            width, height = img.size
            left = int(width * 0.15)  
            right = int(width * 0.98)  
            top = 0  + int(height * 0.12)
            bottom = height
            cropped_img = img.crop((left, top, right, bottom))
            cropped_img.save(f"{self.folder}\\{self.image_name}\\{self.image_name}_oic.png")

class OICImageCropperV2(WebImageCropper):
    def __init__(self, image_name, folder):
        self.image_name = image_name
        self.folder = folder

    def crop(self):
        if self.image_name[-1] == ".":
            self.image_name = self.image_name[:-1]
        with Image.open(f"{self.folder}\\{self.image_name}\\{self.image_name}_oic.png") as img:
            width, height = img.size
            left = int(width * 0.15)  
            right = int(width * 0.99)  
            top = 0  + int(height * 0.12)
            bottom = height - (height*0.10)
            cropped_img = img.crop((left, top, right, bottom))
            cropped_img.save(f"{self.folder}\\{self.image_name}\\{self.image_name}_oic.png")



class WebAutoScreenshot(ABC):
    def __init__(self):
        pass

    @abstractmethod
    def check_url(self, driver):
        pass
    
    @abstractmethod
    def search_candidate(self, driver):
        pass

    def check_driver(self, driver):
        # global driver
        if not driver:
            driver = Driver(uc=True, d_height=1080, d_width=1920, headless=True)

class OFACAutoScreenshot(WebAutoScreenshot):

    def __init__(self, candidate_name, folder, timeout):
        self.timeout = timeout
        self.candidate_name = candidate_name
        self.url = "https://sanctionssearch.ofac.treas.gov/"
        self.clean_result = "0"
        self.folder = folder
    
    def check_url(self, driver, input_field):
        try:
            if not input_field:
                driver.get(self.url)
                logger.info(f"Navigating to {self.url}")
            input_field = pa.wait_for_element_visible(driver, 'input[name="ctl00$MainContent$txtLastName"]', timeout= self.timeout)
            return input_field
        except Exception as e:
            logger.error(f"Failed to navigate {self.url}")
            logger.error(f"Automation for OFAC stopped. Error: {str(e)}")
            raise e
            

    def search_candidate(self, driver, input_field):
        try:

            super().check_driver(driver)
            input_field = self.check_url(driver, input_field)
            input_field.clear()
            time.sleep(random.random())
            input_field.send_keys(self.candidate_name)
            time.sleep(random.randint(1,3))
            search_button = pa.wait_for_element_clickable(driver, 'input[type="submit"][value="Search"]', timeout= self.timeout)
            search_button.click()
            input_field = pa.wait_for_element_visible(driver,'input[name="ctl00$MainContent$txtLastName"]', timeout= self.timeout)
            result = pa.wait_for_element(driver, 'span#ctl00_MainContent_lblResults.fieldHeader').text[16]
            if result != self.clean_result:
                logger.warning(f"{self.candidate_name} has {result} result(s) when country is set to All.")
                country_dropdown = pa.wait_for_element_visible(driver, 'select#ctl00_MainContent_ddlCountry', timeout= self.timeout)
                Select(country_dropdown).select_by_value("Thailand")
                search_button = pa.wait_for_element_clickable(driver, 'input[type="submit"][value="Search"]', timeout= self.timeout)
                search_button.click()
                result = pa.wait_for_element(driver, 'span#ctl00_MainContent_lblResults.fieldHeader').text[16]
                if result != self.clean_result:
                    logger.warning(f"{self.candidate_name} has {result} result(s) when country is set to Thailand")
                if self.candidate_name[-1] == ".":
                    self.candidate_name = self.candidate_name[:-1]
                pa.save_screenshot(driver, f"{self.folder}\\{self.candidate_name}\\{self.candidate_name}_ofac.png")
                time.sleep(random.random())
                country_dropdown = pa.wait_for_element_visible(driver, 'select#ctl00_MainContent_ddlCountry', timeout= self.timeout)
                Select(country_dropdown).select_by_visible_text("All")
                input_field = pa.wait_for_element_visible(driver, 'input[name="ctl00$MainContent$txtLastName"]', timeout= self.timeout)
            else:
                if self.candidate_name[-1] == ".":
                    self.candidate_name = self.candidate_name[:-1]
                    pa.save_screenshot(driver, f"{self.folder}\\{self.candidate_name}\\{self.candidate_name}_ofac.png")
                    input_field = pa.wait_for_element_visible(driver, 'input[name="ctl00$MainContent$txtLastName"]', timeout= self.timeout)
                else:
                    pa.save_screenshot(driver, f"{self.folder}\\{self.candidate_name}\\{self.candidate_name}_ofac.png")
                    input_field = pa.wait_for_element_visible(driver, 'input[name="ctl00$MainContent$txtLastName"]', timeout= self.timeout)
            OFACImageCropper(self.candidate_name, self.folder).crop()
            logger.success(f"Screenshot for {self.candidate_name} is done.")
            return input_field
        except (Exception) as e:
            logger.error(f"Automation for OFAC has failed. Error: {e}")
            raise e

class OICAutoScreenshotV1(WebAutoScreenshot):
    def __init__(self, first_name, last_name, folder, timeout):
        self.timeout = timeout
        self.first_name = first_name
        self.last_name = last_name
        self.folder = folder
        self.url_individual = "https://smart.oic.or.th/EService/Menu1"
        self.candidate_name = (str(self.first_name) + " " + str(self.last_name)).strip()
        self.clean_result = "0"

    def check_url(self, driver, input_first_name, input_last_name):
        try:
            if (input_first_name == None and input_last_name == None):
                driver.get(self.url_individual)
                logger.info(f"Navigating to {self.url_individual}")
            if driver.current_url != self.url_individual:
                driver.get(self.url_individual)
            input_first_name = pa.wait_for_element_present(driver, 'input#b15-Input_FirstName.form-control.OSFillParent', timeout=self.timeout)
            input_last_name = pa.wait_for_element_present(driver, 'input#b15-Input_LastName.form-control.OSFillParent', timeout=self.timeout)
            return input_first_name, input_last_name
        except Exception as e:
            logger.error(f"Failed to navigate {self.url_individual}")
            logger.error(f"Automation for OIC failed. Error: {str(e)}")
            raise e

    def search_candidate(self, driver, input_first_name, input_last_name):
        try:
            super().check_driver(driver)
            input_first_name, input_last_name = self.check_url(driver, input_first_name, input_last_name)
            input_first_name.clear()
            input_last_name.clear()
            time.sleep(random.random())
            input_first_name.send_keys(self.first_name)
            input_last_name.send_keys(self.last_name)
            time.sleep(random.randint(1,3))
            search_button = pa.wait_for_element_present(driver, 'button.btn.background-secondary', timeout= self.timeout)
            # search_button.click()
            driver.execute_script("arguments[0].click()", search_button)
            input_first_name = pa.wait_for_element_present(driver,'input#b15-Input_FirstName.form-control.OSFillParent', timeout= self.timeout)
            input_last_name = pa.wait_for_element_present(driver, 'input#b15-Input_LastName.form-control.OSFillParent', timeout= self.timeout)
            pa.wait_for_element_present(driver, 'div#b15-b14-Content', timeout=self.timeout)
            scale_ratio = 0.9
            driver.execute_script(f"document.body.style.zoom='{scale_ratio}'")
            time.sleep(1)
            if self.candidate_name[-1] == ".":
                self.candidate_name = self.candidate_name[:-1]
                pa.save_screenshot(driver, f"{self.folder}\\{self.candidate_name}\\{self.candidate_name}_oic.png")
            else:
                pa.save_screenshot(driver, f"{self.folder}\\{self.candidate_name}\\{self.candidate_name}_oic.png")
            OICImageCropperV1(self.candidate_name, self.folder).crop()
            logger.success(f"Screenshot for {self.candidate_name} is done.")
            return input_first_name, input_last_name
        except (Exception) as e:
            logger.error(f"Automation for OIC failed. Error: {e}")
            raise e

class OICAutoScreenshotV2(WebAutoScreenshot):
    def __init__(self, company_name, folder, timeout):
        self.timeout = timeout
        self.company_name = company_name
        self.folder = folder
        self.url_company = "https://smart.oic.or.th/EService/Menu2"
        self.clean_result = "0"

    def check_url(self, driver, input_company_name):
        try:
            if not input_company_name:
                driver.get(self.url_company)
                logger.info(f"Navigating to {self.url_company}")
            if driver.current_url != self.url_company:
                driver.get(self.url_company)
            input_company_name = pa.wait_for_element_visible(driver, 'input#b13-Input_TextVar.form-control.OSFillParent', timeout=self.timeout)
            return input_company_name
        except Exception as e:
                logger.error(f"Failed to navigate {self.url_company}")
                logger.error(f"Automation for OIC failed. Error: {str(e)}")
                raise e

    def search_candidate(self,driver, input_company_name):
        try:    
            super().check_driver(driver)
            input_company_name = self.check_url(driver, input_company_name)
            input_company_name.clear()
            time.sleep(random.random())
            input_company_name.send_keys(self.company_name)
            time.sleep(random.randint(1,3))
            search_button = pa.wait_for_element_clickable(driver, 'button.btn.background-secondary', timeout=self.timeout)
            # search_button.click()
            driver.execute_script("arguments[0].click()", search_button)
            input_company_name = pa.wait_for_element_visible(driver, 'input#b13-Input_TextVar.form-control.OSFillParent', timeout=self.timeout)
            scale_ratio = 0.9
            driver.execute_script(f"document.body.style.zoom='{scale_ratio}'")
            time.sleep(1)
            if self.company_name[-1] == ".":
                self.company_name = self.company_name[:-1]
                pa.save_screenshot(driver, f"{self.folder}\\{self.company_name}\\{self.company_name}_oic.png")
            else:
                pa.save_screenshot(driver, f"{self.folder}\\{self.company_name}\\{self.company_name}_oic.png")
            OICImageCropperV2(self.company_name, self.folder).crop()
            logger.success(f"Screenshot for {self.company_name} is done.")
            return input_company_name
        except (Exception) as e:
            logger.error(f"Automation for OIC failed. Error: {e}")
            raise e