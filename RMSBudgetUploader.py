import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from enum import Enum
import getpass
import os


class BudgetCategory(Enum):
    PERSONNEL = "Personnel"
    TEACHING_RELIEF = "Teaching Relief"
    TRAVEL = "Travel"
    FIELD_RESEARCH = "Field Research"
    EQUIPMENT = "Equipment"
    MAINTENANCE = "Maintenance"
    OTHER = "Other"
    TOTAL = "Total"

class RMSBudgetBuilder():
  def __init__(self, options):
    self.driver = None
    self.vars = {}
    self.options = options


  def __init__(self):
    options = Options()
    options.binary_location = ("/Applications/Firefox Developer Edition.app/Contents/MacOS/firefox")
    #options.headless = True
    options.add_argument("--width=1920")
    options.add_argument("--height=1080")
    self.driver = None
    self.vars = {}
    self.options = options
    

  def setup(self):
    self.driver = webdriver.Firefox(options=self.options)
    self.vars = {}
  
  def __del__(self):
    if self.driver is not None:
      self.driver.quit()
  
  def get_credentials(self):
    credentials = None
    # check if the credentials file exists
    if os.path.exists("credentials.json"):
      with open("credentials.json", "r") as file:
        credentials = json.load(file)
        return credentials
    else:
      # prompt the user for username and password
      username = input("Enter your RMS Email: ")
      password = getpass.getpass("Enter your password: ")

      # save the username and password to a file
      credentials = {
        "username": username,
        "password": password
      }
      with open("credentials.json", "w") as file:
        json.dump(credentials, file)

      # make the file only readable by the user
      os.chmod("credentials.json", 0o600)

      return credentials

  def login(self):
    credentials = self.get_credentials()
    
    self.driver.get("https://rms.arc.gov.au/RMS/ActionCentre/Account/Login?ReturnUrl=%2FRMS%2FActionCentre")
    self.driver.set_window_size(550, 691)
    self.driver.find_element(By.ID, "emailAddress").send_keys(credentials["username"])
    self.driver.find_element(By.ID, "password").click()
    self.driver.find_element(By.ID, "password").send_keys(credentials["password"])
    self.driver.find_element(By.ID, "login").click()

  def goto_budget(self):
    # wait until the page is loaded
    WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_element_located((By.LINK_TEXT, "Edit")))
    element = self.driver.find_element(By.LINK_TEXT, "Edit")
    self.driver.execute_script("arguments[0].click();", element)
    #element.click()
    #.click()
    # wait until the page is loaded with the -delta-form-part-button-60c7008b-bde5-443a-976c-290c3eaa6de8__b9af8d46-0fab-4ece-bed1-d76af1982525 element
    WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_element_located((By.ID, "-delta-form-part-button-60c7008b-bde5-443a-976c-290c3eaa6de8__b9af8d46-0fab-4ece-bed1-d76af1982525"))
    )
    element = self.driver.find_element(By.ID, "-delta-form-part-button-60c7008b-bde5-443a-976c-290c3eaa6de8__b9af8d46-0fab-4ece-bed1-d76af1982525")
    self.driver.execute_script("arguments[0].click();", element)

  def goto_budget_year(self, year):
    #self.driver.execute_script("window.scrollTo(0,0)")
    element = self.driver.find_element(By.ID, f"-delta-budget-year-{year}")
    #click it with js
    self.driver.execute_script("arguments[0].click();", element)

  def create_element(self, category, year, name):
      element = self.driver.find_element(By.CSS_SELECTOR, f'#year{year} .-delta-budget-line[data-name="{category.value}"] .bi')
      self.driver.execute_script("arguments[0].click();", element)
      wait = WebDriverWait(self.driver, 10)
      element = wait.until(expected_conditions.element_to_be_clickable((By.ID, "__bootbox_custom_input")))
      self.driver.find_element(By.ID, "__bootbox_custom_input").send_keys(name)
      self.driver.find_element(By.ID, "__bootbox_custom_input").send_keys(Keys.ENTER)
      self.driver.find_element(By.CSS_SELECTOR, ".btn-primary").click()
      # wait one second
      time.sleep(1)


  def input_category(self, year, category: BudgetCategory, name, arc_cash, admin_cash, admin_inkind):
    name = name.strip()

    last_element = None
    while last_element == None:

      # Fetch all elements with the specified data-parent attribute
      selector = f'#year{year} .-delta-budget-line[data-parent="{category.value}"]'
      elements = self.driver.find_elements(By.CSS_SELECTOR,selector)
      elements = [element for element in elements if element.value_of_css_property('display') != 'none']
      # Select the one whose text matches name and set it to last_element
      last_element = None
      for element in elements:
        if element.text == name:
          last_element = element
          break

      if last_element is None:
        self.create_element(category, year, name)

    #last_element = elements[-1]
    # Now find the child element with the specified headers attribute
    target_elements = last_element.find_elements(By.CSS_SELECTOR,'td > input')
    inputs = [arc_cash, admin_cash, admin_inkind]
    for element in target_elements[-3:]:
      # special case for HDR (Higher Degree by Research stipend) on RMS
      if name == 'HDR (Higher Degree by Research stipend)' and len(inputs) == 3:
        inputs.pop(0)
        continue
      self.driver.execute_script("arguments[0].click();", element)
      self.driver.execute_script("arguments[0].value = '';", element)
      element.send_keys(inputs.pop(0))

    # final click just to save the data
    self.driver.execute_script("arguments[0].click();", target_elements[-2])

  def save_budget(self):
    element = self.driver.find_element(By.ID, "-delta-form-save")
    self.driver.execute_script("arguments[0].click();", element)
  
  def apply_teaching_relief(self):
    # click all .-teaching-relief-approve-button
    # wait for document to finish loading
    WebDriverWait(self.driver, 10).until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR, ".-teaching-relief-approve-button")))
    time.sleep(1)

    while True:
      try:
        element = self.driver.find_element(By.CSS_SELECTOR, ".-teaching-relief-approve-button:not([disabled])")
        self.driver.execute_script("arguments[0].click();", element)
      except:
        break

  def test_build_rms_budget(self):
    self.setup()
    self.login()
    self.goto_budget()
    self.goto_budget_year(1)
    #self.input_category(1, RMSBudgetBuilder.BudgetCategory.PERSONNEL, "Person 1", "10000", "2000", "3000")
    #self.input_category(1, RMSBudgetBuilder.BudgetCategory.PERSONNEL , "Person 2", "20000", "3000", "4000")
    self.save_budget()
    self.__del__()



if __name__ == "__main__":
  options = Options()
  options.binary_location = ("/Applications/Firefox Developer Edition.app/Contents/MacOS/firefox")

  #options.headless = True
  options.add_argument("--width=1920")
  options.add_argument("--height=1080")
  test = RMSBudgetBuilder(options)
  test.test_build_rms_budget()