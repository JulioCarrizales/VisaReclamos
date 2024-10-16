# Generated by Selenium IDE
# DOWNLOAD >>> pip install selenium pytest <<< DOWNLOAD
import pytest
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities


class TestUntitled():
  def setup_method(self, method):
    self.driver = webdriver.Chrome()
    self.vars = {}
  
  def teardown_method(self, method):
    self.driver.quit()
  
  def wait_for_window(self, timeout = 2):
    time.sleep(round(timeout / 1000))
    wh_now = self.driver.window_handles
    wh_then = self.vars["window_handles"]
    if len(wh_now) > len(wh_then):
      return set(wh_now).difference(set(wh_then)).pop()
  
  def test_untitled(self):
    # Test name: Untitled
    # Step # | name | target | value | comment
    # 1 | open | /home |  | 
    self.driver.get("https://secure.visaonline.com/home")
    # 2 | setWindowSize | 559x760 |  | 
    self.driver.set_window_size(559, 760)
    # 3 | mouseOver | id=vol-megamenu-toggle-1 |  | 
    element = self.driver.find_element(By.ID, "vol-megamenu-toggle-1")
    actions = ActionChains(self.driver)
    actions.move_to_element(element).perform()
    # 4 | click | id=ROL |  | 
    self.vars["window_handles"] = self.driver.window_handles
    # 5 | selectWindow | handle=${win562} |  | 
    self.driver.find_element(By.ID, "ROL").click()
    # 6 | mouseOver | id=inquirymenua |  | 
    self.vars["win562"] = self.wait_for_window(2000)
    # 7 | mouseOut | id=inquirymenua |  | 
    self.driver.switch_to.window(self.vars["win562"])
    # 8 | click | linkText=Transaction Inquiry |  | 
    element = self.driver.find_element(By.ID, "inquirymenua")
    actions = ActionChains(self.driver)
    actions.move_to_element(element).perform()
    element = self.driver.find_element(By.CSS_SELECTOR, "body")
    actions = ActionChains(self.driver)
    actions.move_to_element(element, 0, 0).perform()
    self.driver.find_element(By.LINK_TEXT, "Transaction Inquiry").click()
  
  