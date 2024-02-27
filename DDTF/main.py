from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service

from Locators.locators import Web_Locators

from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.action_chains import ActionChains

from excel_function.excel_fn import Excel_Fn

excel_file = Excel_Fn("testdata.xlsx", "Sheet1")
rows = excel_file.row_count()

driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()))
driver.implicitly_wait(10)
driver.maximize_window()
driver.get(Web_Locators().url)

for row in range(2, rows + 1):
    username = excel_file.read_data(row, 6)
    password = excel_file.read_data(row, 7)

    driver.find_element(by=By.NAME, value=Web_Locators().username_locator).send_keys(username)
    driver.find_element(by=By.NAME, value=Web_Locators().password_locator).send_keys(password)
    driver.find_element(by=By.XPATH, value=Web_Locators().submit_button).click()

    driver.implicitly_wait(10)

    if Web_Locators().dashboard_url in driver.current_url:
        print("SUCCESS : Login with Username {a}".format(a=username))
        excel_file.write_data(row, 8, "TEST PASSED")
        action = ActionChains(driver)
        logout_button = driver.find_element(by=By.XPATH, value=Web_Locators().logout_button)
        action.click(on_element=logout_button).perform()
        driver.find_element(by=By.LINK_TEXT, value="Logout").click()
    elif Web_Locators().url in driver.current_url:
        print("FAIL : Login failure with Username {a}".format(a=username))
        excel_file.write_data(row, 8, "TEST FAIL")
        driver.refresh()

driver.quit()
