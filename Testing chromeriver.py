from selenium import webdriver

driver = webdriver.Chrome()
driver.get("https://www.sldcguj.com/Energy_Block_New.php")
print(driver.title)
driver.quit()
