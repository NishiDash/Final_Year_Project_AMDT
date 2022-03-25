from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

path = 'C:\Program Files (x86)\chromedriver.exe'
driver = webdriver.Chrome(path)

driver.get('http://127.0.0.1:5000/')
time.sleep(3)
################################### LOGIN #############################################
login = driver.find_element_by_xpath("//*[@id='navbarSupportedContent']/ul[2]/li[1]/a/span")
login.click()
time.sleep(2)
SSOID = driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div/form/input[1]')
SSOID.send_keys('223')
PWD = driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div/form/input[2]')
PWD.send_keys('12345678')
time.sleep(3)
driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div/form/input[3]').click()
################################### FILE MODIFICATION ##################################
time.sleep(3)
upload_file = driver.find_element_by_xpath('/html/body/div/div[2]/div/div/form/input[1]')
upload_file.send_keys("C:/Users/223053192/Desktop/work/Python/Final Year Project/AD.xlsx")
driver.find_element_by_xpath('/html/body/div/div[2]/div/div/form/input[2]').click()
time.sleep(13)
################################### Continue and exit ###################################
#driver.find_element_by_xpath('/html/body/div/div/div/a[1]').click()
driver.find_element_by_xpath('/html/body/div/div/div/a[2]').click()
################################### SIGNUP #############################################
time.sleep(1)
driver.find_element_by_xpath('//*[@id="navbarSupportedContent"]/ul[2]/li[2]/a/span').click()
time.sleep(2)
driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div/form/input[1]').send_keys('Nishita.Dash@ge.com')
driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div/form/input[2]').send_keys('223053193')
driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div/form/input[3]').send_keys('cse73626N')
time.sleep(3)
driver.find_element_by_xpath('/html/body/div/div/div[2]/div/div/form/input[4]').click()
################################### FILE MODIFICATION ##################################
time.sleep(3)
upload_file = driver.find_element_by_xpath('/html/body/div/div[2]/div/div/form/input[1]')
upload_file.send_keys("C:/Users/223053192/Desktop/work/Python/Final Year Project/model.xlsx")
driver.find_element_by_xpath('//*[@id="attribute"]').click()
driver.find_element_by_xpath('//*[@id="attribute"]/option[2]').click()

driver.find_element_by_xpath('/html/body/div/div[2]/div/div/form/input[2]').click()
time.sleep(13)
################################### LOGOUT #############################################

# time.sleep(4)
driver.find_element_by_xpath('//*[@id="navbarSupportedContent"]/ul[2]/li/a/span').click()

time.sleep(3)
driver.quit()


