import time
from selenium import webdriver
from openpyxl import Workbook
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from pynput.keyboard import Key, Controller

filepath = "C:/Users/Raghav/Desktop/contact_list.xlsx"
wb = load_workbook(filepath)
sheet = wb.active
frequency = 3
i = 1
# open chrome
driver = webdriver.Chrome('C:\chromedriver')
## maximize window
#driver.maximize_window()
# open whatsapp homepage for scanning
driver.get("https://web.whatsapp.com")
# wait for some time(in sec)
time.sleep(22)
for i in range(frequency):
    b3 = sheet.cell(row=i + 1, column=2)
    url = "https://wa.me/+91"
    i + 1
    print(url + str(b3.value))
    # open new window
    driver.execute_script("window.open('');")
    # Switch to the new window
    driver.switch_to.window(driver.window_handles[1])
    # open the link
    driver.get(url + str(b3.value))
    # wait for some time(in sec)
    time.sleep(2)
    ## continue_to_chat = driver.find_element(By.xpath("//input[@id='_9rne _9vcv _9u4i _9scb']"))
    continue_to_chat = driver.find_element_by_id("action-button")
    continue_to_chat.click()
    # driver.find_element_by_text('Continue to Chat')
    continue_to_chat.click()
    time.sleep(2)
    use_whatsapp_web = driver.find_element_by_xpath("// a[contains(text(),'use WhatsApp Web')]")
    # use_whatsapp_web = driver.find_element_by_text('use WhatsApp Web')
    # use_whatsapp_web = driver.find_element(By.xpath("//input[@class='_9rne _9vcv _9vcx']")) not possible as same as download button
    use_whatsapp_web.click()
    time.sleep(25)
    # message = driver.find_element(By.xpath("//input[@class='_13NKt copyable-text selectable-text']"))
    # message = driver.find_element_by_text('Type a message')
    message = driver.find_element_by_xpath('//*[@title="Type a message"]')
    # type message
    message.send_keys("Hello")
    keyboard = Controller()
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)
    time.sleep(15)

# Close the browser
driver.quit()
