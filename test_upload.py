from seleniumbase import Basecase

def test_visible_upload(self):
    Brochure = ('C:/Users/Raghav/Desktop/Brochure')
    driver.find_element_by_css_selector('span[data-icon="clip"]').click()
    self.choose_file(driver.find_element_by_css_selector('span[data-icon="attach-image"]').click(), Brochure))
    driver.find_element_by_css_selector('span[data-icon="clip"]').click()
    driver.find_element_by_css_selector('span[data-icon="attach-image"]').click()
    driver.find_element_by_xpath("//input[@type='file']").send_keys("C://Users/Raghav/Desktop/Brochure")