from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

options = Options()

# gets rid of "Chrome is being controlled by automated test software." infobar
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# gets rid of notifications popup window
options.add_experimental_option("prefs", {
    "profile.default_content_setting_values.notifications": 1,
    "credentials_enable_service": False,
    "profile.password_manager_enabled": False
})

try:
    driver = webdriver.Chrome(options=options)
    driver.get('https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=13&ct=1612506865&rver=7.0.6737.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f%3fnlp%3d1%26cobrandid%3dab0455a0-8d03-46b9-b18b-df2f57b9e44c%26RpsCsrfState%3dd0bd3d5c-d539-deac-143a-39af53382784&id=292841&aadredir=1&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=ab0455a0-8d03-46b9-b18b-df2f57b9e44c')

    # code below corresponds to Outlook email
    email_input_field = driver.find_element_by_xpath('//input[contains(@type, "email")]')
    email_input_field.send_keys('your email')

    next_button = driver.find_element_by_xpath('//input[contains(@type, "submit")]')
    next_button.click()

    driver_wait = WebDriverWait(driver, 5)

    # sends keys to password input field
    driver_wait.until(EC.element_to_be_clickable((By.ID, "i0118"))).send_keys('your password')

    sign_in_button = driver.find_element_by_xpath('//input[contains(@id, "idSIButton9")]')
    sign_in_button.click()

    # clicks "New Message" button
    driver_wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "ms-Button-label"))).click()

    # sends keys to "To" email input field
    driver_wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "ms-BasePicker-input"))).send_keys('email to send message to')

    subject_field = driver.find_element_by_xpath('//input[contains(@class, "ms-TextField-field")]')
    subject_field.send_keys('Bot message')

    message_field = driver.find_element_by_xpath('//div[contains(@class, "_4utP_vaqQ3UQZH0GEBVQe")]')
    message_field.send_keys('This email was sent by a bot!')

    send_button = driver.find_element_by_xpath('//i[contains(@data-icon-name, "Send")]')
    send_button.click()
except Exception as error:
    print(f'{error}')
finally:
    # time.sleep() is used to keep browser open long
    # enough to see that the code worked properly
    time.sleep(20)
    driver.quit()
