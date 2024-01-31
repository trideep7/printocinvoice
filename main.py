# Require PDFKit, Chromedriver and Selenium to be installed
import os
import pdfkit
from selenium import webdriver
from datetime import datetime, timedelta
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# Point Pdfkit configuration to htmltopdf
path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path)

# Set up the driver
website = 'https://www.yourwebsite.com/admin/'
path = 'C:\Chromedriver\chromedriver.exe'  # Link to your Chromedriver
service = Service(path)

with webdriver.Chrome(service=service) as driver:
    driver.maximize_window()

    # Login to the website
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "input-username")))
    username_field.send_keys("username")
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "input-password")))
    password_field.send_keys("password")
    login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
    login_button.click()

    # Navigate to Orders page
    sales_link = driver.find_element('id', 'menu-sale')
    sales_link.click()
    order_link = driver.find_element(By.XPATH, "//a[text()='Orders']")
    order_link.click()

    # Calculate the date of the last day of the previous month
    last_day_of_previous_month = datetime.now().replace(day=1) - timedelta(days=1)

    # Create folder for invoices
    folder_path = 'C:\\Invoices\\'
    previous_month_name = last_day_of_previous_month.strftime("%B").upper() + "-" + datetime.now().strftime("%Y")
    download_folder = os.path.join(folder_path, previous_month_name)
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    # Iterate through days in the previous month
    d = 1
    page_number = 1
    while d <= last_day_of_previous_month.day:
        # Filter orders by date
        filter_order = driver.find_element(By.XPATH, '//input[@name="filter_date_added"]')
        filter_order.clear()
        filter_order.send_keys(f'{datetime.now().year}-{datetime.now().month - 1 if datetime.now().month > 1 else 12}-{d}')
        filter_button = driver.find_element(By.XPATH, '//body/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[2]/div[7]/button[1]')
        filter_button.click()

        # Retrieve orders for the day
        wait = WebDriverWait(driver, 10)
        matches = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'tr')))

        # Iterate through orders and print invoices
        for match in matches[1:]:
            # Get order details
            cells = match.find_elements(By.TAG_NAME, 'td')
            file_name = cells[1].text

            # Open order details page
            open_order = driver.find_element(By.CSS_SELECTOR, f'a[href*="order_id={file_name}"]')
            open_order.click()

            # Print invoice
            invoice = driver.find_element(By.XPATH, "//a[@data-original-title='Print Invoice']")
            invoice.click()

            # Switch to invoice tab
            driver.switch_to.window(driver.window_handles[1])
            html_source = driver.page_source

            # Save invoice as PDF
            output_file_path = os.path.join(download_folder, file_name + '.pdf')
            pdfkit.from_string(html_source, output_file_path, configuration=config)

            # Close invoice tab
            driver.close()

            # Switch back to order list tab
            driver.switch_to.window(driver.window_handles[0])
            driver.back()

        # Check for pagination and navigate to next page if it exists
        try:
            next_page = driver.find_element(By.XPATH, f'//a[contains(@class, "next")]')
            next_page.click()
            page_number += 1
            wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'tr')))
        except:
            break

        d += 1

driver.quit()
