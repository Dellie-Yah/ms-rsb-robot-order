from robocorp.tasks import task
from robocorp import browser

# from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

import time
from datetime import timedelta
import os

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=2000,
    )
    
    # webscraper.set_selenium_speed(timedelta(seconds=0.8))
    
    open_robot_order_website()
    download_orders()
    orders = get_orders()
    for order in orders:
        fill_the_form(order)
    archive_receipts()
    
def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    
def close_annoying_modal():
    page = browser.page()
    page.click("//button[@class='btn btn-dark']")

def download_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    excel = Tables()
    worksheet = excel.read_table_from_csv("orders.csv", header=True)
    return worksheet

def fill_the_form(row):
    page = browser.page()
    close_annoying_modal()
    
    page.select_option("#head", str(row['Head']))
    page.click(f"//input[@id='id-body-{str(row['Body'])}']")
    page.fill("//input[@placeholder='Enter the part number for the legs']", str(row['Legs']))
    page.fill("//input[@id='address']", str(row['Address']))
    page.click("#preview")
    page.click("#order")
    
    
    is_error_present = page.is_visible("//div[@class='alert alert-danger']")
    while is_error_present:
        time.sleep(3)
        page.click("#order")
        is_error_present = is_error_present = page.is_visible("//div[@class='alert alert-danger']")
        
    pdf_location = store_receipt_as_pdf(str(row['Order number']))
    screenshot_location = screenshot_robot(str(row['Order number']))
    
    embed_screenshot_to_receipt(screenshot_location, pdf_location)
        
    page = browser.page()
    is_order_fulfilled = page.is_visible("//div[@id='order-completion']")
    if is_order_fulfilled:
        page.click("#order-another")
        
def store_receipt_as_pdf(order_number):
    page = browser.page()
    pdf_handler = PDF()


    robot_order_receipt = page.locator("#receipt").inner_html()
    # robot_order_receipt = page.locator("#receipt").screenshot()

    file_name = f"output/receipts/receipt-order-number-{order_number}.pdf"
    pdf_handler.html_to_pdf(robot_order_receipt, file_name)
    
    path = os.path.abspath(file_name)
    
    return path   
    
def screenshot_robot(order_number):
    page = browser.page()
    
    file_name = f"output/screenshots/screenshot-order-number-{order_number}.png"
    
    page.locator("#robot-preview-image").screenshot(path=file_name, type="png")
    
    path = os.path.abspath(file_name)
    
    return path
    
def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf_handler = PDF()
    
    pdf_handler.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)
    
def archive_receipts():
    archive = Archive()
    
    archive.archive_folder_with_zip(folder="output/receipts", archive_name="output/Receipts.zip")