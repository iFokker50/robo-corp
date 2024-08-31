from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
import zipfile
import time
from pathlib import Path

pdf=PDF()
pdf_list = []

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """

    open_chrome()
    download_csv()
    read_csv()
    archive_receipts(pdf_list,"output/PDF_Order.zip")


def open_chrome():
    """Open Chrome and navigate to Robot page"""
    browser.configure(headless=False)
    page = browser.goto("https://robotsparebinindustries.com/")
    page.click("a:text('Order your robot!')")
    page.click("button:text('OK')")


def download_csv():
    """Download CSV file"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def read_csv():
    """Read CSV data""" 
    table = Tables().read_table_from_csv("orders.csv", header=True)
    print(table)

    for row in table:
        fill_form_information(row)

def fill_form_information(row):
    page = browser.page()
    if page.is_visible("button:text('OK')"):
        page.click("button:text('OK')")

    page.select_option("#head", value=row['Head'])
    page.click(f"#id-body-{row['Body']}")
    page.fill(".form-control[type='number']", row['Legs'])
    page.fill("#address", row['Address'])
    page.click("#order")
    time.sleep(2)
    if page.is_visible("div.alert.alert-danger[role='alert']"):
        page.click("#order")
    if page.is_visible("div.alert.alert-danger[role='alert']"):
        page.click("#order")
    stores_order_as_pdf(row['Order number'])
    screenshot_robot(row['Order number'])
    embed_screenshot_to_receipt(f"output/{row['Order number']}.png",f"output/{row['Order number']}.pdf",row['Order number'])
    pdf_list.append(f"output/{row['Order number']}.pdf")
    page.click("#order-another")


def stores_order_as_pdf(order_number):
    page = browser.page()
    order_results_html = page.locator("#order-completion").inner_html()
    pdf.html_to_pdf(order_results_html, f"output/{order_number}.pdf")

def screenshot_robot(order_number):
    page = browser.page()
    page.screenshot(path=f"output/{order_number}.png")

def embed_screenshot_to_receipt(screenshot,pdf_file,order_number):
    output_pdf_path = f"output/{order_number}.pdf"
    pdf.add_watermark_image_to_pdf(image_path=screenshot, source_path=pdf_file, output_path=output_pdf_path)

def archive_receipts(file_paths, output_zip_path):

    with zipfile.ZipFile(output_zip_path, 'w') as zipf:
        for file_path in file_paths:
            # Add each file to the ZIP
            zipf.write(file_path, arcname=Path(file_path).name)
       
