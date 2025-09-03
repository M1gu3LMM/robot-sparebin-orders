from robocorp.tasks import task
from robocorp import browser

from RPA.Tables import Tables
from RPA.HTTP import HTTP
from RPA.PDF import PDF
from RPA.Archive import Archive

import glob
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
        slowmo=90,
    )
    open_robot_order_website()
    download_csv_file()
    close_annoying_modal()
    orders = get_orders()

    for order in orders:
        fill_the_form(order)
        #get_the_image_robot()
        submit_the_order()
        pdf_file = store_receipt_as_pdf(str(order["Order number"]))
        screenshot = screenshot_robot(str(order["Order number"]))
        embed_screenshot_to_receipt(screenshot, pdf_file)
        order_another_robot()
        close_annoying_modal()

    archive_receipts()
    


def open_robot_order_website():
    """ Navigates to the order website """
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    """Read the data from the csv file"""
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=['Order number','Head','Body','Legs','Address']
    )
    return orders

def close_annoying_modal():
    """Close the pop up message"""
    page = browser.page() 
    page.click("button:text('OK')")

def fill_the_form(order):
    """FIll the form"""
    page = browser.page()


    print(order)
    #head = number_2_model_name(order["Head"])
    
    #print(head)
    page.select_option("#head", str(order["Head"]))

    xpath = f'input[name="body"][value="{order["Body"]}"]'
    page.locator(xpath).check()
    
    page.locator("xpath=//input[@placeholder='Enter the part number for the legs']").fill(str(order["Legs"]))

    page.fill("#address", str(order["Address"]))

def number_2_model_name(number):
    page = browser.page()
    page.click("button:text('Show model info')")
    xpath = f"xpath=//table[@id='model-info']//tbody/tr[{number}]/td[1]"
    model_name = page.locator(xpath).inner_text()
    page.click("button:text('Hide model info')")
    return model_name

def get_the_image_robot():
    """"Obtain the preview image"""
    page = browser.page()
    page.click("button:text('Preview')")

def submit_the_order():
    """"try even if exist the message error"""
    page = browser.page()
    page.click("button:text('Order')")

    while page.locator("div.alert.alert-danger").is_visible():
        page.click("button:text('Order')")

def order_another_robot():
    page = browser.page()
    page.click("button:text('Order another robot')")

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    receipt_info_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    file_name = f"output/{order_number}.pdf"
    pdf.html_to_pdf(receipt_info_html, file_name)
    return file_name

def screenshot_robot(order_number):
    """Take a screenshot of the robot"""
    page = browser.page()
    image_name = f"output/{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=image_name)
    return image_name

def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Append the screenshot image to the end of the PDF receipt"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=pdf_file,
        coverage=0.2  # Ajusta el tamaño relativo (30%) de la imagen
    )

def archive_receipts():
    """Crea un ZIP con los recibos .pdf en la carpeta output/"""
    os.makedirs("output", exist_ok=True)

    zip_path = "output/receipts.zip"
    if os.path.exists(zip_path):
        os.remove(zip_path)

    archive = Archive()
    # Crea el ZIP NUEVO con sólo los PDFs (no subdirectorios)
    archive.archive_folder_with_zip(
        folder="output",
        archive_name=zip_path,
        include="*.pdf",
        recursive=False,       # opcional (por defecto es False)
        # compression="deflated"  # opcional: 'stored' (default), 'deflated', 'bzip2', 'lzma'
    )