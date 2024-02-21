from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

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
        slowmo=100
    )
    open_robot_order_website()
    download_excel_file()
    process_orders_from_file()
    create_archive()

def open_robot_order_website():
    """Navigates to the website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def process_orders_from_file():
    """Opens the file, loops through the orders and processes each of them"""
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", columns=["Order number", "Head", "Body", "Legs", "Address"]
    )
    for order in orders:
        fill_in_form(order)

def fill_in_form(order):
    """Takes the order data and uploads it into the system"""
    page = browser.page()
    close_modal()

    page.select_option("#head", str(order["Head"]))
    page.click("#id-body-" + str(order["Body"]))
    page.fill("*[placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", str(order["Address"]))

    while not page.query_selector("#order-another"):
        page.click("#order")
        
    store_receipt_as_pdf(str(order["Order number"]))
    page.click("#order-another")

def close_modal():
    """Closes modal on page"""
    page = browser.page()
    page.click("text=OK")

def store_receipt_as_pdf(order_number):
    """Creates PDF containing order and screenshot"""
    page = browser.page()
    
    screenshot = f"output/receipts/{order_number}.png"
    page.screenshot(path=screenshot)
   
    receipt = page.locator("#receipt").inner_html()
    pdf = PDF()
    file = f"output/receipts/{order_number}.pdf"
    pdf.html_to_pdf(receipt, file)

    combined_pdf = PDF()
    combined_pdf.add_files_to_pdf(
        files=[screenshot], 
        target_document=file, 
        append=True
    )

def create_archive():
    """Creates a ZIP archive of the PDFs"""
    archive = Archive()
    archive.archive_folder_with_zip("output/receipts", "output/receipts.zip", include="*.pdf")
