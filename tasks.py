from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.FileSystem import FileSystem
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
    browser.configure(slowmo=10)

    open_robot_order_website()
    get_csv()
    accept_terms()
    place_order()
    zip_receipts()
    delete_obsolete_files()


def open_robot_order_website():
    browser.goto(
        "https://robotsparebinindustries.com/#/robot-order")


def accept_terms():
    page = browser.page()
    page.click("button:text('Ok')")


def get_csv():
    http = HTTP()
    http.download(
        "https://robotsparebinindustries.com/orders.csv", target_file='output/orders.csv', overwrite=True)


def get_orders():
    library = Tables()
    orders = library.read_table_from_csv(
        "output/orders.csv", header=True)
    return orders


def store_receipt_as_pdf(order_nr, receipt):
    pdf = PDF()

    pdf.html_to_pdf(receipt, "output/receipts/"+order_nr+".pdf")
    pdf.add_files_to_pdf(["output/"+order_nr +
                         ".png:align=Center"], "output/receipts/"+order_nr+".pdf", append=True)


def place_order():
    orders = get_orders()
    FileSystem.create_directory(self=FileSystem, path='output/receipts')
    for row in orders:
        page = browser.page()
        page.select_option('#head', row['Head'])
        page.set_checked('#id-body-'+row['Body'], True)
        page.fill('.form-control', row['Legs'])
        page.fill('#address', row['Address'])
        page.click('#preview')
        page.click('#order')
        # ToDo retry logic
        while page.is_visible('.alert-danger'):
            page.click('#order')

        order_nr = page.inner_text('.badge-success')
        receipt = page.locator('#receipt').inner_html()
        page.locator(
            "#robot-preview-image").screenshot(path="output/"+order_nr+".png")

        store_receipt_as_pdf(order_nr, receipt)
        FileSystem.remove_file(
            self=FileSystem, path="output/"+order_nr+".png")
        page.click('#order-another')
        accept_terms()

# Function to zip files


def zip_receipts():

    Archive.archive_folder_with_zip(
        self='Archive', folder='output/receipts/', archive_name='output/receipts.zip', recursive=True)


def delete_obsolete_files():
    FileSystem.remove_file(self=FileSystem, path='output/orders.csv')
    FileSystem.remove_directory(
        self=FileSystem, path="output/receipts", recursive=True)
