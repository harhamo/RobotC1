from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from robocorp import vault

LOGIN_CREDENTIALS = "robotsparebin"

page = browser.page()
http = HTTP()

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""   
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data()
    collect_results()
    export_as_pdf()
    log_out()

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.configure(
        slowmo = 100,
    )

    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""
    secret = vault.get_secret(LOGIN_CREDENTIALS)
    page.get_by_label("username").fill(secret["username"])
    page.get_by_label("password").fill(secret["password"])
    page.get_by_role("button", name = ("Log in")).click()

def download_excel_file():
    """Downloads excel file from the given URL"""
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_form_with_excel_data():
    """Read data from excel and fill in the sales form"""
    excel = Files()
    excel.open_workbook("SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()
    for row in worksheet:
        fill_and_submit_sales_form(row)

def fill_and_submit_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page.get_by_label("First name").fill(sales_rep["First Name"])
    page.get_by_label("Last name").fill(sales_rep["Last Name"])
    page.get_by_label("Sales result ($)").fill(str(sales_rep["Sales"]))
    page.get_by_label("Sales target ($)").select_option(str(sales_rep["Sales Target"]))
    page.get_by_role("button", name = ("Submit")).click()

def collect_results():
    """Take a screenshot of the page"""
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf():
    """Export the data to a odf file"""
    pdf = PDF()
    sales_result_html = page.locator("#sales-results").inner_html() 
    pdf.html_to_pdf(sales_result_html, "output/sales_results.pdf")

def log_out():
    """Press the 'Log out' button"""
    page.get_by_role("button", name = "Log out").click()

