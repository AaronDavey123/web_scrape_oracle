import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException, WebDriverException, StaleElementReferenceException
from multiprocessing import Process, Pool

# Function to read expected names from a file
def read_expected_names(file_path):
    with open(file_path, 'r') as file:
        return [line.strip() for line in file if line.strip()]

# Function to extract data from the target page
def extract_data(driver, section_name, page_name):
    print(f"Extracting data for table: {page_name} in section: {section_name}")
    data = {}

    # Extract header
    header = driver.find_element(By.XPATH, "//header/h1[@class='fa-chapter topic_link']")
    data['header'] = header.text

    # Extract paragraph below the header
    paragraph = driver.find_element(By.XPATH, "//p[@class='p']")
    data['paragraph'] = paragraph.text

    # Extract details section
    try:
        details_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Details']]")
        details = details_section.find_elements(By.XPATH, ".//li/p[@class='p']")
        data['details'] = [detail.text for detail in details]
    except Exception:
        data['details'] = []

    # Extract primary key section
    try:
        primary_key_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Primary-Key']]")
        primary_key_table = primary_key_section.find_element(By.XPATH, ".//table[@summary='Primary Key']")
        primary_key_rows = primary_key_table.find_elements(By.XPATH, ".//tr[@class='row']")
        primary_key_data = []
        for row in primary_key_rows:
            row_data = [cell.text for cell in row.find_elements(By.XPATH, ".//td[@class='entry']")]
            if len(row_data) == 2:  # Ensure row matches the expected number of columns
                primary_key_data.append(row_data)
            else:
                pass  # Skipping invalid row in Primary Key
        data['primary_key'] = primary_key_data
    except Exception:
        data['primary_key'] = []

    # Extract columns section
    try:
        columns_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Columns']]")
        columns_table = columns_section.find_element(By.XPATH, ".//table[@summary='Columns']")
        columns_headers = [header.text for header in columns_table.find_elements(By.XPATH, ".//thead/tr/th")]
        columns_rows = columns_table.find_elements(By.XPATH, ".//tbody/tr")
        data['columns'] = [[cell.text for cell in row.find_elements(By.XPATH, ".//td")] for row in columns_rows]
        data['columns_headers'] = columns_headers
    except Exception:
        data['columns'] = []
        data['columns_headers'] = []

    # Extract indexes section
    try:
        indexes_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Indexes']]")
        indexes_table = indexes_section.find_element(By.XPATH, ".//table[@summary='Indexes']")
        indexes_headers = [header.text for header in indexes_table.find_elements(By.XPATH, ".//thead/tr/th")]
        indexes_rows = indexes_table.find_elements(By.XPATH, ".//tbody/tr")
        data['indexes'] = [[cell.text for cell in row.find_elements(By.XPATH, ".//td")] for row in indexes_rows]
        data['indexes_headers'] = indexes_headers
    except Exception:
        data['indexes'] = []
        data['indexes_headers'] = []

    # Extract foreign keys section
    try:
        foreign_keys_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Foreign-Keys']]")
        foreign_keys_table = foreign_keys_section.find_element(By.XPATH, ".//table[@summary='Foreign Keys']")
        foreign_keys_headers = [header.text for header in foreign_keys_table.find_elements(By.XPATH, ".//thead/tr/th")]
        foreign_keys_rows = foreign_keys_table.find_elements(By.XPATH, ".//tbody/tr")
        data['foreign_keys'] = [[cell.text for cell in row.find_elements(By.XPATH, ".//td")] for row in foreign_keys_rows]
        data['foreign_keys_headers'] = foreign_keys_headers
    except Exception:
        data['foreign_keys'] = []
        data['foreign_keys_headers'] = []

    return data

# Function to extract data from the views page
def extract_view_data(driver, section_name, page_name):
    print(f"Extracting data for view: {page_name} in section: {section_name}")
    data = {}

    # Extract header
    header = driver.find_element(By.XPATH, "//header/h1[@class='fa-chapter topic_link']")
    data['header'] = header.text

    # Extract details section
    try:
        details_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Details']]")
        details = details_section.find_elements(By.XPATH, ".//li/p[@class='p']")
        data['details'] = [detail.text for detail in details]
    except Exception:
        data['details'] = []

    # Extract columns section
    try:
        columns_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Columns']]")
        columns_table = columns_section.find_element(By.XPATH, ".//table[@summary='Columns']")
        columns_headers = [header.text for header in columns_table.find_elements(By.XPATH, ".//thead/tr/th")]
        columns_rows = columns_table.find_elements(By.XPATH, ".//tbody/tr")
        columns_data = []
        for row in columns_rows:
            row_data = [cell.text for cell in row.find_elements(By.XPATH, ".//td")]
            if len(row_data) == len(columns_headers):  # Ensure row matches header length
                columns_data.append(row_data)
        data['columns'] = columns_data
        data['columns_headers'] = columns_headers
    except Exception:
        data['columns'] = []
        data['columns_headers'] = []

    # Extract query section
    try:
        query_section = driver.find_element(By.XPATH, "//section[@class='section'][h2[@id='Query']]")
        query_table = query_section.find_element(By.XPATH, ".//table[@summary='Query']")
        query_headers = [header.text for header in query_table.find_elements(By.XPATH, ".//thead/tr/th")]
        query_rows = query_table.find_elements(By.XPATH, ".//tbody/tr")
        query_data = []
        for row in query_rows:
            row_data = [cell.text for cell in row.find_elements(By.XPATH, ".//td")]
            if len(row_data) == len(query_headers):  # Ensure row matches header length
                query_data.append(row_data)
        data['query'] = query_data
        data['query_headers'] = query_headers
    except Exception:
        data['query'] = []
        data['query_headers'] = []

    return data

# Helper function to save a DataFrame to an Excel sheet
def save_dataframe_to_excel(writer, data, sheet_name, headers=None):
    if data:
        try:
            if headers and len(headers) > 1:
                if not all(isinstance(row, list) and len(row) == len(headers) for row in data):
                    raise ValueError(f"Data shape mismatch: Expected {len(headers)} columns, but got inconsistent row lengths.")
            df = pd.DataFrame(data, columns=headers) if headers else pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception:
            pass

# Refactored save_to_excel function
def save_to_excel(data, file_path):
    with pd.ExcelWriter(file_path) as writer:
        save_dataframe_to_excel(writer, [{'Header': data['header']}], 'Header')
        save_dataframe_to_excel(writer, data['details'], 'Details', ['Details'])
        save_dataframe_to_excel(writer, data.get('primary_key'), 'Primary Key', ['Name', 'Columns'])
        save_dataframe_to_excel(writer, data.get('columns'), 'Columns', data.get('columns_headers'))
        save_dataframe_to_excel(writer, data.get('indexes'), 'Indexes', data.get('indexes_headers'))
        save_dataframe_to_excel(writer, data.get('foreign_keys'), 'Foreign Keys', data.get('foreign_keys_headers'))
        save_dataframe_to_excel(writer, data.get('query'), 'Query', data.get('query_headers'))

# Function to expand a dropdown with retries and fallback mechanism
def expand_dropdown_with_retries(driver, dropdown_id, dropdown_type, section_name, retries=10):
    dropdown = None  # Initialize dropdown to avoid unbound local variable error
    for attempt in range(retries):
        try:
            print(f"Attempting to click {dropdown_type} dropdown for section: {section_name} (Attempt {attempt + 1}/{retries})")
            
            # Scroll to the dropdown area to ensure it is visible
            dropdown_area = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, f"//li[@id='{dropdown_id}']"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_area)
            print(f"Scrolled to {dropdown_type} dropdown area for section: {section_name}")

            # Locate and click the dropdown
            dropdown = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[@id='{dropdown_id}'][contains(@class, 'oj-collapsed')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
            ActionChains(driver).move_to_element(dropdown).click().perform()
            print(f"Successfully clicked {dropdown_type} dropdown for section: {section_name}")

            # Wait for and click the first page in the dropdown
            first_page = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, f"//li[@id='{dropdown_id}']//li[@class='oj-typography-body-xs tree-view-row oj-treeview-item oj-treeview-leaf'][1]//span[@class='oj-treeview-item-text tree-view-item']"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", first_page)
            ActionChains(driver).move_to_element(first_page).click().perform()
            print(f"Successfully clicked the first page in {dropdown_type} dropdown for section: {section_name}")
            return
        except Exception as e:
            print(f"Error clicking {dropdown_type} dropdown or first page for section: {section_name}: {e}")
            # Retry after a delay
            time.sleep(5)
    print(f"Failed to click {dropdown_type} dropdown or first page for section: {section_name} after {retries} retries")
    raise Exception(f"Failed to click {dropdown_type} dropdown or first page for section: {section_name}")

# Function to refresh the WebDriver session if it becomes unresponsive
def refresh_driver_session():
    global driver
    print("Refreshing WebDriver session...")
    driver.quit()
    driver = start_webdriver_session()

# Function to extract all pages in a dropdown sequentially
def extract_all_pages(driver, dropdown_id, section_name, save_dir, extract_function):
    try:
        page_elements = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.XPATH, f"//li[@id='{dropdown_id}']//li[@class='oj-typography-body-xs tree-view-row oj-treeview-item oj-treeview-leaf']"))
        )
        page_names = [
            page.find_element(By.XPATH, ".//span[@class='oj-treeview-item-text tree-view-item']").text.strip()
            for page in page_elements
        ]

        for page_name in page_names:
            retries = 3  # Retry up to 3 times for transient errors
            for attempt in range(retries):
                try:
                    print(f"Processing page: {page_name} in section: {section_name} (Attempt {attempt + 1}/{retries})")
                    # Re-locate the page element to avoid stale element issues
                    page_element = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.XPATH, f"//span[text()='{page_name}']"))
                    )
                    ActionChains(driver).move_to_element(page_element).click().perform()
                    time.sleep(2)
                    data = extract_function(driver, section_name, page_name)
                    save_path = os.path.join(save_dir, f'{page_name.lower()}.xlsx')
                    save_to_excel(data, save_path)
                    break  # Exit the retry loop if successful
                except StaleElementReferenceException:
                    print(f"Stale element encountered for page {page_name}. Retrying...")
                    time.sleep(2)  # Wait before retrying
                except Exception as e:
                    print(f"Error processing page {page_name} in section {section_name}: {e}")
                    if attempt == retries - 1:
                        raise  # Raise the exception if all retries fail
    except Exception as e:
        print(f"Error extracting pages for section: {section_name}: {e}")

# Helper function to create directories dynamically
def create_save_directories(base_path, section_name, create_tables=True, create_views=True):
    tables_dir = os.path.join(base_path, section_name, "Tables") if create_tables else None
    views_dir = os.path.join(base_path, section_name, "Views") if create_views else None
    if tables_dir:
        os.makedirs(tables_dir, exist_ok=True)
    if views_dir:
        os.makedirs(views_dir, exist_ok=True)
    return tables_dir, views_dir

# Helper function to retry an operation
def retry_operation(operation, retries=5, delay=5):
    for attempt in range(retries):
        try:
            return operation()
        except Exception:
            time.sleep(delay)
    raise Exception(f"Operation failed after {retries} retries")

# Prompt the user to input the path for saving directories
def get_save_path():
    default_path = os.path.join(os.path.expanduser("~"), "Desktop", "Oracle_Excel_Files")
    print(f"Please input the path you want to save the Excel files to. We recommend {default_path}")
    user_input = input("Enter the path (or press Enter to use the recommended path): ").strip()
    return user_input if user_input else default_path

# Define sections with their table and view requirements
sections = [
    ("2-AI", True, True),
    ("3-Absence-Management", True, True),
    ("4-Benefits", True, True),
    ("5-Career-Development", True, True),
    ("6-Celebrate", True, True),
    ("7-Compensation", True, True),
    ("8-Corporate-Social-Responsibility", True, True),
    ("9-Fast-Formula", True, True),
    ("10-Global-Human-Resources", True, True),
    ("11-Global-Payroll", True, True),
    ("12-Global-Payroll-Interface", True, True),
    ("13-Goal-Management", True, True),
    ("14-HCM-Common", True, True),
    ("15-HCM-Communicate", False, True),
    ("16-HCM-Configuration-Workbench", True, False),
    ("17-HCM-Country-and-Vertical-Extensions", False, True),
    ("18-HCM-Extracts", True, False),
    ("19-Performance-Management", True, True),
    ("20-Profile-Management", True, True),
    ("21-Questionnaire", True, True),
    ("22-Recruiting", True, True),
    ("23-Social-Connection", True, False),
    ("24-Succession-Management", True, True),
    ("25-Talent-Review", True, True),
    ("26-Time-and-Labor", True, True),
    ("27-Touchpoints", True, True),
    ("28-Work-Life", True, True),
    ("29-Workforce-Directory-Management", True, True),
    ("30-Workforce-Health-and-Safety-Incidents", True, True),
    ("31-Workforce-Management", True, True),
    ("32-Workforce-Modeling", True, True),
    ("33-Workforce-Predictions", True, True),
    ("34-Workforce-Reputation-Management", True, True),
    ("35-Workforce-Scheduling", True, True),
]

# Function to start a new WebDriver session
def start_webdriver_session():
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager

    options = Options()
    options.add_argument("--headless")  # Run Chrome in headless mode
    options.add_argument("--disable-gpu")  # Disable GPU acceleration
    options.add_argument("--window-size=1920,1080")  # Set window size for headless mode
    options.add_argument("--no-sandbox")  # Bypass OS security model (useful for some environments)
    options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems in some environments

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://docs.oracle.com/en/cloud/saas/human-resources/25a/oedmh/index.html")
    time.sleep(5)
    
    # Check for iframes and switch to the correct one
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    if len(iframes) > 1:
        driver.switch_to.frame(iframes[1])
    else:
        driver.switch_to.frame(iframes[0])
        
    # Handle cookie consent if present
    try:
        accept_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@class='call' and contains(text(), 'Accept all')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", accept_button)
        driver.execute_script("arguments[0].click();", accept_button)
        time.sleep(2)
    except Exception:
        pass
    driver.refresh()
    time.sleep(2)
    driver.execute_script("document.body.style.zoom='75%'")                     
    return driver

# Function to process a single section in a separate browser instance
def process_section(section):
    section_title, table_dropdown_id, view_dropdown_id, tables_save_dir, views_save_dir = section
    driver = start_webdriver_session()  # Start a new WebDriver session for this section

    try:
        # Process tables dropdown if available
        if table_dropdown_id:
            try:
                print(f"Processing tables for section: {section_title}")
                expand_dropdown_with_retries(driver, table_dropdown_id, "tables", section_title)
                extract_all_pages(driver, table_dropdown_id, section_title, tables_save_dir, extract_data)
            except Exception as e:
                print(f"Error processing tables for section {section_title}: {e}")
        else:
            print(f"Skipping tables for section: {section_title} as no tables dropdown is defined.")

        # Process views dropdown if available
        if view_dropdown_id:
            try:
                print(f"Processing views for section: {section_title}")
                expand_dropdown_with_retries(driver, view_dropdown_id, "views", section_title)
                extract_all_pages(driver, view_dropdown_id, section_title, views_save_dir, extract_view_data)
            except Exception as e:
                print(f"Error processing views for section {section_title}: {e}")
        else:
            print(f"Skipping views for section: {section_title} as no views dropdown is defined.")
    except Exception as e:
        print(f"Error processing section {section_title}: {e}")
    finally:
        driver.quit()  # Ensure the browser is closed after processing

# Function to process sections using a pool of processes in order
def process_sections_with_pool(sections, directories, pool_size=5):
    treeview_mapping = {
        "2-AI": ("treeview4_0", "treeview4_1"),
        "3-Absence-Management": ("treeview6_0", "treeview6_1"),
        "4-Benefits": ("treeview8_0", "treeview8_1"),
        "5-Career-Development": ("treeview10_0", "treeview10_1"),
        "6-Celebrate": ("treeview12_0", "treeview12_1"),
        "7-Compensation": ("treeview14_0", "treeview14_1"),
        "8-Corporate-Social-Responsibility": ("treeview16_0", "treeview16_1"),
        "9-Fast-Formula": ("treeview18_0", "treeview18_1"),
        "10-Global-Human-Resources": ("treeview20_0", "treeview20_1"),
        "11-Global-Payroll": ("treeview22_0", "treeview22_1"),
        "12-Global-Payroll-Interface": ("treeview24_0", "treeview24_1"),
        "13-Goal-Management": ("treeview26_0", "treeview26_1"),
        "14-HCM-Common": ("treeview28_0", "treeview28_1"),
        "15-HCM-Communicate": (None, "treeview30_0"),
        "16-HCM-Configuration-Workbench": ("treeview32_0", None),
        "17-HCM-Country-and-Vertical-Extensions": (None, "treeview34_0"),
        "18-HCM-Extracts": ("treeview36_0", None),
        "19-Performance-Management": ("treeview38_0", "treeview38_1"),
        "20-Profile-Management": ("treeview40_0", "treeview40_1"),
        "21-Questionnaire": ("treeview42_0", "treeview42_1"),
        "22-Recruiting": ("treeview44_0", "treeview44_1"),
        "23-Social-Connection": ("treeview46_0", None),
        "24-Succession-Management": ("treeview48_0", "treeview48_1"),
        "25-Talent-Review": ("treeview50_0", "treeview50_1"),
        "26-Time-and-Labor": ("treeview52_0", "treeview52_1"),
        "27-Touchpoints": ("treeview54_0", "treeview54_1"),
        "28-Work-Life": ("treeview56_0", "treeview56_1"),
        "29-Workforce-Directory-Management": ("treeview58_0", "treeview58_1"),
        "30-Workforce-Health-and-Safety-Incidents": ("treeview60_0", "treeview60_1"),
        "31-Workforce-Management": ("treeview62_0", "treeview62_1"),
        "32-Workforce-Modeling": ("treeview64_0", "treeview64_1"),
        "33-Workforce-Predictions": ("treeview66_0", "treeview66_1"),
        "34-Workforce-Reputation-Management": ("treeview68_0", "treeview68_1"),
        "35-Workforce-Scheduling": ("treeview70_0", "treeview70_1"),
    }

    with Pool(processes=pool_size) as pool:
        pool.map(process_section, [
            (
                section_name,
                treeview_mapping[section_name][0] if create_tables else None,
                treeview_mapping[section_name][1] if create_views else None,
                directories[section_name]["tables"],
                directories[section_name]["views"]
            )
            for section_name, create_tables, create_views in sections
        ])

# Main function to extract data for all sections using a process pool
def extract_data_with_pool():
    # Get the base path from the user
    global base_path
    base_path = get_save_path()

    # Create save directories dynamically for all sections
    directories = {}
    for section_name, create_tables, create_views in sections:
        tables_dir, views_dir = create_save_directories(base_path, section_name, create_tables, create_views)
        directories[section_name] = {"tables": tables_dir, "views": views_dir}

    # Process all sections
    print("Processing all sections...")
    process_sections_with_pool(sections, directories, pool_size=15)
    print("Data extraction completed for all sections.")

# Start the batch extraction process
if __name__ == "__main__":
    extract_data_with_pool()

# End of script

