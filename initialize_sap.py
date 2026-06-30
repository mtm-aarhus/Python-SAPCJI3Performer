from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import random
import string
import time
import psutil
import os
import win32com.client


def download_sap(driver: webdriver.Chrome, downloads_folder, orchestrator_connection, parent_tab): 
    before = set(os.listdir(downloads_folder))
    driver.execute_script("arguments[0].click();", parent_tab)
    
    start_time = time.time()
    timeout = 10  # seconds

    while time.time() - start_time < timeout:
        time.sleep(0.25)
        after = set(os.listdir(downloads_folder))
        new_files = after - before
        if new_files:
            # Only consider .sap files
            for file in new_files:
                if file.endswith(".sap"):
                    full_path = os.path.join(downloads_folder, file)
                    orchestrator_connection.log_info(f"Found SAP file: {file}")
                    return full_path
    raise TimeoutError("SAP file not downloaded.")

    
    
def initialize_sap(orchestrator_connection: OrchestratorConnection):
    # Opus bruger
    OpusLogin = orchestrator_connection.get_credential("OpusBruger")
    OpusUser = OpusLogin.username
    OpusPassword = OpusLogin.password
    
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    
    
    # Configure Chrome options
    chrome_options = Options()
    chrome_options.add_argument('--remote-debugging-pipe')
    # chrome_options.add_argument("--headless=new")  
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_experimental_option("prefs", {
    "download.default_directory": downloads_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    })


    driver = webdriver.Chrome(options=chrome_options)
    driver.get(orchestrator_connection.get_constant("OpusAdgangUrl").value)
    orchestrator_connection.log_info("Navigating to Opus login page")
    
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "logonuidfield")))
    
    driver.find_element(By.ID, "logonuidfield").send_keys(OpusUser)
    driver.find_element(By.ID, "logonpassfield").send_keys(OpusPassword)
    driver.find_element(By.ID, "buttonLogon").click()
    
    orchestrator_connection.log_info("Logged in to Opus portal successfully")
    
    WebDriverWait(driver, 60).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
   
    tab_label_xpath = "//div[contains(@class, 'TabText_SmallTabs') and contains(text(), 'Mine Genveje')]"
    
    try:
        # Wait until tab is present
        tab_label = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, tab_label_xpath))
        )
        # Go up to parent tab container 
        parent_tab = tab_label.find_element(By.XPATH, "./ancestor::div[contains(@id, 'tabIndex')]")
            
    except:
        orchestrator_connection.log_info('Trying to find change button')
        WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "changeButton")))
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "changeButton")))
        
        lower = string.ascii_lowercase
        upper = string.ascii_uppercase
        digits = string.digits
        special = "!@#&%"

        password_chars = []
        password_chars += random.choices(lower, k=2)
        password_chars += random.choices(upper, k=2)
        password_chars += random.choices(digits, k=4)
        password_chars += random.choices(special, k=2)

        random.shuffle(password_chars)
        password = ''.join(password_chars)

        driver.find_element(By.ID, "inputUsername").send_keys(OpusPassword)
        driver.find_element(By.NAME, "j_sap_password").send_keys(password)
        driver.find_element(By.NAME, "j_sap_again").send_keys(password)
        driver.find_element(By.ID, "changeButton").click()
        orchestrator_connection.update_credential('OpusBruger', OpusUser, password)
        orchestrator_connection.log_info('Password changed and credential updated')
        time.sleep(2)
        driver.get(orchestrator_connection.get_constant("OpusAdgangUrl").value)
        # Wait until tab is present
        tab_label = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, tab_label_xpath))
        )
        # Go up to parent tab container 
        parent_tab = tab_label.find_element(By.XPATH, "./ancestor::div[contains(@id, 'tabIndex')]")
        
    filepath = download_sap(driver, downloads_folder, orchestrator_connection, parent_tab)
    driver.quit()
    
    success = False
    timeout = 30
    start_time = time.time()
    
    os.startfile(filepath)

    while time.time() - start_time < timeout:
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and 'saplogon' in proc.info['name'].lower():
                success = True
                break  # Exit the for-loop once found
        if success:
            break  # Exit the while-loop if process was found
        time.sleep(1)
        
    dismiss_until_easy_access(30)
    return success

def dismiss_until_easy_access(timeout=30):
    start_time = time.time()


    session = None
    while time.time() - start_time < timeout:
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine

            if application.Children.Count > 0:
                connection = application.Children(0)
                if connection.Children.Count > 0:
                    session = connection.Children(0)
                    break
        except Exception:
            pass  # Keep waiting if not ready yet

        time.sleep(0.5)

    if not session:
        raise TimeoutError("SAP session not available within timeout.")

    # Step 2: Loop until SAP Easy Access is reached
    while time.time() - start_time < timeout:
        try:
            active_window = session.ActiveWindow
            window_name = active_window.Name
            window_title = active_window.Text.strip()

            if window_name == "wnd[0]" and window_title.startswith("SAP Easy Access"):
                return True

            if window_name != "wnd[0]":
                print(f"Non-main window detected: {window_name} '{window_title}'")
                try:
                    btn = active_window.FindById("tbar[0]/btn[0]")
                    btn.Press()
                    print(f"Dismissed window '{window_title}' using btn[0]")
                except Exception as e:
                    print(f"Could not dismiss '{window_title}': {e}")
            else:
                print(f"Main window open but not Easy Access: '{window_title}'")
        except Exception as e:
            print(f"Error during SAP GUI window check: {e}")

        time.sleep(0.5)

    raise TimeoutError("SAP Easy Access screen not reached within timeout.")