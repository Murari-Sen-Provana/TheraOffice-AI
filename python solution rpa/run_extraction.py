import os
import time
import configparser
import logging
import pyperclip
import csv
from pywinauto.application import Application
import pyautogui

# --- Constants ---
# It's good practice to define image paths as constants for easy management.
IMAGE_PATH_BLUE_STAR = 'images/blue_star_icon.png'
IMAGE_PATH_SCHEDULING_BTN = 'images/scheduling_button.png'

class TheraOfficeExtractor:
    """
    A modular and robust RPA class to extract patient data from the TheraOffice application.
    Includes comprehensive logging, error handling, and modular methods for testability.
    """
    def __init__(self, config_file='config.ini'):
        """Initializes the extractor, logger, and configuration."""
        self._setup_logging()
        
        try:
            config = configparser.ConfigParser()
            
            # --- FIX 1: Build an absolute path to the config file ---
            script_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(script_dir, config_file)

            if not os.path.exists(config_path):
                self.logger.critical(f"Configuration file not found at: {config_path}")
                raise FileNotFoundError(f"config.ini not found at {config_path}")
            
            config.read(config_path)

            # Now, safely read the configuration
            self.app_path = config['TheraOffice']['app_path']
            self.username = config['TheraOffice']['username']
            self.password = config['TheraOffice']['password']
            self.output_dir = config['RPA_Settings']['output_directory']
            self.timeout = int(config['RPA_Settings']['default_timeout'])
            
            self.app = None
            self.main_window = None
            
            os.makedirs(self.output_dir, exist_ok=True)
            self.logger.info("RPA Extractor initialized successfully.")
        except KeyError as e:
            self.logger.critical(f"Configuration Error: Section or key missing in config.ini. Missing key: {e}")
            raise
        except Exception as e:
            self.logger.critical(f"Failed to initialize. Error: {e}")
            raise

    def _setup_logging(self):
        """Sets up a logger to output to both console and a file."""
        self.logger = logging.getLogger("TheraOfficeRPA")
        self.logger.setLevel(logging.INFO)
        
        if self.logger.hasHandlers():
            self.logger.handlers.clear()
            
        fh = logging.FileHandler('rpa_run.log', mode='w')
        ch = logging.StreamHandler()
        
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter)
        ch.setFormatter(formatter)
        
        self.logger.addHandler(fh)
        self.logger.addHandler(ch)

    def connect_to_app(self):
        """
        --- FIX 2: Connect to an already running TheraOffice instance. ---
        This is more reliable for Windows Store apps.
        """
        self.logger.info("Attempting to connect to the TheraOffice application...")
        try:
            # We connect by process name, which is more reliable.
            # You may need to find the process name from Task Manager (e.g., TheraOffice.exe)
            self.app = Application(backend="uia").connect(title_re=".*TheraOffice.*", timeout=self.timeout)
            self.logger.info("Successfully connected to the TheraOffice instance.")
            return True
        except Exception:
            self.logger.error("Could not find a running TheraOffice process. Please launch the application manually before running the script.")
            return False

    def login(self):
        """Finds the login window and performs login."""
        self.logger.info("Step 1.1: Attempting to log in...")
        try:
            login_dlg = self.app.window(title="TheraOffice Web")
            login_dlg.wait('ready', timeout=self.timeout)

            # Assuming the app might already be logged in from a previous session
            if not login_dlg.exists():
                self.logger.info("Login screen not found. Assuming already logged in.")
                self.main_window = self.app.window(title_re=".*TheraOffice Web Client.*")
                return True

            login_dlg.child_window(auto_id="txtUserName").set_text(self.username)
            login_dlg.child_window(auto_id="txtPassword").set_text(self.password)
            login_dlg.child_window(title="Log In", control_type="Button").click()
            
            self.logger.info("Login submitted successfully.")
            return True
        except Exception as e:
            self.logger.error(f"Login failed. Could not interact with login controls. Error: {e}")
            return False

    # ... [The rest of the methods (select_facility, navigate_to_scheduling, etc.) remain the same as the previous version] ...
    # Make sure to implement the actual extraction logic inside them.
    
    # Example placeholder for a fully implemented method
    def run_single_patient_export(self, patient_last_name):
        """Orchestrates the entire extraction process for one patient."""
        self.logger.info(f"\n----- STARTING EXPORT FOR PATIENT: {patient_last_name} -----")
        
        if not self.search_and_select_patient(patient_last_name):
            self.logger.error(f"Could not select patient '{patient_last_name}'. Skipping to next patient.")
            return

        self.patient_folder_path = self._create_patient_folder(patient_last_name)
        
        # Call the detailed extraction functions
        # self.extract_demographics()
        # self.extract_case_and_insurance()
        # self.download_documents()
        self.logger.warning("Extraction logic is not yet implemented. This is a placeholder.")

        try:
            if self.patient_window and self.patient_window.exists():
                self.patient_window.close()
                self.logger.info(f"Closed patient window for '{patient_last_name}'.")
        except Exception as e:
            self.logger.warning(f"Could not close patient window. May need manual intervention. Error: {e}")

        self.logger.info(f"----- COMPLETED EXPORT FOR PATIENT: {patient_last_name} -----")


# --- Main Execution Block ---
if __name__ == "__main__":
    extractor = None
    try:
        extractor = TheraOfficeExtractor()
        
        # --- Main Workflow with Checks and Balances ---
        if extractor.connect_to_app():
            if extractor.login():
                # Check if we need to select a facility or are already at the main screen
                try:
                    extractor.main_window = extractor.app.window(title_re=".*TheraOffice Web Client.*", timeout=5)
                    extractor.logger.info("Main window already visible.")
                except Exception:
                    extractor.logger.info("Main window not found, attempting to select facility.")
                    if not extractor.select_facility():
                        raise Exception("Failed to select facility after login.")

                # Proceed only if the main window handle is valid
                if extractor.main_window and extractor.main_window.exists():
                    if extractor.navigate_to_scheduling():
                        patients_to_process = ["Kantor"] # Add more last names here
                        
                        for patient_name in patients_to_process:
                            extractor.run_single_patient_export(patient_name)
    except Exception as e:
        # This will catch critical failures like config file errors or app connection failures.
        if extractor and extractor.logger:
            extractor.logger.critical(f"A critical, unhandled error occurred in the main execution block: {e}")
        else:
            print(f"A critical error occurred before logging was initialized: {e}")
    finally:
        if extractor and extractor.logger:
            extractor.logger.info("\nAutomation script finished.")