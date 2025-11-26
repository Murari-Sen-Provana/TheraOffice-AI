# main.py
import os
import configparser
import time
from logger_setup import setup_logger
from theraoffice_automation import TheraOfficeExtractor
# We DO NOT import the bruteforce_launcher anymore

def load_config(logger):
    """Load config.ini from same directory"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_dir, "config.ini")

    if not os.path.exists(config_path):
        logger.critical(f"Configuration file not found at: {config_path}")
        raise FileNotFoundError(f"config.ini not found at {config_path}")

    config = configparser.ConfigParser()
    config.read(config_path)
    logger.info(f"Configuration loaded successfully from {config_path}")
    return config

def main():
    logger = setup_logger()
    logger.info("==============================================")
    logger.info("=== TheraOffice Automation (DYNAMIC) STARTING ===")
    logger.info("==============================================")

    try:
        # 1. Load config
        config = load_config(logger)
        facility_name = config.get("RPA_Settings", "facility_name", fallback="Brookline/Allston")
        max_login_wait = config.getint("RPA_Settings", "max_login_wait_seconds", fallback=180)

        # 2. Initialize and perform the login sequence
        extractor = TheraOfficeExtractor(config, logger)

        if not extractor.launch_and_connect():
            logger.critical("PROCESS FAILED: Dynamic launch failed.")
            return

        if not extractor.login():
            logger.critical("PROCESS FAILED: Dynamic login failed.")
            return

        logger.info("Dynamic login succeeded.")
        
        # 3. Wait for and handle the Facility selection window
        state = extractor.wait_until_logged_in(max_wait_seconds=max_login_wait)
        if state == "facility":
            if not extractor.select_facility(facility_name):
                logger.critical("Failed to select facility. Aborting.")
                return
            logger.info("Facility selected successfully.")

        # --- UPDATED POPUP HANDLING WITH DIAGNOSTIC ---
        logger.info("Waiting 3 seconds for popup to appear...")
        time.sleep(3)
        # Optional: Run diagnostic to see popup structure (comment out after first run)
        # extractor.debug_popup_structure()

        # DIAGNOSTIC: List all windows to find the popup
        logger.info("Running diagnostic to find all windows...")
        extractor.debug_all_windows()

        # Now try to dismiss the popup
        logger.info("\nAttempting to dismiss popup...")
        if not extractor.dismiss_shared_user_accounts_warning():
            logger.warning("Could not dismiss popup automatically.")
            # Give extra time for manual dismissal
            # Final diagnostic: try one more time to find it
            logger.info("Running final window scan...")
            extractor.debug_all_windows()
    
            logger.info("Waiting 15 seconds for manual dismissal...")
            logger.info(">>> PLEASE CLICK THE OK BUTTON MANUALLY NOW <<<")
            time.sleep(15)  
        # 5. Now, wait for the main window to become fully active.
        logger.info("Waiting for main window to become active...")
        state = extractor.wait_until_logged_in(max_wait_seconds=60)
        
        if state != "main":
            logger.critical(f"Could not detect active main window after setup. Final state: {state}")
            return
        
        logger.info("Main window is ready.")

        # 6. Navigate to the Scheduling module.
        if extractor.navigate_to_scheduling():
            logger.info("Navigated to Scheduling.")
            
            # Handle the facility selection inside Scheduling
            if extractor.handle_scheduling_facility(facility_name):
                logger.info("Scheduling facility handled.")
            
            # --- NEW STEP: Search and Diagnose ---
            patient_name = "Kantor" # Replace with a real test name if needed
            if extractor.search_for_patient(patient_name):
                logger.info("Search executed. Running diagnostic to find the results table...")
                extractor.debug_search_results()
        
        logger.info("Automation workflow finished.")


    finally:
        logger.info("==============================================")
        logger.info("=== TheraOffice Automation (DYNAMIC) FINISHED ===")
        logger.info("==============================================")


if __name__ == "__main__":
    main()