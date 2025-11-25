# main.py
import os
import configparser

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
        
        # --- THE CORRECTED LOGIC ---
        # 4. Immediately after facility selection, check for and dismiss the pop-up.
        # This is the correct time to look for it.
        extractor.dismiss_shared_user_accounts_warning()

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
            
            # --- NEW STEP: Handle the facility selection inside Scheduling ---
            if extractor.handle_scheduling_facility(facility_name):
                logger.info("Scheduling facility handled.")
            
            logger.info("Now ready to search for patients...")
            # extractor.search_for_patient("Kantor") # Uncomment when ready
        
        logger.info("Automation workflow finished.")

    finally:
        logger.info("==============================================")
        logger.info("=== TheraOffice Automation (DYNAMIC) FINISHED ===")
        logger.info("==============================================")


if __name__ == "__main__":
    main()