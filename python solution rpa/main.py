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

        # 2. Use the DYNAMIC automation class from theraoffice_automation.py
        extractor = TheraOfficeExtractor(config, logger)

        # Launch the app dynamically
        if not extractor.launch_and_connect():
            logger.critical("PROCESS FAILED: Dynamic launch failed.")
            return

        # Log in dynamically (without using coordinates)
        if not extractor.login():
            logger.critical("PROCESS FAILED: Dynamic login failed.")
            return

        logger.info("Dynamic login succeeded.")
        
        # 3. Wait for the next screen after login
        logger.info("Waiting for a logged-in state (facility dialog or main window)...")
        state = extractor.wait_until_logged_in(max_wait_seconds=max_login_wait)

        if state is None or state == "timeout":
            logger.critical("Timed out waiting for login to complete. Aborting.")
            return

        # 4. Handle the Facility selection window
        if state == "facility":
            logger.info(f"Facility selection dialog detected; selecting '{facility_name}'...")
            if not extractor.select_facility(facility_name):
                logger.critical("Failed to select facility. Aborting.")
                return

            logger.info("Facility selected. Waiting for main window again...")
            state = extractor.wait_until_logged_in(max_wait_seconds=60)

        if state != "main":
            logger.critical(f"Could not find main window after login. State: {state}")
            return
        
        logger.info("Main window is ready.")

        # 5. Dismiss any popups and continue your automation
        extractor.dismiss_shared_user_accounts_warning()

        logger.info("Automation workflow can now proceed.")
        # Your patient processing logic would go here

    except Exception as e:
        logger.critical(f"An unhandled error occurred in main(): {e}", exc_info=True)

    finally:
        logger.info("==============================================")
        logger.info("=== TheraOffice Automation (DYNAMIC) FINISHED ===")
        logger.info("==============================================")

if __name__ == "__main__":
    main()