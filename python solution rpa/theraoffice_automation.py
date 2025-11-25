# theraoffice_automation.py
import os
import time
import subprocess
import psutil
import pywinauto
from pywinauto.application import Application
from playwright.sync_api import sync_playwright
from pywinauto.keyboard import send_keys
import time
from pywinauto import Desktop
from playwright.sync_api import sync_playwright


class TheraOfficeExtractor:
    """
    Handles all automation of the TheraOffice application using a hybrid approach:
    - subprocess.Popen with debugging for reliable launch.
    - Playwright for robust interaction with the web-based login screen.
    - pywinauto for interaction with native Windows controls post-login.
    """
    def __init__(self, config, logger):
        """Initializes the extractor with configuration and logger."""
        self.config = config
        self.logger = logger
        
        # Extract values from config for easy access
        self.executable_path = self.config['TheraOffice']['executable_path']
        self.process_name = self.config['TheraOffice']['process_name']
        self.username = self.config['TheraOffice']['username']
        self.password = self.config['TheraOffice']['password']
        self.timeout = int(self.config['RPA_Settings']['default_timeout'])

        self.app = None
        self.main_window = None

    def _kill_existing_processes(self):
        """Ensures a clean state by terminating any lingering TheraOffice processes."""
        self.logger.info(f"Checking for and terminating any existing '{self.process_name}' processes...")
        killed_count = 0
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() == self.process_name.lower():
                    self.logger.info(f"Terminating PID {proc.pid} ({proc.info['name']})")
                    proc.terminate()
                    proc.wait(timeout=5)
                    killed_count += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired) as e:
                self.logger.warning(f"Could not terminate process {proc.pid}: {e}")
        if killed_count == 0:
            self.logger.info("No existing processes found.")

    def launch_and_connect(self):
        """
        Launches the app after setting an environment variable to force the remote 
        debugging port to open, then connects the main window handle.
        """
        self._kill_existing_processes()
        self.logger.info("Step 1.0: Launching TheraOffice with debugging enabled via environment variable...")
        try:
            debugging_port = 9223  # The port we will connect to
            
            # --- THE KEY FIX ---
            # Create a copy of the current environment variables
            launch_env = os.environ.copy()
            # Set the special variable that tells WebView2 apps to open the debugging port
            launch_env["WEBVIEW2_REMOTE_DEBUGGING_PORT"] = str(debugging_port)
            self.logger.info(f"Set WEBVIEW2_REMOTE_DEBUGGING_PORT={debugging_port}")

            # Launch the app using Popen, but pass in the modified environment
            subprocess.Popen([self.executable_path], env=launch_env)
            
            self.logger.info("Launch command sent. Waiting for process to appear...")
            time.sleep(12)

            self.logger.info(f"Connecting main window handle with pywinauto: '{self.process_name}'...")
            self.app = Application(backend="uia").connect(path=self.process_name, timeout=self.timeout)
            self.logger.info("Successfully launched TheraOffice and connected main window handle.")
            return True
        except Exception as e:
            self.logger.error(f"Failed to launch or connect. Check executable path in config.ini. Error: {e}")
            return False
        

    def find_window_patiently(self, title_re, class_name, timeout=45):
        """
        Actively polls the desktop every second to find a window that matches
        the criteria. This is the most reliable way to handle slow-starting apps.
        """
        self.logger.info(f"Patiently searching for window (title: '{title_re}', class: '{class_name}') for up to {timeout}s...")
        from pywinauto import Desktop
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Perform a search across the entire desktop
                target_window = Desktop(backend="uia").window(
                    title_re=title_re,
                    class_name=class_name,
                    visible_only=True
                )
                # If the window exists and is visible, we've found it!
                if target_window.exists() and target_window.is_visible():
                    self.logger.info("Window found successfully!")
                    return target_window
            except Exception:
                # Ignore errors while the window doesn't exist yet
                pass
            # Wait one second before trying again
            time.sleep(1)
        
        # If the loop finishes, the window was never found
        raise TimeoutError(f"Timed out after {timeout} seconds waiting for the application window.")

    def login(self):
        """
        Performs login by finding the application window by its unique class name,
        which is the most reliable method for apps with slow-loading titles.
        """
        self.logger.info("Step 1.1: Attempting login by finding window via its unique Class Name...")
        from pywinauto import Desktop
        
        try:
            # --- THE FINAL WORKING SOLUTION ---
            # We search ONLY by the unique class name we discovered in the diagnostic logs.
            window_classname = "WindowsForms10.Window.8.app.0.1629f15_r7_ad1"
            
            self.logger.info(f"Searching for window with Class Name: '{window_classname}'...")
            
            # Find the window using only its class name for a perfect match.
            login_window = Desktop(backend="uia").window(
                class_name=window_classname
            )
            login_window.wait('visible', timeout=30, retry_interval=1)
            
            self.logger.info(f"Correct login window found! Title is now: '{login_window.window_text()}'")
            login_window.set_focus()

            # The rest of the logic uses the reliable AutomationIds.
            username_field = login_window.child_window(auto_id="email", control_type="Edit")
            username_field.wait('enabled', timeout=10)
            username_field.set_text(self.username)
            self.logger.info("Filled username.")

            password_field = login_window.child_window(auto_id="password", control_type="Edit")
            password_field.wait('enabled', timeout=10)
            password_field.set_text(self.password)
            self.logger.info("Filled password.")

            login_button = login_window.child_window(auto_id="btn-login", control_type="Button")
            login_button.wait('enabled', timeout=10)
            login_button.click_input()
            self.logger.info("Clicked Log In button.")

            self.logger.info("Login submitted. Waiting 5 seconds for the next screen...")
            time.sleep(5)
            return True

        except Exception as e:
            self.logger.error("A critical error occurred during the login process.")
            self.logger.error(f"Error details: {str(e)}")
            return False

    def select_facility(self, facility_name: str = "Brookline/Allston") -> bool:
        """
        Finds the facility dialog and selects an item using a name-to-ID mapping,
        the most robust method for handling this application's custom table control.
        """
        from pywinauto import Desktop
        import time

        self.logger.info(f"Attempting to select facility '{facility_name}'...")

        # --- THIS MAPPING IS THE KEY TO THE FINAL SOLUTION ---
        # It connects the visible name to the internal ID we found in the logs.
        facility_map = {
            "Brookline/Allston": "Row0_SRFACILITY",
            "Concord": "Row1_SRFACILITY",
            "Downtown": "Row2_SRFACILITY",
            "Fort Point": "Row3_SRFACILITY",
            "Government Center": "Row4_SRFACILITY",
            "Kendall Square": "Row5_SRFACILITY",
            "Kenmore Square": "Row6_SRFACILITY",
            "Leominster": "Row7_SRFACILITY",
            "Needham": "Row8_SRFACILITY",
            "Peabody": "Row9_SRFACILITY",
            "Post Office Square": "Row10_SRFACILITY",
            "Prudential Center": "Row11_SRFACILITY",
            "Quincy": "Row12_SRFACILITY",
            "Wayland": "Row13_SRFACILITY",
            "Wellesley": "Row14_SRFACILITY",
            # Add other facilities here if the list changes
        }

        # Look up the internal ID for the given facility name.
        target_item_id = facility_map.get(facility_name)
        if not target_item_id:
            self.logger.error(f"Facility '{facility_name}' not found in the internal facility_map. Please update the script.")
            return False
        
        # This entire section for finding the window handle is preserved.
        desk = Desktop(backend="uia")
        facility_handle = self.last_facility_handle
        if not facility_handle:
            self.logger.info("No remembered facility handle; trying to locate it again...")
            for w in desk.windows(process=self.app.process):
                if "services rendered facility" in (w.window_text() or "").lower():
                    facility_handle = w.handle
                    break
        if not facility_handle:
            self.logger.error("Could not find 'Services Rendered Facility' window.")
            return False

        facility_spec = self.app.window(handle=facility_handle)
        facility_win = facility_spec.wrapper_object()
        self.logger.info(f"'Services Rendered Facility' window ready: handle={facility_handle:#x}")
        facility_win.set_focus()
        time.sleep(0.5)

        # --- THE FINAL, CORRECTED SELECTION LOGIC ---
        try:
            # Find the DataItem directly using its unique internal title (ID).
            self.logger.info(f"Searching for facility item with internal ID: '{target_item_id}'...")
            facility_item = facility_spec.child_window(
                title=target_item_id,
                control_type="DataItem"
            )
            facility_item.wait('visible', timeout=5)

            # Click the item.
            self.logger.info(f"Clicking the row for '{facility_name}' (ID: {target_item_id})...")
            facility_item.click_input()
            time.sleep(0.5)

        except Exception as e:
            self.logger.error(f"Failed to click item with ID '{target_item_id}': {e}", exc_info=True)
            return False
        # --- END OF DEFINITIVE LOGIC ---

        # Click OK button (This part was already correct).
        try:
            self.logger.info("Looking for 'OK' button...")
            ok_button = facility_spec.child_window(title="OK", control_type="Button")
            ok_button.click_input()
            self.logger.info("Clicked 'OK' on facility dialog.")
        except Exception as e:
            self.logger.error(f"Could not click OK on facility dialog: {e}", exc_info=True)
            return False

        time.sleep(3.0)
        self.logger.info(f"Facility '{facility_name}' has been selected successfully.")
        return True
    

     # Placeholder for the rest of your automation workflow
    def run_single_patient_export(self, patient_last_name):
        self.logger.info(f"--- Starting data export for patient: {patient_last_name} ---")
        # 1. Navigate to scheduling (if not already there)
        # 2. Search for the patient
        # 3. Extract demographics
        # 4. Extract case/insurance info
        # 5. Download documents
        # ... etc.
        self.logger.info(f"--- Placeholder: Completed data export for {patient_last_name} ---")
        time.sleep(2) # Simulate work

    def connect_to_running_app(self):
        """
        Attach to an already-running TheraOffice instance.
        Use this when a human has launched & logged in (or is logging in) manually.
        """
        self.logger.info("Trying to connect to existing TheraOffice.exe process...")
        try:
            self.app = Application(backend="uia").connect(
                path=self.process_name,
                timeout=self.timeout
            )
            self.logger.info("Successfully connected to existing TheraOffice instance.")
            return True
        except Exception as e:
            self.logger.error(f"Could not connect to running TheraOffice.exe. Error: {e}")
            return False

    import time

    def wait_until_logged_in(self, max_wait_seconds=120):
        """
        Polls for the logged-in state using an unambiguous search (class and title)
        to correctly handle applications that create multiple top-level windows.
        """
        self.logger.info("Waiting for a logged-in state (final, unambiguous check)...")
        start = time.time()
        desk = Desktop(backend="uia")

        main_window_class = "WindowsForms10.Window.8.app.0.1629f15_r7_ad1"

        while time.time() - start < max_wait_seconds:
            try:
                # --- THE FINAL, DEFINITIVE FIX ---
                # We now search for the window using BOTH class_name AND title_re.
                # This is unambiguous and will only ever find the one, correct window.
                main_app_window = desk.window(
                    class_name=main_window_class,
                    title_re=".*TheraOffice.*"
                )

                if main_app_window.exists():
                    # --- 1. Check for the facility dialog as a child ---
                    facility_dialog = main_app_window.child_window(title="Services Rendered Facility")
                    if facility_dialog.exists() and facility_dialog.is_visible():
                        self.logger.info("SUCCESS: Detected 'Services Rendered Facility' child dialog.")
                        self.last_facility_handle = facility_dialog.handle
                        return "facility"

                    # --- 2. If no dialog, check if the main window is active ---
                    if main_app_window.is_enabled() and "(" in main_app_window.window_text():
                        self.main_window = main_app_window
                        self.logger.info(f"SUCCESS: Detected active main window: '{main_app_window.window_text()}'")
                        return "main"

            except Exception as e:
                # This can happen if the window search is ambiguous for a moment.
                self.logger.debug(f"Polling... waiting for windows to stabilize. Error: {e}")

            self.logger.debug("Polling for an active window (facility or main)...")
            time.sleep(2)

        self.logger.error("Timed out waiting for an active logged-in state.")
        return None
    
    def get_main_window(self):
        """
        Resolves ambiguity when multiple TheraOffice Web windows exist.
        Returns the real interactive main window.
        """
        try:
            # Grab all windows matching the regex
            windows = self.app.windows(title_re=r"TheraOffice Web.*")

            if len(windows) == 1:
                return windows[0]

            # If more than one: choose the visible, enabled, non-offscreen window
            candidates = []
            for w in windows:
                try:
                    if w.is_visible() and w.is_enabled():
                        rect = w.rectangle()
                        if rect.width() > 100 and rect.height() > 100:  # exclude tiny container windows
                            candidates.append(w)
                except Exception:
                    pass

            if not candidates:
                self.logger.error("No suitable visible main window found among candidates.")
                return None

            # If still >1, try choosing the one that can get focus
            for w in candidates:
                try:
                    w.set_focus()
                    return w
                except Exception:
                    continue

            # fallback: return the first visible candidate
            return candidates[0]

        except Exception as e:
            self.logger.error(f"Error resolving main window: {e}")
            return None
    def debug_dump_windows(self, note=""):
        """Log all top-level windows for this process (for debugging)."""
        try:
            desk = Desktop(backend="uia")
            wins = desk.windows(process=self.app.process)
            self.logger.debug(f"[WIN-DUMP] {note} - Found {len(wins)} top-level windows for process {self.app.process}")
            for idx, w in enumerate(wins):
                try:
                    rect = w.rectangle()
                    self.logger.debug(
                        f"[WIN-DUMP] {idx}: title='{w.window_text()}', "
                        f"handle=0x{w.handle:x}, visible={w.is_visible()}, enabled={w.is_enabled()}, "
                        f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom})"
                    )
                except Exception as e:
                    self.logger.debug(f"[WIN-DUMP] {idx}: error reading window properties: {e}")
        except Exception as e:
            self.logger.debug(f"[WIN-DUMP] Failed to dump windows: {e}")
    def debug_inspect_main_window(self):
        """
        Dump the main window's control tree to stdout/log for debugging.
        """
        if not self.main_window:
            self.logger.error("debug_inspect_main_window called but main_window is None.")
            return

        self.logger.info("=== BEGIN MAIN WINDOW CONTROL TREE DUMP ===")
        try:
            spec = self.app.window(handle=self.main_window.handle)
            spec.print_control_identifiers()
        except Exception as e:
            self.logger.error(f"Failed to print control identifiers: {e}", exc_info=True)
        self.logger.info("=== END MAIN WINDOW CONTROL TREE DUMP ===")
    
    
    def dismiss_shared_user_accounts_warning(self, timeout=30):
        """
        Aggressively attempts to dismiss the 'Shared User Accounts' popup using
        multiple strategies until it is confirmed to be gone.
        """
        self.logger.info("Scanning for 'Shared User Accounts' popup...")
        from pywinauto import Desktop
        from pywinauto.keyboard import send_keys
        import time

        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # 1. Check if the popup exists
                popup = Desktop(backend="uia").window(auto_id="frmSharedAccounts")
                
                if not popup.exists() or not popup.is_visible():
                    # If popup is gone, check if main window is enabled to be sure
                    main_window_class = "WindowsForms10.Window.8.app.0.1629f15_r7_ad1"
                    main_win = Desktop(backend="uia").window(class_name=main_window_class)
                    if main_win.exists() and main_win.is_enabled():
                        self.logger.info("Popup is gone and main window is enabled.")
                        return True
                    else:
                        # Popup might be gone, but main window still disabled? Wait a bit.
                        time.sleep(0.5)
                        continue

                self.logger.info("Popup found. Attempting to dismiss...")

                # --- STRATEGY 1: Click OK Button ---
                try:
                    ok_btn = popup.child_window(auto_id="SimpleButtonOK", control_type="Button")
                    if ok_btn.exists():
                        self.logger.info("Strategy 1: Clicking 'OK' button...")
                        ok_btn.set_focus()
                        ok_btn.click_input()
                except: pass

                # --- STRATEGY 2: Send ENTER ---
                try:
                    self.logger.info("Strategy 2: Sending ENTER...")
                    popup.set_focus()
                    send_keys('{ENTER}')
                except: pass

                # --- STRATEGY 3: Click Close (X) ---
                try:
                    close_btn = popup.child_window(title="Close", control_type="Button")
                    if close_btn.exists():
                        self.logger.info("Strategy 3: Clicking 'Close' button...")
                        close_btn.click_input()
                except: pass
                
                # Wait a moment to see if it worked
                time.sleep(1)

            except Exception as e:
                self.logger.warning(f"Error during popup handling: {e}")
                time.sleep(1)

        self.logger.warning("Timeout reached. Popup may still be present.")
        return False


    def find_window(self, name=None, automation_id=None, class_name=None):
        for w in pywinauto.findwindows.find_elements():
            if name and name.lower() not in (w.name or "").lower():
                continue
            if automation_id and automation_id != w.automation_id:
                continue
            if class_name and class_name.lower() not in (w.class_name or "").lower():
                continue
            return w
        return None
    def click_ok_on_window(self, window):
        app = pywinauto.Application(backend="uia").connect(process=window.process_id)
        w = app.window(handle=window.handle)
        ok = w.child_window(auto_id="SimpleButtonOK", control_type="Button")
        ok.click_input()

    def wait_and_dismiss_popup(self, title="Shared User Accounts", timeout=40):
        self.logger.info(f"Waiting for popup: {title}")
        deadline = time.time() + timeout

        while time.time() < deadline:
            win = self.find_window(name=title)
            if win:
                self.logger.info(f"Popup '{title}' found. Clicking OK.")
                self.click_ok_on_window(win)
                return True
            time.sleep(0.5)

        self.logger.warning(f"Popup '{title}' not found in time.")
        return False

    
    
    def navigate_to_scheduling(self) -> bool:
        """
        Navigates to the Scheduling module by finding the 'Modules' toolbar
        and then clicking the 'Scheduling' button specifically within that toolbar.
        """
        self.logger.info("Navigating to the Scheduling module...")
        try:
            if not self.main_window or not self.main_window.exists():
                self.logger.error("Main window not found. Cannot navigate to Scheduling.")
                return False

            # --- THE FINAL, DEFINITIVE FIX ---
            # 1. First, find the parent toolbar named "Modules".
            self.logger.info("Finding the 'Modules' toolbar...")
            modules_toolbar = self.main_window.child_window(
                title="Modules",
                control_type="ToolBar"
            )
            modules_toolbar.wait('visible', timeout=10)

            # 2. Then, find the "Scheduling" button INSIDE that specific toolbar.
            self.logger.info("Finding the 'Scheduling' button within the 'Modules' toolbar...")
            scheduling_button = modules_toolbar.child_window(
                title="Scheduling",
                control_type="Button"
            )
            scheduling_button.wait('visible', timeout=5)
            
            # 3. Click the button.
            self.logger.info("Clicking the 'Scheduling' button...")
            scheduling_button.click_input()
            
            self.logger.info("Successfully navigated to Scheduling.")
            time.sleep(3) # Wait for the scheduling view to load
            return True

        except Exception as e:
            self.logger.error(f"Failed to navigate to Scheduling. Error: {e}", exc_info=True)
            return False
        
    def search_for_patient(self, patient_name: str) -> bool:
        """
        Finds the patient search box, clicks it to focus, and types the patient's name.
        This method uses keyboard simulation which is more reliable for search bars.
        """
        self.logger.info(f"Searching for patient: '{patient_name}'...")
        try:
            if not self.main_window or not self.main_window.exists():
                self.logger.error("Main window not found. Cannot search for patient.")
                return False

            # Find the search edit box by its title "Search:"
            # We use the 'Edit' control type as identified in previous logs.
            search_box = self.main_window.child_window(
                title="Search:",
                control_type="Edit"
            )
            search_box.wait('visible', timeout=10)
            
            # --- THE FIX: CLICK AND TYPE ---
            # 1. Click to focus the search box
            self.logger.info("Clicking search box to focus...")
            search_box.click_input()
            time.sleep(0.5)
            
            # 2. Type the name using keyboard simulation
            self.logger.info(f"Typing '{patient_name}'...")
            search_box.type_keys(patient_name, with_spaces=True)
            time.sleep(0.5)
            
            # 3. Press Enter to search
            self.logger.info("Pressing Enter...")
            search_box.type_keys('{ENTER}')
            
            self.logger.info("Patient search initiated.")
            time.sleep(3) # Wait for search results to load
            return True

        except Exception as e:
            self.logger.error(f"Failed to search for patient. Error: {e}", exc_info=True)
            return False


    def debug_scheduling_screen(self):
        """
        DIAGNOSTIC: Prints all controls on the Scheduling screen so we can
        find the 'Facility' button or dropdown.
        """
        self.logger.info("--- STARTING SCHEDULING SCREEN DIAGNOSTIC ---")
        time.sleep(3) # Wait for screen to fully render
        try:
            # Print the control identifiers to the log
            self.main_window.print_control_identifiers()
            self.logger.info("--- FINISHED SCHEDULING SCREEN DIAGNOSTIC ---")
            self.logger.error("Please copy the 'Control Identifiers' output from the terminal.")
        except Exception as e:
            self.logger.error(f"Diagnostic failed: {e}")


    def handle_scheduling_facility(self, facility_name: str = "Brookline/Allston") -> bool:
        """
        Handles the 'Select Facility' dialog in Scheduling using the robust
        Name-to-ID mapping, which is proven to work for this application's tables.
        """
        self.logger.info(f"Checking for 'Select Facility' dialog in Scheduling...")
        from pywinauto import Desktop
        import time

        # --- REUSE THE PROVEN MAPPING ---
        facility_map = {
            "Brookline/Allston": "Row0_SRFACILITY",
            "Concord": "Row1_SRFACILITY",
            "Downtown": "Row2_SRFACILITY",
            "Fort Point": "Row3_SRFACILITY",
            "Government Center": "Row4_SRFACILITY",
            "Kendall Square": "Row5_SRFACILITY",
            "Kenmore Square": "Row6_SRFACILITY",
            "Leominster": "Row7_SRFACILITY",
            "Needham": "Row8_SRFACILITY",
            "Peabody": "Row9_SRFACILITY",
            "Post Office Square": "Row10_SRFACILITY",
            "Prudential Center": "Row11_SRFACILITY",
            "Quincy": "Row12_SRFACILITY",
            "Wayland": "Row13_SRFACILITY",
            "Wellesley": "Row14_SRFACILITY",
        }

        target_item_id = facility_map.get(facility_name)
        if not target_item_id:
            self.logger.error(f"Facility '{facility_name}' not found in map.")
            return False

        try:
            # 1. Find the dialog.
            if not self.main_window: return False
            
            # Search for the dialog as a child of the main window
            facility_dialog = self.main_window.child_window(title="Select Facility", control_type="Window")
            
            if not facility_dialog.exists(timeout=5):
                self.logger.info("'Select Facility' dialog did not appear. Continuing...")
                return True

            self.logger.info("'Select Facility' dialog found. Selecting facility...")
            facility_dialog.set_focus()

            # 2. Find the table control.
            table = facility_dialog.child_window(auto_id="fg", control_type="Table")
            table.wait('visible', timeout=5)

            # 3. Find the item directly using the ID from the map.
            self.logger.info(f"Searching for facility item with ID: '{target_item_id}'...")
            facility_item = facility_dialog.child_window(
                title=target_item_id,
                control_type="DataItem"
            )
            facility_item.wait('visible', timeout=5)

            # 4. Click the item.
            self.logger.info(f"Clicking the row for '{facility_name}'...")
            facility_item.click_input()
            time.sleep(0.5)

            # 5. Click OK.
            ok_button = facility_dialog.child_window(title="OK", control_type="Button")
            ok_button.click_input()
            self.logger.info("Clicked 'OK' on Select Facility dialog.")
            time.sleep(2)
            return True

        except Exception as e:
            self.logger.error(f"Error handling Scheduling facility selection: {e}")
            return False