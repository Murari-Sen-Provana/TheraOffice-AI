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
        Use the previously-detected facility dialog (self.last_facility_handle)
        to select Brookline/Allston and click OK.
        """
        from pywinauto import Desktop
        import time

        self.logger.info(f"Attempting to select facility '{facility_name}'...")

        desk = Desktop(backend="uia")

        facility_handle = self.last_facility_handle

        if facility_handle:
            self.logger.info(f"Using remembered facility handle 0x{facility_handle:x}")
        else:
            self.logger.info("No remembered facility handle; trying to locate it again...")

        # If we don't have a handle yet, try to find it again (top-level or child)
        if not facility_handle:
            for w in desk.windows(process=self.app.process):
                title = (w.window_text() or "").lower()
                if "services rendered facility" in title:
                    facility_handle = w.handle
                    self.logger.info(
                        f"Found facility window as top-level: title='{w.window_text()}', handle=0x{w.handle:x}"
                    )
                    break

            if not facility_handle and self.main_window is not None:
                try:
                    main_spec = self.app.window(handle=self.main_window.handle)
                    facility_spec = main_spec.child_window(title_re=".*Services Rendered Facility.*")
                    self.logger.info("Re-scanning main window for facility dialog as child...")
                    if facility_spec.exists(timeout=2):
                        facility_handle = facility_spec.wrapper_object().handle
                        self.logger.info(
                            f"Found facility window as child of main window: handle=0x{facility_handle:x}"
                        )
                except Exception as e:
                    self.logger.error(f"Error locating facility as child of main window: {e}", exc_info=True)

        if not facility_handle:
            self.logger.error("Could not find 'Services Rendered Facility' window to select facility.")
            return False

        # Wrap handle into WindowSpecification / UIAWrapper
        facility_spec = self.app.window(handle=facility_handle)
        facility_win = facility_spec.wrapper_object()

        self.logger.info(
            f"'Services Rendered Facility' window ready: handle=0x{facility_handle:x}, "
            f"title='{facility_win.window_text()}'"
        )

        try:
            facility_win.set_focus()
            facility_win.set_focus()  # call twice just in case
            self.logger.info("Facility dialog focused.")
        except Exception as e:
            self.logger.warning(f"Could not explicitly focus facility dialog: {e}")

        time.sleep(0.5)

        # Try to click the Brookline/Allston row by internal name Row0_SRFACILITY
        try:
            self.logger.info("Searching for row 'Row0_SRFACILITY' (Brookline/Allston)...")
            row_spec = facility_spec.child_window(
                title="Row0_SRFACILITY",
                control_type="DataItem"
            )
            if not row_spec.exists(timeout=3):
                self.logger.error("DataItem 'Row0_SRFACILITY' not found in facility dialog.")
                return False

            row = row_spec.wrapper_object()
            rect = row.rectangle()
            self.logger.info(
                f"Clicking facility row Row0_SRFACILITY at "
                f"({rect.left},{rect.top},{rect.right},{rect.bottom})"
            )
            row.click_input()
            time.sleep(0.5)
        except Exception as e:
            self.logger.error(f"Failed to click facility row Row0_SRFACILITY: {e}", exc_info=True)
            return False

        # Click OK button
        try:
            self.logger.info("Looking for 'OK' button in facility dialog...")
            ok_spec = facility_spec.child_window(title="OK", control_type="Button")
            if not ok_spec.exists(timeout=3):
                self.logger.error("OK button not found in facility dialog.")
                return False

            ok_btn = ok_spec.wrapper_object()
            self.logger.info("Clicking 'OK' on facility dialog...")
            ok_btn.click_input()
        except Exception as e:
            self.logger.error(f"Could not click OK on facility dialog: {e}", exc_info=True)
            return False

        time.sleep(2.0)
        self.logger.info(f"Facility '{facility_name}' should now be selected.")
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
        Polls for:
          - 'Services Rendered Facility' dialog  -> returns 'facility'
          - main 'TheraOffice Web ( ... )' window -> returns 'main'
        """
        self.logger.info("Waiting for a logged-in state (facility dialog or main window)...")
        start = time.time()
        desk = Desktop(backend="uia")

        self.last_facility_handle = None

        while time.time() - start < max_wait_seconds:
            try:
                wins = desk.windows(process=self.app.process)
                self.logger.debug(f"[WIN-DUMP] poll - Found {len(wins)} top-level windows for process {self.app.process}")
                for idx, w in enumerate(wins):
                    try:
                        rect = w.rectangle()
                        self.logger.debug(
                            f"[WIN-DUMP] {idx}: title='{w.window_text()}', handle=0x{w.handle:x}, "
                            f"visible={w.is_visible()}, enabled={w.is_enabled()}, "
                            f"rect=({rect.left},{rect.top},{rect.right},{rect.bottom})"
                        )
                    except Exception as e:
                        self.logger.debug(f"[WIN-DUMP] {idx}: error reading window properties: {e}")

                # ---------- 1) Try to find facility as a child of each top-level window ----------
                for w in wins:
                    spec = self.app.window(handle=w.handle)  # WindowSpecification
                    try:
                        facility_spec = spec.child_window(
                            title_re=".*Services Rendered Facility.*"
                        )
                        if facility_spec.exists(timeout=0.2):
                            facility_wrapper = facility_spec.wrapper_object()
                            if facility_wrapper.is_visible():
                                self.last_facility_handle = facility_wrapper.handle
                                self.logger.info(
                                    f"Detected facility selection dialog as child of window "
                                    f"'{w.window_text()}' handle=0x{w.handle:x}"
                                )
                                return "facility"
                    except Exception as e:
                        self.logger.debug(f"Child facility search failed in window 0x{w.handle:x}: {e}")

                # ---------- 2) If no facility, look for main 'TheraOffice Web ...' window ----------
                for w in wins:
                    title = (w.window_text() or "").lower()
                    if title.startswith("theraoffice web"):
                        try:
                            rect = w.rectangle()
                            vis = w.is_visible()
                            en = w.is_enabled()
                            self.logger.debug(
                                f"[MAIN-CAND] title='{title}', visible={vis}, "
                                f"enabled={en}, size=({rect.width()}x{rect.height()})"
                            )
                            if vis and rect.width() > 100 and rect.height() > 100:
                                try:
                                    w.set_focus()
                                except Exception as e:
                                    self.logger.debug(f"Could not focus main candidate window: {e}")
                                self.main_window = w  # UIAWrapper
                                self.logger.info(
                                    f"Detected main TheraOffice window (logged-in state) "
                                    f"(enabled={en}): '{w.window_text()}' handle=0x{w.handle:x}"
                                )
                                return "main"
                        except Exception as e:
                            self.logger.debug(f"Main candidate check failed: {e}")
            except Exception as e:
                self.logger.warning(f"Error while polling for logged-in state: {e}")

            time.sleep(2)

        self.logger.error("Timed out waiting for a logged-in state. User may not have finished login.")
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
    
    def dismiss_shared_user_accounts_warning(self, timeout=40):
        """
        Dismiss the Shared User Accounts popup (OK button).
        Uses deep UIA recursive search so it works even on WinForms/DevExpress dialogs.
        """
        self.logger.info("Scanning for Shared User Accounts OK button...")

        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                # Enumerate ALL top-level windows in the app
                for win in self.app.windows():
                    try:
                        # Try direct search first
                        ok = win.child_window(auto_id="SimpleButtonOK", control_type="Button")
                        if ok.exists() and ok.is_visible():
                            self.logger.info(f"Found OK button in window '{win.window_text()}'. Clicking...")
                            try:
                                ok.invoke()
                            except:
                                ok.click_input()
                            time.sleep(1)
                            self.logger.info("OK popup dismissed.")
                            return True
                    except:
                        pass

                    # Deep recursive UIA search
                    try:
                        descendants = win.descendants()
                        for d in descendants:
                            try:
                                if d.control_type == "Button" and d.element_info.automation_id == "SimpleButtonOK":
                                    self.logger.info(
                                        f"Found OK button deep inside window '{win.window_text()}'. Clicking..."
                                    )
                                    try:
                                        d.invoke()
                                    except:
                                        d.click_input()

                                    time.sleep(1)
                                    self.logger.info("OK popup dismissed.")
                                    return True
                            except:
                                continue
                    except:
                        continue

            except Exception as e:
                self.logger.error(f"Error scanning windows for OK popup: {e}")

            time.sleep(0.3)

        self.logger.warning("Did NOT detect Shared User Accounts popup within timeout.")
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

