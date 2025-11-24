# bruteforce_launcher.py

import os
import time
import psutil
import subprocess
import pyautogui
import pygetwindow as gw
import pytesseract
from PIL import Image
import cv2
import numpy as np


# --- Global defaults (can be overridden by config / env) ---

# Configure Tesseract OCR path as needed
TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Fallbacks; we'll override from config if available
DEFAULT_EXE_PATH = r"C:\Program Files\WindowsApps\NetsmartTechnologies.TheraOfficeWeb_14.42.0.0_x86__qxsw3pe1m8nt6\VFS\ProgramFilesX86\NetSmart Technologies\TheraOffice Web\TheraOffice.exe"
DEFAULT_PROCESS_NAME = "TheraOffice.exe"
DEFAULT_ORGANIZATION = "Joint Ventures"

# Hard-coded coordinates (must be calibrated for your machine)
USERNAME_FIELD_X, USERNAME_FIELD_Y = 1133, 691
PASSWORD_FIELD_X, PASSWORD_FIELD_Y = 1131, 762
LOGIN_BUTTON_X, LOGIN_BUTTON_Y = 1107, 905
USERNAME_REGION = (1115, 685, 220, 35)  # (left, top, width, height)

pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True


def _log(logger, msg):
    if logger:
        logger.info(msg)
    else:
        print(msg)


def configure_from_config(config, logger=None):
    """Build a simple settings dict from config.ini (with fallbacks)."""
    section_to = config["TheraOffice"]
    exe_path = section_to.get("executable_path", DEFAULT_EXE_PATH)
    process_name = section_to.get("process_name", DEFAULT_PROCESS_NAME)

    username = section_to.get("username", os.getenv("THERA_USERNAME", "megbus"))
    password = section_to.get("password", os.getenv("THERA_PASSWORD", "Jvpt2025!"))

    organization = section_to.get("organization", DEFAULT_ORGANIZATION)

    pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE

    _log(logger, f"Brute-force launcher configured. EXE={exe_path}, USER={username}, ORG={organization}")

    return {
        "exe_path": exe_path,
        "process_name": process_name,
        "username": username,
        "password": password,
        "organization": organization,
    }


def kill_theraoffice_processes(process_name, logger=None):
    _log(logger, f"Killing existing {process_name} processes...")
    for proc in psutil.process_iter(["name"]):
        try:
            if proc.info["name"] and proc.info["name"].lower() == process_name.lower():
                _log(logger, f"Terminating PID {proc.pid} ({proc.info['name']})")
                proc.terminate()
                proc.wait(timeout=5)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.TimeoutExpired):
            continue


def launch_theraoffice(exe_path, logger=None):
    try:
        subprocess.Popen(exe_path, shell=True)
        _log(logger, "Launched TheraOffice. Waiting 10 seconds for UI to appear...")
        time.sleep(10)
        return True
    except FileNotFoundError:
        _log(logger, f"Could not find the TheraOffice executable at {exe_path}")
        return False


def print_mouse_position(label="Position", logger=None):
    pos = pyautogui.position()
    _log(logger, f"{label}: X={pos.x}, Y={pos.y}")


def extract_text_from_region(region, logger=None):
    screenshot = pyautogui.screenshot(region=region)
    img = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.medianBlur(gray, 3)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    temp_file = "temp_user.png"
    cv2.imwrite(temp_file, thresh)
    text = pytesseract.image_to_string(Image.open(temp_file), config="--psm 7").strip().lower()
    os.remove(temp_file)
    _log(logger, f"OCR read from username field: '{text}'")
    return text


def handle_username_field(username, logger=None):
    pyautogui.moveTo(USERNAME_FIELD_X, USERNAME_FIELD_Y, duration=0.3)
    pyautogui.click(USERNAME_FIELD_X, USERNAME_FIELD_Y)
    print_mouse_position("Username Clicked", logger=logger)
    time.sleep(0.5)

    current_username = extract_text_from_region(USERNAME_REGION, logger=logger)
    if current_username == username.lower():
        _log(logger, "Username already correct; tabbing to password.")
        pyautogui.press("tab")
        time.sleep(1)
    else:
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.2)
        pyautogui.press("delete")
        _log(logger, "Cleared username; typing correct username.")
        pyautogui.write(username)
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(1)


def focus_window_containing(org, user, logger=None):
    """Finds window whose title contains both org and username (case insensitive) and activates it."""
    org_lower = org.lower()
    user_lower = user.lower()
    windows = gw.getAllWindows()
    for w in windows:
        title_lower = w.title.lower()
        if org_lower in title_lower and user_lower in title_lower:
            try:
                w.activate()
                _log(logger, f"Activated window: '{w.title}'")
                time.sleep(2)
                return True
            except Exception as e:
                _log(logger, f"Could not activate window '{w.title}': {e}")
    _log(logger, "No matching window found for organization + username.")
    return False


def wait_for_dashboard(org, user, timeout=60, poll_interval=1, min_height=800, logger=None):
    _log(logger, f"Waiting for main dashboard window for up to {timeout} seconds...")
    start_time = time.time()

    while time.time() - start_time < timeout:
        if focus_window_containing(org, user, logger=logger):
            windows = [
                w for w in gw.getAllWindows()
                if org.lower() in w.title.lower() and user.lower() in w.title.lower()
            ]
            if not windows:
                _log(logger, "Dashboard window temporarily missing, retrying...")
                time.sleep(poll_interval)
                continue
            window = windows[0]
            if window.height > min_height:
                try:
                    window.maximize()
                    _log(logger, f"Dashboard detected and maximized (size: {window.size})")
                    return True
                except Exception as e:
                    _log(logger, f"Error maximizing window: {e}")
                    return True
            else:
                _log(logger, f"Window found but height {window.height} < {min_height}, waiting...")
                time.sleep(poll_interval)
        else:
            _log(logger, "Dashboard window not found yet, waiting...")
            time.sleep(poll_interval)

    _log(logger, "Timed out waiting for dashboard window.")
    return False


def automate_login(settings, logger=None):
    org = settings["organization"]
    username = settings["username"]
    password = settings["password"]

    # Bring login window to front
    if not focus_window_containing(org, username, logger=logger):
        _log(logger, "Login window not focused; waiting briefly before continuing...")
        time.sleep(2)

    # Username
    handle_username_field(username, logger=logger)

    # Password
    pyautogui.moveTo(PASSWORD_FIELD_X, PASSWORD_FIELD_Y, duration=0.3)
    pyautogui.click(PASSWORD_FIELD_X, PASSWORD_FIELD_Y)
    print_mouse_position("Password Clicked", logger=logger)
    pyautogui.write(password)
    time.sleep(0.5)

    # Login button
    pyautogui.moveTo(LOGIN_BUTTON_X, LOGIN_BUTTON_Y, duration=0.3)
    pyautogui.click(LOGIN_BUTTON_X, LOGIN_BUTTON_Y)
    print_mouse_position("Login Clicked", logger=logger)
    _log(logger, "Login submitted; waiting for dashboard...")
    time.sleep(5)

    if wait_for_dashboard(org, username, logger=logger):
        _log(logger, "Dashboard ready after brute-force login.")
        return True
    else:
        _log(logger, "Failed to detect dashboard window after login.")
        return False


def launch_and_login_bruteforce(config, logger=None):
    """
    High-level helper:
      1) Kill existing TheraOffice.exe processes
      2) Launch app
      3) Perform brute-force login (visual/coordinate + OCR)
    """
    settings = configure_from_config(config, logger=logger)

    kill_theraoffice_processes(settings["process_name"], logger=logger)

    if not launch_theraoffice(settings["exe_path"], logger=logger):
        _log(logger, "Failed to launch TheraOffice; aborting brute-force login.")
        return False

    if not automate_login(settings, logger=logger):
        _log(logger, "Brute-force login did not succeed.")
        return False

    return True


if __name__ == "__main__":
    # Simple standalone run (without logger/config.ini integration)
    class DummyCfg:
        def __getitem__(self, key):
            return {}

    dummy_config = {"TheraOffice": {}}
    launch_and_login_bruteforce(dummy_config, logger=None)
