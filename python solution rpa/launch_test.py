# launch_test.py
import subprocess
import time
import os

# --- Paste the exact details from your config.ini here ---
APP_FAMILY_PACKAGE_NAME = "NetsmartTechnologies.TheraOfficeWeb_qxsw3pe1m8nt6"
APP_ID = "NetsmartTechnologies.TheraOfficeWeb"
# ---------------------------------------------------------

# This forms the unique ID for the UWP app
UWP_APP_ID = f"{APP_FAMILY_PACKAGE_NAME}!{APP_ID}"

def test_method_1_explorer_shell():
    """The method we have been trying. For baseline comparison."""
    print("\n--- [METHOD 1] Trying to launch via 'explorer.exe shell:appsFolder'...")
    try:
        command = f'explorer.exe shell:appsFolder\\{UWP_APP_ID}'
        print(f"Executing: {command}")
        subprocess.run(command, shell=True, check=False)
        print("Command sent. Check if TheraOffice launched or if a folder opened.")
    except Exception as e:
        print(f"Method 1 failed with an error: {e}")

def test_method_2_powershell():
    """
    A more direct and often more reliable method using PowerShell.
    This is the most likely to succeed on a restricted system.
    """
    print("\n--- [METHOD 2] Trying to launch via PowerShell 'start' command...")
    try:
        # PowerShell command to get the AppX package and pipe it to start
        command = f'powershell.exe -Command "Get-AppxPackage {APP_FAMILY_PACKAGE_NAME} | ForEach-Object {{ Start-Process -FilePath shell:appsfolder\\$($_.PackageFamilyName)!$($_.Name) }}"'
        print(f"Executing PowerShell command...")
        subprocess.run(command, shell=True, check=False)
        print("Command sent. Check if TheraOffice launched.")
    except Exception as e:
        print(f"Method 2 failed with an error: {e}")

def test_method_3_os_startfile():
    """
    A high-level Python method. Less likely to work, but easy to test.
    """
    print("\n--- [METHOD 3] Trying to launch via Python's 'os.startfile'...")
    try:
        uri = f'shell:appsFolder\\{UWP_APP_ID}'
        print(f"Executing os.startfile with URI: {uri}")
        os.startfile(uri)
        print("Command sent. Check if TheraOffice launched.")
    except Exception as e:
        print(f"Method 3 failed with an error: {e}")


if __name__ == "__main__":
    print("Starting UWP Launch Diagnostic Script...")
    print("We will test 3 different methods to launch TheraOffice.")
    print("Please observe what happens after each method is tried.")

    # Test Method 1
    test_method_1_explorer_shell()
    time.sleep(5)  # Wait 5 seconds to see the result

    # Test Method 2
    test_method_2_powershell()
    time.sleep(5)  # Wait 5 seconds to see the result

    # Test Method 3
    test_method_3_os_startfile()

    print("\nDiagnostic script finished. Please report which method (if any) worked.")