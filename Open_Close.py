import subprocess
import time
import psutil
import os


def open_outlook_app():
    try:
        outlook_exe_path = r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"  # Replace with the correct path to your Outlook executable
        process = subprocess.Popen([outlook_exe_path])
        print("Outlook app is now open.")

        # Wait for one minute
        time.sleep(30)
        # Forcefully terminate the process if it hasn't exited gracefully
        for child in psutil.Process(process.pid).children(recursive=True):
            child.kill()
        process.kill()
        print("Outlook app has been forcefully closed.")

        # Close the command prompt window
        print("Command prompt window has been closed.")
        time.sleep(0.8)
        os.system("taskkill /f /im cmd.exe")
        
    except Exception as e:
        print("An error occurred while trying to open or close Outlook.")
        print(e)


open_outlook_app()





















