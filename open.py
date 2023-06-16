import subprocess
import time
import pyautogui

def open_outlook_app(search_query):
    try:
        outlook_exe_path = r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"  # Replace with the correct path to your Outlook executable
        process = subprocess.Popen([outlook_exe_path])
        print("Outlook app is now open.")

        # Wait for one minute
        time.sleep(10)

        # Locate the search box
        search_box_position = pyautogui.locateOnScreen('search_box.png')
        if search_box_position is None:
            print("Search box not found.")
            return

        # Click on the search box to focus it
        search_box_center = pyautogui.center(search_box_position)
        pyautogui.click(search_box_center)

        # Type the search query
        pyautogui.typewrite(search_query)

        # Press Enter to perform the search
        pyautogui.press('enter')

        # Wait for the search results to load
        time.sleep(45)

        # Get the screen coordinates of the top right corner
        screen_width, screen_height = pyautogui.size()
        exit_button_x = screen_width - 10  # Adjust this value as needed
        exit_button_y = 10  # Adjust this value as needed

        # Move the mouse to the exit button and click it
        pyautogui.moveTo(exit_button_x, exit_button_y)
        pyautogui.click()

        # Wait for the process to exit or forcefully terminate it after a timeout
        timeout = 10  # Timeout in seconds
        start_time = time.time()
        while time.time() - start_time < timeout:
            if process.poll() is not None:
                print("Outlook app has been gracefully closed.")
                return
            time.sleep(1)

        # Forcefully terminate the process if it hasn't exited gracefully
        process.terminate()
        process.kill()
        print("Outlook app has been forcefully closed.")
    except Exception as e:
        print("An error occurred while trying to open or close Outlook.")
        print(e)

# Provide the search query
search_query = 'Your search query'  # Replace with your specific search query

open_outlook_app(search_query)

