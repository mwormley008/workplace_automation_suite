import pywinauto
from pywinauto import Application
from pywinauto.timings import Timings

Timings.window_find_timeout = .1  # Example value, adjust as necessary
Timings.window_find_retry = 0.3  # Example value, adjust as necessary


def click_duplicate_button(app_path, main_window_title_regex, child_window_title, button_id):
    """Clicks the 'Duplicate Invoice' button in QuickBooks."""
    
    # Connect to the application
    app = Application(backend="uia").connect(path=app_path)
    
    # # Check if app is running
    # if not app.is_process_running():
    #     print(f"Unable to connect to {app_path}")
    #     return

    # # Find the main window
    main_window = app.window(title_re=main_window_title_regex)
    # if not main_window.exists():
    #     print("Main window does not exist.")
    #     return

    # Navigate to the 'Duplicate Invoice' button
    duplicate_invoice_btn = main_window.child_window(auto_id="DuplicateBtn")
    duplicate_invoice_btn.click_input()

    # Check if the button is found and click it
    # if duplicate_invoice_btn.exists():
    #     print("Found 'Duplicate Invoice' button!")
    #     
    # else:
    #     print("'Duplicate Invoice' button not found!")

# Usage
if __name__ == "__main__":
    APP_PATH = r"C:\Program Files\Intuit\QuickBooks 2024\QBW.EXE"
    MAIN_WINDOW_TITLE_REGEX = ".*QuickBooks Desktop Pro*"
    CHILD_WINDOW_TITLE = "Create Invoices (Editing Transaction...) "
    BUTTON_ID = "DuplicateBtn"
    
    click_duplicate_button(APP_PATH, MAIN_WINDOW_TITLE_REGEX, CHILD_WINDOW_TITLE, BUTTON_ID)
