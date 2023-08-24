import pywinauto
from pywinauto import Application

def click_duplicate_button(app_path, main_window_title_regex, child_window_title, button_id):
    """Clicks the 'Duplicate Invoice' button in QuickBooks."""
    
    # Connect to the application
    app = Application(backend="uia").connect(path=app_path)
    
    # Check if app is running
    if not app.is_process_running():
        print(f"Unable to connect to {app_path}")
        return

    # Find the main window
    main_window = app.window(title_re=main_window_title_regex)
    if not main_window.exists():
        print("Main window does not exist.")
        return

    # Navigate to the 'Duplicate Invoice' button
    duplicate_invoice_btn = main_window.child_window(auto_id="DuplicateBtn")


    # Check if the button is found and click it
    if duplicate_invoice_btn.exists():
        print("Found 'Duplicate Invoice' button!")
        duplicate_invoice_btn.click_input()
    else:
        print("'Duplicate Invoice' button not found!")

# Usage
if __name__ == "__main__":
    APP_PATH = r"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32.EXE"
    MAIN_WINDOW_TITLE_REGEX = ".*QuickBooks Desktop Pro 2019.*"
    CHILD_WINDOW_TITLE = "Create Invoices (Editing Transaction...) "
    BUTTON_ID = "DuplicateBtn"
    
    click_duplicate_button(APP_PATH, MAIN_WINDOW_TITLE_REGEX, CHILD_WINDOW_TITLE, BUTTON_ID)
