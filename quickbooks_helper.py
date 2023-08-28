import pywinauto
from pywinauto import Application, mouse
import time
from pywinauto.controls.win32_controls import ButtonWrapper
# from pywinauto import 


output_file = open("output.txt", "w")


def dual_print(message):
    print(message)              # Print to console
    print(message, file=output_file)   # Write to file

def print_element_info(element, indent=0):
    """Recursive function to print UI element info."""
    
    # Print current element info
    prefix = '  ' * indent  # Increase indentation based on depth
    dual_print(f"{prefix}{element.window_text()} - {element.element_info.control_type} - {element.class_name()}")
    
    # Recursive call for child elements
    for child in element.children():
        print_element_info(child, indent + 1)

def check_exist(element, element_name):
    if element.exists():
        dual_print(f"{element_name} exists.")
    else:
        dual_print(f"{element_name} does not exist.")

def find_message_window(main_window):
    try:
        message_window = main_window.child_window(class_name="MauuiMessage", title="QuickBooks Message")
        return message_window.exists(), message_window
    except Exception as e:
        print("Error finding message window:", str(e))
        return False


def click_duplicate_invoice_button(main_window):
    create_invoices_window_title = "Create Invoices (Editing Transaction...) "
    create_invoices_window = main_window.child_window(title=create_invoices_window_title)
    
    duplicate_invoice_btn = main_window.child_window(auto_id="DuplicateBtn")

    if duplicate_invoice_btn.exists():
        dual_print("Found 'Duplicate Invoice' button!")
        duplicate_invoice_btn.click_input()
        
        qb_wpf_container = app.window(auto_id="QB2WPFContainer")
        print(qb_wpf_container.children())

        buttons = [child for child in create_invoices_window.children() if isinstance(child, ButtonWrapper)]

        for button in buttons:
            print(button)
            button.click_input()
            print(button)
            time.sleep(4)  # Pause for 4 seconds to observe

        check_exist(duplicate_invoice_btn, 'create copy button')
        
        if duplicate_invoice_btn.exists():
            dual_print("Found 'Create a Copy' button!")
            duplicate_invoice_btn.click_input()
        else:
            dual_print("'Create a Copy' button not found!")
    else:
        dual_print("'Duplicate Invoice' button not found!")

def handle_error(message_window, main_window):
    cancel_button = message_window.child_window(title="Cancel")

    if cancel_button.exists():
        dual_print("Cancel button found!")
        cancel_button.click_input()
    else:
        dual_print("Cancel button not found!")

    # Access and interact with the "Enter Bills" window
    bill_entry_window = main_window.child_window(title_re=".*Enter Bills.*")

    if bill_entry_window.exists():
        dual_print("Enter Bills window found!")
        clear_button = bill_entry_window.child_window(title="Clear")

        if clear_button.exists():
            dual_print("Clear button found!")
            clear_button.click_input()
        else:
            dual_print("Clear button not found!")
    else:
        dual_print("Enter Bills window not found!")


if __name__ == "__main__":
    # Connect to the QuickBooks application
    app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32.EXE")  


    if app.is_process_running():
        dual_print("QuickBooks is running!")
    else:
        dual_print("QuickBooks is not running!")

    main_window = app.window(title_re='.*QuickBooks Desktop Pro 2019.*')

    if main_window.exists():
        dual_print("Main window exists")
        message_window_exists, message_window = find_message_window(main_window)
        if message_window_exists:
            handle_error(message_window, main_window)

    output_file.close()