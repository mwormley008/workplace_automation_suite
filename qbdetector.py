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

# Connect to the QuickBooks application
app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32.EXE")  


if app.is_process_running():
    dual_print("QuickBooks is running!")
else:
    dual_print("QuickBooks is not running!")

main_window = app.window(title_re='.*QuickBooks Desktop Pro 2019.*')
if main_window.exists():
    dual_print("Main window exists")
    create_invoices_window_title = "Create Invoices (Editing Transaction...) "
    create_invoices_window = main_window.child_window(title=create_invoices_window_title)
    # Navigate the control hierarchy to get to the "Duplicate Invoice" button
    duplicate_invoice_btn = main_window.child_window(auto_id="DuplicateBtn")

    # Check if the button is found and click it
    if duplicate_invoice_btn.exists():
        print("Found 'Duplicate Invoice' button!")
        duplicate_invoice_btn.click_input()
    else:
        print("'Duplicate Invoice' button not found!")
    qb_wpf_container = app.window(auto_id="QB2WPFContainer")
    print(qb_wpf_container.children())
    

    buttons = [child for child in create_invoices_window.children()]# if isinstance(child, ButtonWrapper)]

    for button in buttons:
        print(button)
        button.click_input()
        print(button)

        time.sleep(4)  # Pause for 3 seconds to observe

    check_exist(duplicate_invoice_btn, 'create copy button')

    # duplicate_invoice_btn = create_invoices_window.child_window(automation_id="DuplicateBtn", class_name="RibbonButton")
    # check_exist(duplicate_invoice_btn, 'create copy button')
    # unnamed_pane = create_invoices_window.child_window(control_type="Pane", title="")
    # unnamed_custom = unnamed_pane.child_window(control_type="Custom", title="")
    # invoice_ribbon = unnamed_custom.child_window(auto_id="Invoice Ribbon", control_type="Tab")
    # main_tab = invoice_ribbon.child_window(auto_id="Main", control_type="TabItem")
    # ribbon_group = main_tab.child_window(control_type="Group", title_re="Microsoft.Windows.Controls.Ribbon.RibbonGroup Header: Items.Count:6")
    # create_copy_btn = ribbon_group.child_window(control_type="ListItem", title="Create a Copy")
    # check_exist(unnamed_pane, 'unnamed pane') 
    # print_element_info(create_invoices_window)
    if duplicate_invoice_btn.exists():
        dual_print("Found 'Create a Copy' button!")
        duplicate_invoice_btn.click_input()
    else:
        dual_print("'Create a Copy' button not found!")


# print_element_info(main_window)


# Attempt to fetch the "QuickBooks Message" window
message_window = app.window(title_re=".*QuickBooks Message.*")

check_exist(message_window, 'message window')

cancel_button = message_window.child_window(title="Cancel")

if cancel_button.exists():
    dual_print("Cancel button found!")
    cancel_button.click_input()
else:
    dual_print("Cancel button not found!")

# Access and interact with the "Enter Bills" window
bill_entry_window = main_window.child_window(title_re=".*Enter Bills.*")


if bill_entry_window.exists():
    # print_element_info(bill_entry_window)
    dual_print("Enter Bills window found!")

    # Find and click the "Clear" button
    clear_button = bill_entry_window.child_window(title="Clear")

    if clear_button.exists():
        dual_print("Clear button found!")
        clear_button.click_input()
    else:
        dual_print("Clear button not found!")
else:
    dual_print("Enter Bills window not found!")

output_file.close()