from pywinauto import Application

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

def check_exist(element):
    if element.exists():
        dual_print(f"{element} exists.")
    else:
        dual_print(f"{element} does not exist.")

# Connect to the QuickBooks application
app = Application(backend="win32").connect(path=r"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32.EXE")  


if app.is_process_running():
    dual_print("QuickBooks is running!")
else:
    dual_print("QuickBooks is not running!")

main_window = app.window(title_re='.*QuickBooks Desktop Pro 2019.*')
if main_window.exists():
    dual_print('mwe xists')

# print_element_info(main_window)


# for child in main_window.children():
#     print(child.window_text() + "child")
#     print(child.window_text(), "-", child.element_info.control_type, "-", child.class_name())

# Attempt to fetch the "QuickBooks Message" window
message_window = app.window(title_re=".*QuickBooks Message.*")

check_exist(message_window)

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