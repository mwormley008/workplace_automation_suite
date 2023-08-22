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

# Connect to the QuickBooks application
app = Application(backend="win32").connect(path=r"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32.EXE")  


if app.is_process_running():
    dual_print("QuickBooks is running!")
else:
    dual_print("QuickBooks is not running!")

main_window = app.window(title_re='.*QuickBooks Desktop Pro 2019.*')

# print_element_info(main_window)


# for child in main_window.children():
#     print(child.window_text() + "child")
#     print(child.window_text(), "-", child.element_info.control_type, "-", child.class_name())

# Attempt to fetch the "QuickBooks Message" window
message_window = main_window.child_window(title="QuickBooks Message", control_type="Window")
cancel_button = message_window.child_window(title="Cancel", control_type="Pane")

if cancel_button.exists():
    dual_print("Cancel button found!")
    try:
        cancel_button.click_input()
    except Exception as e:
        dual_print("Error while clicking:", e)
else:
    dual_print("Cancel button not found!")


# Access the Workspace pane from the main window
workspace_pane = main_window.child_window(title="Workspace", control_type="Pane", class_name="MDIClient")

# From the Workspace pane, access the Enter Bills window
bill_entry_window = workspace_pane.child_window(title_re='.*Enter Bills.*', control_type="Window")

if bill_entry_window.exists():
    dual_print("Enter Bills window found!")
else: 
    dual_print("Enter Bills window not found!")

# Once we've accessed the Enter Bills window, find the Clear button
clear_button = bill_entry_window.child_window(title="Clear", control_type="Pane", class_name="MauiPushButton")

# Check if the button exists and click on it
if clear_button.exists():
    clear_button.click_input()
else:
    dual_print("Clear button not found!")


output_file.close()