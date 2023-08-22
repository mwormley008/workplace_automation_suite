from pywinauto import Application


def print_element_info(element, indent=0):
    """Recursive function to print UI element info."""
    
    # Print current element info
    prefix = '  ' * indent  # Increase indentation based on depth
    print(f"{prefix}{element.window_text()} - {element.element_info.control_type} - {element.class_name()}")
    
    # Recursive call for child elements
    for child in element.children():
        print_element_info(child, indent + 1)

# Connect to the QuickBooks application
app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\Intuit\QuickBooks 2019\QBW32.EXE")  


if app.is_process_running():
    print("QuickBooks is running!")
else:
    print("QuickBooks is not running!")

# Access the main window
for w in app.windows():
    print(w.window_text() + "we did it")

main_window = app.window_(title='WBR Roofing Company, Inc.  - QuickBooks Desktop Pro 2019(multi-user)(Michael)')

print_element_info(main_window)


for child in main_window.children():
    print(child.window_text() + "child")
    print(child.window_text(), "-", child.element_info.control_type, "-", child.class_name())

message_window = main_window.child_window(window_text="QuickBooks Message", control_type="Window")

if message_window.exists():
    cancel_button = message_window.child_window(window_text="Cancel", control_type="Pane")
    
    if cancel_button.exists():
        cancel_button.click()

