import win32com.client

# Path to your Word document
file_path = r"C:\Users\Michael\Desktop\CEnvelopes\The Tax Shoppe.docx"

# Start an instance of Word
word = win32com.client.Dispatch('Word.Application')

# Open the document
doc = word.Documents.Open(file_path)

# Print the document to the default printer
doc.PrintOut()

# Close the document
doc.Close()

# Quit Word
word.Quit()

