import re, docx
import sys
import os
import docx
from docx import Document


def find_date_paragraph_index(doc):
    """
    This function finds and returns the index of the paragraph containing the date.
    """
    date_pattern = re.compile(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b')
    for i, para in enumerate(doc.paragraphs):
        if date_pattern.search(para.text):
            return i  # Return the index instead of the paragraph object
    return None

def mark_revision(doc, keyword="Revised"):
    """
    This function checks the paragraph following the date for the revision keyword 
    and updates the text based on whether it is already revised.
    """
    date_para_index = find_date_paragraph_index(doc)

    # Initialize to 1, assuming it's the first revision if no previous keyword is found
    revision_number = 1  

    # If there's no date paragraph, we'll still proceed with the assumption that it's a valid document for revision.
    if date_para_index is not None and date_para_index + 1 < len(doc.paragraphs):
        target_para = doc.paragraphs[date_para_index + 1]
        text = target_para.text.split()

        for i, word in enumerate(text):
            if word == keyword and i + 1 < len(text) and text[i + 1].isdigit():
                revision_number = int(text[i + 1]) + 1  # increment the revision number
                text[i + 1] = str(revision_number)
                break
        else:
            # If "Revised" keyword is present but without a number, add the number.
            if keyword in target_para.text:
                text.append(str(revision_number))
            else:
                # If the "Revised" keyword doesn't exist, add "Revised 1" at the end.
                text.append(f"{keyword} 1")

        # Update paragraph text with the tabs
        tabs = '\t' * 6
        new_text = ' '.join(text)
        target_para.clear()
        target_para.add_run(tabs + new_text)  # Adding the new text with tabs

    else:
        # If we don't find a paragraph following the date, we will create a new paragraph with the revision note.
        tabs = '\t' * 6
        new_para = doc.add_paragraph()
        new_para.add_run(tabs + f"{keyword} {revision_number}")  # Adding the new text with tabs

    return revision_number  # Return the revision number
    
def main(file_path):  # Now main accepts an argument, which is the file path of the document.
    try:
        # Load the existing document
        doc = Document(file_path)

        # Mark the document as revised
        revision_number = mark_revision(doc)
        
        if revision_number is not None:
                # If a revision was made, save the document with the new revision number in the name
                revised_file_path = os.path.splitext(file_path)[0] + f" Rev {revision_number}" + os.path.splitext(file_path)[1]
                doc.save(revised_file_path)
                print(f"Document revised and saved as '{revised_file_path}'")
        else:
            print("No revisions were made to the document.")

    except Exception as e:
        print(f"An error occurred: {e}")
        return None  # Return None or a suitable message indicating the failure.

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Error: No file path provided.")
        sys.exit(1)

    # sys.argv[1] will be the file path argument passed from the command line.
    file_path = sys.argv[1]
    main(file_path)