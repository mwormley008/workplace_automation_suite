import re, docx
import sys
import os
import docx
from docx import Document
from docx.shared import Inches

def insert_text_in_table(doc, paragraph_index, text):
    # Add a table with one row and one column after the target paragraph
    table = doc.add_table(rows=1, cols=1)
    
    # Move the table below the target paragraph
    for _ in range(len(doc.paragraphs) - 2, paragraph_index, -1):
        move_table_after(doc.paragraphs[_], table)
    
    # No border for aesthetics
    set_table_border(table, False)
    
    # Inserting text
    cell = table.cell(0, 0)
    cell.text = text

def set_table_border(table, has_border):
    # A function to set or remove borders from a table
    for row in table.rows:
        for cell in row.cells:
            for key, value in cell._element.attrib.items():
                if 'border' in key:
                    cell._element.attrib[key] = '1' if has_border else '0'

def move_table_after(paragraph, table):
    """
    Move the given table after the given paragraph.
    """
    tbl, p = table._tbl, paragraph._p
    p.addnext(tbl)

def find_date_paragraph_index(doc):
    """
    This function finds and returns the index of the paragraph containing the date.
    """
    date_pattern = re.compile(r'\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\b')
    for i, para in enumerate(doc.paragraphs):
        if date_pattern.search(para.text):
            return i  # Return the index instead of the paragraph object
    return None


def mark_revision(doc, tab_count, keyword="Revised"):
    """
    This function marks the revision in the paragraph following the date paragraph,
    updating the text based on whether it is already revised.
    """
    # Initialize to 1, assuming it's the first revision if no previous keyword is found
    revision_number = 1  

    # Find the index of the date paragraph
    date_para_index = None
    for i, para in enumerate(doc.paragraphs):
        if "January" in para.text or "February" in para.text or \
           "March" in para.text or "April" in para.text or \
           "May" in para.text or "June" in para.text or \
           "July" in para.text or "August" in para.text or \
           "September" in para.text or "October" in para.text or \
           "November" in para.text or "December" in para.text:
            date_para_index = i
            break

    target_para = doc.paragraphs[date_para_index + 1]
    text = target_para.text.split()

    # Check if revision number needs to be updated
    for i, word in enumerate(text):
        if word == keyword and i + 1 < len(text) and text[i + 1].isdigit():
            revision_number = int(text[i + 1]) + 1  # increment the revision number
            text[i + 1] = str(revision_number)
            break
    else:
        target_para.paragraph_format.space_before = Inches(.1)
        target_para.add_run(f"{keyword} {revision_number}")

   

    return revision_number

def set_tabs_in_paragraph(paragraph, tab_positions):
    # Clear existing tab stops
    paragraph.paragraph_format.tab_stops.clear_all()
    
    # Add new tab stops at specified positions
    for pos in tab_positions:
        paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(pos))

def main(file_path):
    try:
        # Load the existing document
        doc = Document(file_path)

        # Number of tab spaces you want before the revision note
        tab_count = 6

        # Mark the document as revised, specifying the tab count.
        revision_number = mark_revision(doc, tab_count)

        # Save the revised document with a new file name indicating the revision number
        revised_file_path = os.path.splitext(file_path)[0] + f" Rev {revision_number}" + os.path.splitext(file_path)[1]
        doc.save(revised_file_path)
        print(f"Document revised and saved as '{revised_file_path}'")
        return revised_file_path

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Error: No file path provided.")
        sys.exit(1)

    # sys.argv[1] will be the file path argument passed from the command line.
    file_path = sys.argv[1]
    revised_file_path = main(file_path)
    print(revised_file_path)