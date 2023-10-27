import re, docx
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

    if date_para_index is None:
        print("No paragraph containing a valid date was found.")
        return

    # Check if a paragraph exists after the date paragraph
    if date_para_index + 1 >= len(doc.paragraphs):
        print("No paragraph exists after the date paragraph.")
        return

    target_para = doc.paragraphs[date_para_index + 1]

    if keyword in target_para.text:
        # The keyword exists, so we want to increment the revision number.
        revision_number = 1
        text = target_para.text.split()

        for i, word in enumerate(text):
            if word == keyword and i + 1 < len(text) and text[i + 1].isdigit():
                revision_number = int(text[i + 1]) + 1
                text[i + 1] = str(revision_number)
                break
        else:
            # If there was no number, we add one at the end.
            text.append(str(revision_number))

        # Update paragraph text
        new_text = ' '.join(text)
        target_para.clear()  # Clear the existing paragraph content
        # Add six tab characters before the revised text
        tabs = '\t' * 6
        run = target_para.add_run(tabs + new_text)

    else:
        # If the keyword doesn't exist in the paragraph, we add it.
        if target_para.text.strip():  # ensuring paragraph is not empty
            # Add six tab characters before the keyword
            tabs = '\t' * 6
            run = target_para.add_run(tabs + f"({keyword})")

def main():
    # Load the existing document
    doc = Document('Cobalt Self-Storage Bensenville.docx')  # replace 'your_document.docx' with your actual file

    mark_revision(doc)

    # Save the revised document
    doc.save('revised_document.docx')

if __name__ == "__main__":
    main()
