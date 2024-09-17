from docx import Document
from docx.shared import Inches
from docx.shared import Pt

def add_table_with_images(doc, header_text, image_path1, image_path2):
   # Find the paragraph with the header text
    target_paragraph = None
    for paragraph in doc.paragraphs:
        if header_text in paragraph.text:
            target_paragraph = paragraph
            break

    if target_paragraph is None:
        print(f"Header '{header_text}' not found in the document.")
        return

    p1 = target_paragraph.insert_paragraph_before()

    # Add a table after the new paragraph
    # We can't directly control placement through doc.add_table, so we'll insert it programmatically
    table = doc.add_table(rows=1, cols=2)

    # Move the table to the specific location after the target paragraph
    # Using the XML elements for moving the table
    target_paragraph._element.addnext(table._element)

    # Add images to each cell in the table
    for i, image_path in enumerate([image_path1, image_path2]):
        cell = table.cell(0, i)
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(image_path, width=Inches(3))  # Adjust width as needed

    # Save the modified document
    doc.save('modified_document.docx')
    print("Document modified successfully.")

def replace_text_in_paragraph(paragraph, old_texts, new_texts):
    for old_text, new_text in zip(old_texts, new_texts):
        if old_text in paragraph.text:
            paragraph.text = paragraph.text.replace(old_text, new_text)

            run = paragraph.runs[0]
            run.font.name = 'Calibri (Body)'
            run.font.size = Pt(11)

def replace_text_in_table(table, old_texts, new_texts):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_text_in_paragraph(paragraph, old_texts, new_texts)