from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import win32com.client as win32
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.ns import qn



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

    target_paragraph.insert_paragraph_before()

    # Add a table after the new paragraph
    # We can't directly control placement through doc.add_table, so we'll insert it programmatically
    table = doc.add_table(rows=1, cols=2)

    set_table_borders(table)

     # Disable automatic table resizing
    table.autofit = False

    # Set table alignment to center
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths
    for column in table.columns:
        column.width = Inches(3.7)  # Set column width to 2 inches
        for cell in column.cells:
            cell.width = Inches(3.7)

    # Move the table to the specific location after the target paragraph
    # Using the XML elements for moving the table
    target_paragraph._element.addnext(table._element)

    # Add images to each cell in the table
    for i, image_path in enumerate([image_path1, image_path2]):
        cell = table.cell(0, i)
        set_cell_margins(table, left=72, right=72, top=72, bottom=0)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_TABLE_ALIGNMENT.CENTER
        run = paragraph.add_run()
        run.add_picture(image_path, width=Inches(3.6))  # Adjust width as needed

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


def add_captions_with_win32com(doc_path):
    # Open Word application
    word = win32.Dispatch('Word.Application')
    word.Visible = False  # Set to True if you want to see Word while working

    # Open the existing document
    doc = word.Documents.Open(doc_path)

    # Loop through all inline shapes (images) in the document
    for inline_shape in doc.InlineShapes:
        # Select the inline shape (image)
        inline_shape.Select()

        # Insert a caption for the selected image
        word.Selection.InsertCaption(Label="Figure", Title=": This is the caption text.", TitleAutoText="", ExcludeLabel=False)

    # Update all fields (important for cross-references)
    doc.Fields.Update()

    # Save and close the document
    doc.Save()
    # doc.Close()
    # word.Quit()

    print("Captions added successfully.")


def set_table_borders(table):
    """
    Set the borders of the table to the specified color with a width of 2.25 pt.
    """
    # Calculate the border size in eighths of a point
    border_size = 18  # 2.25 pt * 8 = 18 units

    # Define the border color
    border_color = "002060"  # Hex color code for RGB(0,32,96)

    # Define the border style XML
    tbl_borders = parse_xml(r'''
        <w:tblBorders %s>
            <w:top w:val="single" w:sz="%d" w:space="0" w:color="%s"/>
            <w:left w:val="single" w:sz="%d" w:space="0" w:color="%s"/>
            <w:bottom w:val="single" w:sz="%d" w:space="0" w:color="%s"/>
            <w:right w:val="single" w:sz="%d" w:space="0" w:color="%s"/>
            <w:insideH w:val="single" w:sz="%d" w:space="0" w:color="%s"/>
            <w:insideV w:val="single" w:sz="%d" w:space="0" w:color="%s"/>
        </w:tblBorders>
        ''' % (
            nsdecls('w'),
            border_size, border_color,
            border_size, border_color,
            border_size, border_color,
            border_size, border_color,
            border_size, border_color,
            border_size, border_color
        ))

    # Access or create the table properties element
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)

    # Remove any existing borders element
    tblBorders = tblPr.find(qn('w:tblBorders'))
    if tblBorders is not None:
        tblPr.remove(tblBorders)

    # Append the borders element to the table properties
    tblPr.append(tbl_borders)


def set_cell_margins(table, left=0, right=0, top=0, bottom=0):
    tc = table._element
    tblPr = tc.tblPr
    tblCellMar = OxmlElement('w:tblCellMar')
    kwargs = {"left":left, "right":right, "top":top, "bottom":bottom}
    for m in ["left","right", "top", "bottom"]:
        node = OxmlElement("w:{}".format(m))
        node.set(qn('w:w'), str(kwargs.get(m)))
        node.set(qn('w:type'), 'dxa')
        tblCellMar.append(node)

    tblPr.append(tblCellMar)