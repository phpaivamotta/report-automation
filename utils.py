from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import win32com.client as win32
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls
from docx.oxml.ns import qn
import time



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
    # doc.save('modified_document.docx')
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

        # Insert a caption for the selected image (exclude custom title, only label + number)
        word.Selection.InsertCaption(Label="Figure", Title=". Figure caption", ExcludeLabel=False, Position=win32.constants.wdCaptionPositionBelow)

        # Move the selection to the end of the caption
        word.Selection.MoveRight(Unit=win32.constants.wdWord, Count=1, Extend=win32.constants.wdExtend)

        # Delete any text after the caption label (if any text remains after "Figure X")
        word.Selection.TypeBackspace()

    # Update all fields (important for cross-references)
    doc.Fields.Update()

    # Save and close the document
    doc.Save()
    doc.Close()
    word.Quit()

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


def add_bullets_above_tables(output_doc_file_path):
    #Open doc
    doc = Document(output_doc_file_path)

    # Loop through all tables in the document
    tables = doc.tables

    for i, table in enumerate(tables):
        # Skip the first table
        if i == 0:
            continue

        # Find the paragraph just before the table
        paragraph_before_table = table._element.getprevious()

        if paragraph_before_table is not None:
            # Insert two bullet points above the table
            bullet_1 = doc.add_paragraph("Bullet point 1", style='List Bullet 2')
            bullet_2 = doc.add_paragraph("Bullet point 2", style='List Bullet 2')
            
            # Insert the bullet points before the table
            paragraph_before_table.addnext(bullet_2._element)
            bullet_2._element.addprevious(bullet_1._element)
            # Insert an empty paragraph (for space) after the bullets
            empty_space = doc.add_paragraph("")
            bullet_2._element.addnext(empty_space._element)

    doc.save(output_doc_file_path)
    print(f"Bullets added above all tables except the first")


def append_cross_references_to_bullets(docx_path):
    """Append cross-references to the beginning of each bullet point without deleting text."""
    # Open Word application
    word = win32.Dispatch('Word.Application')
    word.Visible = False  # Set to True if you want to see Word while working

    # Open the existing document
    doc = word.Documents.Open(docx_path)

    # Add a small delay to ensure the document is ready
    time.sleep(2)  # Delay for 2 seconds

    # Get the cross-reference items for "Figure"
    ref_items = doc.GetCrossReferenceItems("Figure")
    
    # Debug: Print available cross-reference items
    print("Available cross-reference items for 'Figure':")
    for idx, item in enumerate(ref_items):
        print(f"{idx+1}: {item}")
    
    if len(ref_items) < 2:
        print("Error: There are fewer than two figures to reference.")
        doc.Close(False)
        word.Quit()
        return

    # Set the figure references for bullet 1 and bullet 2
    figure_1_ref = 1  # Reference to Figure 1
    figure_2_ref = 2  # Reference to Figure 2

    # Loop through the paragraphs to find bullet points and append cross-references
    for para in doc.Paragraphs:
        if para.Range.Text.strip() == "Bullet point 1":
            # Move the cursor to the beginning of the paragraph and insert the cross-reference
            word.Selection.SetRange(para.Range.Start, para.Range.Start)
            word.Selection.InsertCrossReference(
                ReferenceType="Figure",
                ReferenceKind=3,  # 3 corresponds to wdOnlyLabelAndNumber (Figure X)
                ReferenceItem=figure_1_ref,
                InsertAsHyperlink=True,  # Optional: make it a hyperlink
                IncludePosition=False,
                SeparateNumbers=False,
                SeparatorString=" "
            )

            # Bold the inserted Figure 1 text
            word.Selection.MoveLeft(Unit=win32.constants.wdCharacter, Count=len(ref_items[figure_1_ref-1]) + 1, Extend=True)  # Move selection to include Figure text
            word.Selection.Font.Bold = True

            # Move cursor to the end of the bolded Figure 1 text
            word.Selection.Collapse(Direction=win32.constants.wdCollapseEnd)

            # Append the original text
            word.Selection.TypeText(" ")  # Add a space before the original text

            set_font_formatting(para, word)

        elif para.Range.Text.strip() == "Bullet point 2":
            # Move the cursor to the beginning of the paragraph and insert the cross-reference
            word.Selection.SetRange(para.Range.Start, para.Range.Start)
            word.Selection.InsertCrossReference(
                ReferenceType="Figure",
                ReferenceKind=3,  # 3 corresponds to wdOnlyLabelAndNumber (Figure X)
                ReferenceItem=figure_2_ref,
                InsertAsHyperlink=True,  # Optional: make it a hyperlink
                IncludePosition=False,
                SeparateNumbers=False,
                SeparatorString=" "
            )
            
            # Bold the inserted Figure 2 text
            word.Selection.MoveLeft(Unit=win32.constants.wdCharacter, Count=len(ref_items[figure_2_ref-1]) + 1, Extend=True)  # Move selection to include Figure text
            word.Selection.Font.Bold = True

            # Move cursor to the end of the bolded Figure 2 text
            word.Selection.Collapse(Direction=win32.constants.wdCollapseEnd)

            # Append the original text
            word.Selection.TypeText(" ")  # Add a space before the original text

            set_font_formatting(para, word)

    # Save the document with cross-references
    doc.SaveAs(docx_path)
    doc.Close()
    word.Quit()

    print(f"Cross-references appended to bullets and saved to {docx_path}")


def set_font_formatting(para, word):
    """Set font formatting for the paragraph to Calibri 12."""
    # Apply the font to the whole range of the paragraph
    para.Range.Font.Name = 'Calibri (Body)'
    para.Range.Font.Size = 12