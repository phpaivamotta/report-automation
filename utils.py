from docx import Document
from docx.shared import Inches, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from copy import deepcopy
import csv
import os
import glob

_XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'


# ---------------------------------------------------------------------------
# Internal formatting helpers
# ---------------------------------------------------------------------------

def _set_run_font(run):
    run.font.name = 'Calibri'
    run.font.size = Pt(12)


def _set_para_spacing(para):
    para.paragraph_format.space_before = Pt(6)
    para.paragraph_format.space_after = Pt(6)


def _apply_formatting(para):
    _set_para_spacing(para)
    for run in para.runs:
        _set_run_font(run)


# ---------------------------------------------------------------------------
# Document properties
# ---------------------------------------------------------------------------

def update_document_properties(doc, report_data):
    doc.core_properties.title = report_data['Customer']
    doc.core_properties.author = report_data['From']
    doc.core_properties.subject = report_data['Subject']
    doc.core_properties.keywords = report_data['Maverick Job']

    CP_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
    VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'

    custom_values = {
        'customer address':  report_data.get('Customer Address', ''),
        'inspection site':   report_data.get('Inspection Site', ''),
        'customer po num':   report_data.get('Customer PO No.', ''),
        'customer ccs':      report_data.get('Customer CCs', ''),
        'inspection date':   report_data.get('Inspection Date(s)', ''),
        'maverick ccs':      report_data.get('Maverick CCs', ''),
        'report date':       report_data.get('Report Date', ''),
        'customer contacts': report_data.get('Customer Contact', ''),
    }

    custom_part = None
    for part in doc.part.package.iter_parts():
        if hasattr(part, 'partname') and str(part.partname) == '/docProps/custom.xml':
            custom_part = part
            break

    if custom_part is not None:
        from lxml import etree
        # Custom properties are loaded as a raw Part (bytes), not an XmlPart
        tree = etree.fromstring(custom_part.blob)
        for prop in tree.findall(f'{{{CP_NS}}}property'):
            name = prop.get('name')
            if name in custom_values:
                lpwstr = prop.find(f'{{{VT_NS}}}lpwstr')
                if lpwstr is not None:
                    lpwstr.text = str(custom_values[name])
        custom_part._blob = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
    else:
        print("Warning: custom properties part not found — custom fields not updated.")


# ---------------------------------------------------------------------------
# Text insertion
# ---------------------------------------------------------------------------

def insert_formatted_text_after_header(doc, header_text, content_to_insert):
    target_paragraph = None
    for paragraph in doc.paragraphs:
        if header_text in paragraph.text:
            target_paragraph = paragraph
            break

    if not target_paragraph:
        print(f"Header '{header_text}' not found.")
        return

    new_paragraph = doc.add_paragraph()
    target_paragraph._p.addnext(new_paragraph._p)
    run = new_paragraph.add_run(content_to_insert)
    _set_run_font(run)
    _set_para_spacing(new_paragraph)
    print(f"Formatted text inserted after '{header_text}'.")


def add_formatted_bullets(doc, header_text, new_content_list, is_drawing=True):
    target_paragraph = None
    target_index = None
    for i, paragraph in enumerate(doc.paragraphs):
        if header_text in paragraph.text:
            target_paragraph = paragraph
            target_index = i
            break

    if not target_paragraph or target_index + 1 >= len(doc.paragraphs):
        print(f"Header text '{header_text}' not found or it's the last paragraph.")
        return

    template_bullet = doc.paragraphs[target_index + 1]
    new_bullets = []
    for new_content in reversed(new_content_list):
        new_bullet = deepcopy(template_bullet._element)
        new_para = type(template_bullet)(new_bullet, template_bullet._parent)
        if is_drawing and not new_content.startswith("Equipment Drawing:"):
            new_content = f"Equipment Drawing: {new_content}"
        new_para.text = new_content
        for run in new_para.runs:
            _set_run_font(run)
        _set_para_spacing(new_para)
        new_bullets.append(new_para)

    for new_bullet in new_bullets:
        template_bullet._element.addnext(new_bullet._element)
    template_bullet._element.getparent().remove(template_bullet._element)


# ---------------------------------------------------------------------------
# Table creation
# ---------------------------------------------------------------------------

def add_table_with_images(doc, header_text, table_counter, num_cols, image_path1, image_path2=None):
    if table_counter == 0:
        target_paragraph = None
        for paragraph in doc.paragraphs:
            if header_text in paragraph.text:
                target_paragraph = paragraph
                break
    else:
        last_table = doc.tables[-1]
        tbl_element = last_table._element
        new_paragraph_element = OxmlElement('w:p')
        tbl_element.addnext(new_paragraph_element)
        target_paragraph = doc.add_paragraph()
        target_paragraph._element = new_paragraph_element

    if target_paragraph is None:
        print(f"Header '{header_text}' not found.")
        return None

    target_paragraph.insert_paragraph_before()

    table = doc.add_table(rows=1, cols=num_cols)
    set_table_borders(table)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for column in table.columns:
        column.width = Inches(3.7)
        for cell in column.cells:
            cell.width = Inches(3.7)

    target_paragraph._element.addnext(table._element)

    if num_cols == 1:
        cell = table.cell(0, 0)
        set_cell_margins(table, left=72, right=72, top=72, bottom=0)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_TABLE_ALIGNMENT.CENTER
        paragraph.add_run().add_picture(image_path1, width=Inches(3.6))

    elif num_cols == 2:
        cell = table.cell(0, 0)
        set_cell_margins(table, left=72, right=72, top=72, bottom=0)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_TABLE_ALIGNMENT.CENTER
        paragraph.add_run().add_picture(image_path1, width=Inches(3.6))

        cell = table.cell(0, 1)
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_TABLE_ALIGNMENT.CENTER
        paragraph.add_run().add_picture(image_path2, width=Inches(3.6))

    print("Table with images added successfully.")
    return table


# ---------------------------------------------------------------------------
# Bullets above tables
# ---------------------------------------------------------------------------

def add_bullets_above_tables(doc, table, num_cols):
    paragraph_before_table = table._element.getprevious()
    bullets = []

    if paragraph_before_table is not None:
        if num_cols >= 2:
            bullet_1 = doc.add_paragraph("Bullet point 1", style='List Bullet 2')
            bullet_2 = doc.add_paragraph("Bullet point 2", style='List Bullet 2')
            bullet_1.paragraph_format.keep_with_next = True
            bullet_2.paragraph_format.keep_with_next = True
            paragraph_before_table.addnext(bullet_2._element)
            bullet_2._element.addprevious(bullet_1._element)
            bullets = [bullet_1, bullet_2]
        else:
            bullet_1 = doc.add_paragraph("Bullet point 1", style='List Bullet 2')
            bullet_1.paragraph_format.keep_with_next = True
            paragraph_before_table.addnext(bullet_1._element)
            bullets = [bullet_1]

    print("Bullets added above table.")
    return bullets


# ---------------------------------------------------------------------------
# Caption paragraph — SEQ field with bookmark
# ---------------------------------------------------------------------------

def add_caption_paragraph(after_element, caption_text, figure_index):
    """
    Insert a caption paragraph immediately after after_element.

    Pass the image paragraph's _p element (inside the table cell) so the caption
    lands as para[1] inside that cell — matching Word's native InsertCaption behaviour.
    Returns (bookmark_name, caption_xml_element).
    """
    bm_name = f'_Ref_Fig_{figure_index}'
    bm_id = figure_index

    p = OxmlElement('w:p')

    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'Caption')
    pPr.append(pStyle)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)
    p.append(pPr)

    bm_start = OxmlElement('w:bookmarkStart')
    bm_start.set(qn('w:id'), str(bm_id))
    bm_start.set(qn('w:name'), bm_name)
    p.append(bm_start)

    r_fig = OxmlElement('w:r')
    t_fig = OxmlElement('w:t')
    t_fig.set(_XML_SPACE, 'preserve')
    t_fig.text = 'Figure '
    r_fig.append(t_fig)
    p.append(r_fig)

    # fldSimple matches the structure Word's InsertCaption produces
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), ' SEQ Figure \\* ARABIC ')
    r_num = OxmlElement('w:r')
    rPr_num = OxmlElement('w:rPr')
    rPr_num.append(OxmlElement('w:noProof'))
    r_num.append(rPr_num)
    t_num = OxmlElement('w:t')
    t_num.text = str(figure_index)
    r_num.append(t_num)
    fld.append(r_num)
    p.append(fld)

    bm_end = OxmlElement('w:bookmarkEnd')
    bm_end.set(qn('w:id'), str(bm_id))
    p.append(bm_end)

    r_cap = OxmlElement('w:r')
    t_cap = OxmlElement('w:t')
    t_cap.set(_XML_SPACE, 'preserve')
    t_cap.text = f': {caption_text}'
    r_cap.append(t_cap)
    p.append(r_cap)

    after_element.addnext(p)
    return bm_name, p


# ---------------------------------------------------------------------------
# Cross-reference bullet — REF field pointing to caption bookmark
# ---------------------------------------------------------------------------

def build_cross_reference_bullet(para, description, bookmark_name, figure_index):
    """
    Rewrite a placeholder bullet paragraph in-place as:
      [bold REF field → "Figure N"] [" shows "] [description]
    """
    p = para._element

    # Remove existing runs and hyperlinks
    for child in list(p):
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if local in ('r', 'hyperlink', 'ins', 'del'):
            p.remove(child)

    if description.startswith("Figure shows "):
        description = description[13:]

    def _bold_calibri_rpr():
        # \* Charformat copies the formatting of the instrText run's first character
        # onto the entire field result. All three properties must be set here so
        # that after a field update the result stays bold Calibri 12pt.
        rPr = OxmlElement('w:rPr')
        rPr.append(OxmlElement('w:b'))
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Calibri')
        rFonts.set(qn('w:hAnsi'), 'Calibri')
        rPr.append(rFonts)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '24')  # 12pt in half-points
        rPr.append(sz)
        return rPr

    def _text_rpr():
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Calibri')
        rFonts.set(qn('w:hAnsi'), 'Calibri')
        rPr.append(rFonts)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '24')
        rPr.append(sz)
        return rPr

    r_begin = OxmlElement('w:r')
    r_begin.append(_bold_calibri_rpr())
    fc_begin = OxmlElement('w:fldChar')
    fc_begin.set(qn('w:fldCharType'), 'begin')
    r_begin.append(fc_begin)
    p.append(r_begin)

    # instrText run carries the formatting \* Charformat will copy to the result
    r_instr = OxmlElement('w:r')
    r_instr.append(_bold_calibri_rpr())
    instr = OxmlElement('w:instrText')
    instr.set(_XML_SPACE, 'preserve')
    instr.text = f' REF {bookmark_name} \\h \\* Charformat '
    r_instr.append(instr)
    p.append(r_instr)

    r_sep = OxmlElement('w:r')
    fc_sep = OxmlElement('w:fldChar')
    fc_sep.set(qn('w:fldCharType'), 'separate')
    r_sep.append(fc_sep)
    p.append(r_sep)

    # Cached display — bold Calibri "Figure N" (Word replaces this on field update,
    # using the formatting it copied from the instrText run via \* Charformat)
    r_val = OxmlElement('w:r')
    r_val.append(_bold_calibri_rpr())
    t_val = OxmlElement('w:t')
    t_val.text = f'Figure {figure_index}'
    r_val.append(t_val)
    p.append(r_val)

    r_end = OxmlElement('w:r')
    fc_end = OxmlElement('w:fldChar')
    fc_end.set(qn('w:fldCharType'), 'end')
    r_end.append(fc_end)
    p.append(r_end)

    # " shows <description>" in Calibri 12pt
    r_desc = OxmlElement('w:r')
    r_desc.append(_text_rpr())
    t_desc = OxmlElement('w:t')
    t_desc.set(_XML_SPACE, 'preserve')
    t_desc.text = f' shows {description}'
    r_desc.append(t_desc)
    p.append(r_desc)

    _set_para_spacing(para)


# ---------------------------------------------------------------------------
# Cleanup
# ---------------------------------------------------------------------------

def remove_empty_paragraphs_after_table(doc):
    for i, table in enumerate(doc.tables):
        if i == 0:
            continue
        next_element = table._element.getnext()
        while next_element is not None and next_element.tag.endswith('p'):
            paragraph_text = "".join(next_element.itertext()).strip()
            if not paragraph_text:
                parent = next_element.getparent()
                parent.remove(next_element)
                next_element = table._element.getnext()
            else:
                break


def remove_first_empty_paragraph_above_text(doc, text):
    for paragraph in doc.paragraphs:
        if text in paragraph.text:
            prev_paragraph = paragraph._element.getprevious()
            if prev_paragraph is not None and prev_paragraph.tag.endswith('p'):
                prev_text = "".join(prev_paragraph.itertext()).strip()
                if not prev_text:
                    prev_paragraph.getparent().remove(prev_paragraph)
            break


# ---------------------------------------------------------------------------
# Table styling helpers
# ---------------------------------------------------------------------------

def set_table_borders(table):
    border_size = 18  # 2.25pt × 8
    border_color = "002060"
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
            border_size, border_color,
        ))
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(tbl_borders)


def set_cell_margins(table, left=0, right=0, top=0, bottom=0):
    tc = table._element
    tblPr = tc.tblPr
    tblCellMar = OxmlElement('w:tblCellMar')
    for m, v in [('left', left), ('right', right), ('top', top), ('bottom', bottom)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(v))
        node.set(qn('w:type'), 'dxa')
        tblCellMar.append(node)
    tblPr.append(tblCellMar)


# ---------------------------------------------------------------------------
# Misc utilities (unchanged)
# ---------------------------------------------------------------------------

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
    print("Project details in Table 1 were modified successfully.")


def delete_template_bullets(doc):
    count = 0
    for para in doc.paragraphs:
        if para.style.name in ["List Bullet", "List Bullet 2", "List Bullet 3"]:
            if count >= 3:
                break
            p = para._element
            p.getparent().remove(p)
            count += 1


def get_images_from_folder(folder_path):
    image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.gif', '*.bmp', '*.tiff']
    image_paths = []
    for extension in image_extensions:
        image_paths.extend(glob.glob(os.path.join(folder_path, extension)))
    return sorted(image_paths, key=lambda x: x.lower())


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._element = None


def add_page_break_below_table(doc):
    for i, table in enumerate(doc.tables):
        if i == 0:
            continue
        if i % 2 == 0:
            tbl_element = table._element
            new_paragraph = doc.add_paragraph()
            tbl_element.addnext(new_paragraph._element)
            new_paragraph.add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------------------
# CSV readers
# ---------------------------------------------------------------------------

def read_report_data(report_csv_path, report_id):
    with open(report_csv_path, 'r', newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if int(row['Report ID']) == report_id:
                return row
    return None


def read_picture_data(picture_csv_path, report_id):
    pictures = []
    with open(picture_csv_path, 'r', newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if int(row['Report ID']) == report_id:
                pictures.append(row)
    return pictures
