import os
import csv
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Parent must be a Document or _Cell instance")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_and_save(docx_path, output_folder):
    document = Document(docx_path)
    
    os.makedirs(output_folder, exist_ok=True)
    
    # Define CSV file paths
    table1_csv_path = os.path.join(output_folder, 'table1_data.csv')
    data_tables_csv_path = os.path.join(output_folder, 'data_tables.csv')
    picture_data_csv_path = os.path.join(output_folder, 'picture_data.csv')
    
    # Open CSV files
    with open(table1_csv_path, 'w', newline='', encoding='utf-8') as table1_csvfile, \
         open(data_tables_csv_path, 'w', newline='', encoding='utf-8') as data_tables_csvfile, \
         open(picture_data_csv_path, 'w', newline='', encoding='utf-8') as picture_csvfile:
        
        table1_writer = csv.writer(table1_csvfile)
        data_tables_writer = csv.writer(data_tables_csvfile)
        picture_csv_writer = csv.writer(picture_csvfile)
        
        # Write headers for each CSV
        table1_writer.writerow(['Field1', 'Field2'])
        data_tables_writer.writerow(['Content'])
        picture_csv_writer.writerow(['Index', 'Description', 'Caption', 'Picture File Name'])
        
        image_count = 1
        table_number = 0
        
        for block in iter_block_items(document):
            if isinstance(block, Table):
                # Determine the number of columns in the table
                first_row = block.rows[0]
                num_columns = len(first_row.cells)
                
                if table_number == 0:
                    # Process the first table (4 columns)
                    for row in block.rows:
                        cells = row.cells
                        if len(cells) == 4:
                            # Write columns 1 and 2 into Field1 and Field2
                            field1 = cells[0].text.strip()
                            field2 = cells[1].text.strip()
                            table1_writer.writerow([field1, field2])
                            # Write columns 3 and 4 into Field1 and Field2
                            field1 = cells[2].text.strip()
                            field2 = cells[3].text.strip()
                            table1_writer.writerow([field1, field2])
                elif num_columns == 1:
                    # Process one-column tables (like Table 2)
                    for row in block.rows:
                        cell_text = row.cells[0].text.strip()
                        data_tables_writer.writerow([cell_text])
                elif num_columns == 4:
                    # Process picture tables
                    for row in block.rows[1:]:  # Skip header row
                        cells = row.cells
                        if len(cells) == 4:
                            index = cells[0].text.strip()
                            description = cells[1].text.strip()
                            caption = cells[2].text.strip()
                            
                            # Handle picture
                            picture_cell = cells[3]
                            picture_filename = ""
                            for paragraph in picture_cell.paragraphs:
                                for run in paragraph.runs:
                                    # Check if the run contains a drawing (image)
                                    if 'Drawing' in run._element.xml:
                                        # Extract the alt-text (description) of the image
                                        docPr_elements = run._element.xpath('.//wp:docPr')
                                        if docPr_elements:
                                            alt_text = docPr_elements[0].get('descr')
                                            if alt_text:
                                                # Extract the file name from the file path in alt-text
                                                file_name = os.path.basename(alt_text)
                                                picture_filename = file_name
                                                # Get the image part
                                                blip = run._element.xpath('.//a:blip')[0]
                                                embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                                image_part = run.part.related_parts[embed]
                                                
                                                # Save image using file name
                                                image_extension = os.path.splitext(image_part.partname)[1]
                                                # Ensure the file name has the correct extension
                                                if not file_name.lower().endswith(image_extension.lower()):
                                                    file_name = os.path.splitext(file_name)[0] + image_extension
                                                # Add a unique prefix to prevent overwriting files
                                                image_name = f"{image_count:03d}_{file_name}"
                                                image_path = os.path.join(output_folder, image_name)
                                                with open(image_path, 'wb') as f:
                                                    f.write(image_part.blob)
                                                
                                                image_count += 1
                                                break  # Assume one image per cell
                                        else:
                                            print(f"No alt-text found for image at index {index}")
                                if picture_filename:
                                    break
                            
                            # Write to picture CSV
                            picture_csv_writer.writerow([index, description, caption, picture_filename])
                else:
                    # If the table doesn't match any known format, you can choose to log it or handle it differently
                    print(f"Unknown table format with {num_columns} columns at table number {table_number + 1}")
                table_number += 1
        
    print(f"Data from the first table extracted and saved to {table1_csv_path}")
    print(f"Additional data extracted and saved to {data_tables_csv_path}")
    print(f"Picture data extracted and saved to {picture_data_csv_path}")
    print(f"Images saved in {output_folder}")

# Usage NEEDS DOCM TO BE SAVED AS DOCX
docx_path = r'C:\Users\dfernandez\OneDrive - Maverick Applied Science\Desktop\ReportInputsTemplate.docx'
output_folder = 'image_temp'
extract_and_save(docx_path, output_folder)
