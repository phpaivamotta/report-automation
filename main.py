from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.shared import Inches
from utils import add_table_with_images
from utils import replace_text_in_table
from utils import add_formatted_bullets
from utils import add_captions_with_win32com
from utils import add_bullets_above_tables
from utils import append_cross_references_to_bullets
from utils import delete_template_bullets
from utils import get_images_from_folder
from utils import remove_empty_paragraphs_after_table
from utils import remove_first_empty_paragraph_above_text
from utils import insert_formatted_text_after_header
from utils import read_report_data
from utils import read_picture_data
from wordextraction import get_next_report_id
from wordextraction import create_report_folder
from wordextraction import extract_report_data
from wordextraction import process_four_column_row
from wordextraction import process_picture_row
from wordextraction import extract_and_save_image
from wordextraction import append_to_csv
from dotenv import load_dotenv
import os
import time

# Load environment variables from .env file
load_dotenv(override=True)
template_file_path = os.getenv('TEMPLATE_DOC_PATH')
output_folder_path = os.getenv('OUTPUT_REPORT_FOLDER_PATH')
extracted_data_path = os.getenv('EXTRACTED_DATA_PATH')
input_doc_file_path = os.getenv('INPUT_DOC_PATH')

# Core document properties inputs
doc_core_properties = {
    "title": "Title", # customer
    "author": "author", # from
    "subject": "subject", #
    "keywords": "keywords", # maverick job num
}

# Custom document properties inputs
custom_properties = {
    "customer address": "customer address",
    "customer contact": "customer contact",
    "inspection site": "inspection site",
    "customer po num": "customer po num",
    "customer ccs": "customer css",
    "inspection date": "inspection date",
    "company": "company",
    "maverick contact info": "maverick contact info",
    "maverick ccs": "maverick ccs",
    "report date": "report date",
}

if __name__ == "__main__":

    report_csv_path = os.path.join(extracted_data_path, 'report_info.csv')
    picture_csv_path = os.path.join(extracted_data_path, 'picture_info.csv')

    # Get the latest report ID
    latest_report_id = get_next_report_id(report_csv_path) - 1

    # Ask user which report ID to use
    while True:
        try:
            report_id = int(input(f"Enter the report ID to generate (latest is {latest_report_id}): "))
            if 1 <= report_id <= latest_report_id:
                break
            else:
                print(f"Please enter a valid report ID between 1 and {latest_report_id}.")
        except ValueError:
            print("Please enter a valid integer.")

    # Read report data
    report_data = read_report_data(report_csv_path, report_id)
    if not report_data:
        print(f"No data found for report ID {report_id}")
    else:
        # Read picture data
        picture_data = read_picture_data(picture_csv_path, report_id)

        # Generate unique output file name
        output_file_name = f"Report_{report_id:04d}_{report_data['Customer'].replace(' ', '_')}.docx"
        output_doc_file_path = os.path.join(output_folder_path, output_file_name)

        # Check if file path exists
        if os.path.exists(template_file_path):
            # Open an existing document and save new
            template_doc = Document(template_file_path)
            template_doc.save(output_doc_file_path) # Save template to new location
            working_doc = Document(output_doc_file_path) # Open new file in the new location

            # Modify core properties
            core_properties = working_doc.core_properties
            # Set core properties
            core_properties.title = report_data['Customer']
            core_properties.author = report_data['From']
            core_properties.subject = report_data['Subject']
            core_properties.keywords = report_data['Maverick Job']
            core_properties.comments = report_data['Customer Contact']

            # Update custom properties
            custom_properties = {
                "customer address": report_data['Customer Address'],
                "inspection site": report_data['Inspection Site'],
                "customer po num": report_data['Customer PO No.'],
                "customer ccs": report_data['Customer CCs'],
                "inspection date": report_data['Inspection Date(s)'],
                "company": report_data['Company'],
                "maverick contact info": report_data['Maverick Contact Info'],
                "maverick ccs": report_data['Maverick CCs'],
                "report date": report_data['Report Date'],
            }

            working_doc.save(output_doc_file_path)

            # Iterate through all tables (if any)
            for table in working_doc.tables:
                replace_text_in_table(table, custom_properties.keys(), custom_properties.values())

            working_doc.save(output_doc_file_path)

            #Insert Introduction
            insert_formatted_text_after_header(output_doc_file_path, "Introduction", report_data['Introduction'])

            #Insert Conclusion
            insert_formatted_text_after_header(output_doc_file_path, "Inspection Conclusions and Recommendations", report_data['Conclusions'])

            #Add Drawings
            add_formatted_bullets(output_doc_file_path, "The following drawings were provided and used during the inspection:", ["New Drawing 1"])

            #Reopen document after using win32com
            working_doc = Document(output_doc_file_path)
    
            # Before running, make sure to save working doc
            working_doc.save(output_doc_file_path)

            # if picture_data:
            #     # Construct the path to the report's image folder
            #     report_folder = os.path.join(extracted_data_path, f"report_{report_id:04d}")
            #     table_counter = 0
            #     for i in range(0, len(picture_data), 2):
            #         image_1 = picture_data[i]
            #         image_path_1 = os.path.join(report_folder, image_1['Picture File Name'])
            #         description_1 = image_1['Description']
            #         caption_1 = image_1['Caption']
                    
            #         if i + 1 < len(picture_data):
            #             image_2 = picture_data[i+1]
            #             image_path_2 = os.path.join(report_folder, image_2['Picture File Name'])
            #             description_2 = image_2['Description']
            #             caption_2 = image_2['Caption']
            #             num_cols = 2
            #         else:
            #             image_path_2 = None
            #             description_2 = None
            #             caption_2 = None
            #             num_cols = 1
                    
            #         # Add create tables and add images to them
            #         add_table_with_images(output_doc_file_path, "Inspection Observations:", table_counter, num_cols, image_path_1, image_path_2)
                    
            #         # Add captions to images
            #         add_captions_with_win32com(output_doc_file_path, i, num_cols, image_path_1, image_path_2, caption_1, caption_2)

            #         # Add bullets above table with descriptions
            #         add_bullets_above_tables(output_doc_file_path, table_counter, num_cols)

            #         # Add cross references
            #         append_cross_references_to_bullets(output_doc_file_path, i, num_cols, description_1, description_2)

            #         table_counter += 1
        
            # # Delete the first 3 template bullets (necessary to add bullet styles) dont need this anymore????? It breaks if I leave it in?????
            # #delete_template_bullets(output_doc_file_path)
            
            # remove_empty_paragraphs_after_table(output_doc_file_path)
            
            # remove_first_empty_paragraph_above_text(output_doc_file_path, "Inspection Observations:")
            
            print(f"Report generated successfully: {output_file_name}")

        else:
            print(f"The template file path {template_file_path} does not exist. Please input a valid file path.")