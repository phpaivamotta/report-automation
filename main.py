from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.shared import Inches
from utils import add_table_with_images
from utils import replace_text_in_table
from utils import add_captions_with_win32com
from utils import add_bullets_above_tables
from utils import append_cross_references_to_bullets
from utils import delete_template_bullets
from utils import get_images_from_folder
from dotenv import load_dotenv
import os
import time



# Load environment variables from .env file
load_dotenv(override=True)
template_file_path = os.getenv('TEMPLATE_DOC_PATH')
output_doc_file_path = os.getenv('OUTPUT_REPORT_DOC_PATH')
images_folder_path = os.getenv('IMAGES_FOLDER_PATH')

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

    # Check if file path exists
    if os.path.exists(template_file_path):

        # Open an existing document
        doc = Document(template_file_path)

        # Modify core properties
        core_properties = doc.core_properties
        # Set core properties
        core_properties.title = doc_core_properties["title"]
        core_properties.author = doc_core_properties["author"]
        core_properties.subject = doc_core_properties["subject"]
        core_properties.keywords = doc_core_properties["keywords"]

        # Iterate through all tables (if any)
        for table in doc.tables:
            replace_text_in_table(table, custom_properties.keys(), custom_properties.values())

        # Check if there are any images in the Images folder
        images = get_images_from_folder(images_folder_path)

        if images:
            table_counter = 0
            # Loop through images
            for i in range(0, len(images), 2):
                if i + 1 >= len(images):
                    image_path_1 = images[i]
                    num_cols = 1
                    break
                image_path_1 = images[i]
                image_path_2 = images[i+1]
                num_cols = 2
            
                # Add create tables and add images to them
                add_table_with_images(doc, "Inspection Observations:", table_counter, num_cols, image_path_1, image_path_2)

                # Save the modified document
                doc.save(output_doc_file_path)

                # Add captions to images
                add_captions_with_win32com(output_doc_file_path, i, num_cols, image_path_1, image_path_2)

                # Add bullets above table
                doc = add_bullets_above_tables(output_doc_file_path, table_counter, num_cols)   

                # Add cross references
                append_cross_references_to_bullets(output_doc_file_path, i)

                table_counter += 1

        # Delete the first 3 template bullets (necessary to add bullet styles)
        delete_template_bullets(output_doc_file_path)

    else:
        print(f"The template file path {template_file_path} does not exist. Please input a valid file path.")