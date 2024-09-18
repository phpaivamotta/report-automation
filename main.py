from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.shared import Inches
from utils import add_table_with_images
from utils import replace_text_in_table
from utils import add_captions_with_win32com
from dotenv import load_dotenv
import os



# Load environment variables from .env file
load_dotenv()
image_path_1 = os.getenv('IMAGE_PATH_1')
image_path_2 = os.getenv('IMAGE_PATH_2')
template_file_path = os.getenv('TEMPLATE_DOC_PATH')
output_doc_file_path = os.getenv('OUTPUT_REPORT_DOC_PATH')

# Inputs
doc_core_properties = {
    "title": "Title", # customer
    "author": "author", # from
    "subject": "subject", #
    "keywords": "keywords", # maverick job num
}

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

    add_table_with_images(doc, "Inspection Observations:", image_path_1, image_path_2)

    # Save the modified document
    doc.save('modified_document.docx')

    add_captions_with_win32com(output_doc_file_path)