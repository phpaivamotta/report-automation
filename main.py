from dotenv import load_dotenv
import os
from docx import Document

from utils import add_table_with_images
from utils import add_formatted_bullets
from utils import add_bullets_above_tables
from utils import add_caption_paragraph
from utils import build_cross_reference_bullet
from utils import remove_empty_paragraphs_after_table
from utils import remove_first_empty_paragraph_above_text
from utils import insert_formatted_text_after_header
from utils import read_report_data
from utils import read_picture_data
from utils import update_document_properties

from wordextraction import get_next_report_id


if __name__ == "__main__":

    load_dotenv(override=True)
    template_file_path = os.getenv('TEMPLATE_DOC_PATH')
    output_folder_path = os.getenv('OUTPUT_REPORT_FOLDER_PATH')
    extracted_data_path = os.getenv('EXTRACTED_DATA_PATH')

    report_csv_path = os.path.join(extracted_data_path, 'report_info.csv')
    picture_csv_path = os.path.join(extracted_data_path, 'picture_info.csv')

    latest_report_id = get_next_report_id(report_csv_path) - 1

    while True:
        try:
            report_id = int(input(f"Enter the report ID to generate (latest is {latest_report_id}): "))
            if 1 <= report_id <= latest_report_id:
                break
            else:
                print(f"Please enter a valid report ID between 1 and {latest_report_id}.")
        except ValueError:
            print("Please enter a valid integer.")

    report_data = read_report_data(report_csv_path, report_id)

    if not report_data:
        print(f"No data found for report ID {report_id}")
    else:
        picture_data = read_picture_data(picture_csv_path, report_id)

        output_file_name = f"Report_{report_id:04d}_{report_data['Customer'].replace(' ', '_')}.docx"
        output_doc_file_path = os.path.join(output_folder_path, output_file_name)

        if not os.path.exists(template_file_path):
            print(f"Template file not found: {template_file_path}")
        else:
            # Copy template then open once — all operations happen in memory
            template_doc = Document(template_file_path)
            template_doc.save(output_doc_file_path)
            doc = Document(output_doc_file_path)

            update_document_properties(doc, report_data)

            insert_formatted_text_after_header(doc, "Introduction", report_data['Introduction'])
            insert_formatted_text_after_header(doc, "Entrance Meeting", report_data['Entrance Meeting'])
            insert_formatted_text_after_header(doc, "Inspection Conclusions and Recommendations", report_data['Conclusions'])

            drawings = report_data['Drawings Used'].split('; ')
            specifications = report_data['Specifications Used'].split('; ')

            add_formatted_bullets(
                doc,
                "The following drawings were provided and used during the inspection:",
                drawings,
                is_drawing=True,
            )
            add_formatted_bullets(
                doc,
                "The following specifications were used during the inspection:",
                specifications,
                is_drawing=False,
            )

            if picture_data:
                report_folder = os.path.join(extracted_data_path, f"report_{report_id:04d}")
                table_counter = 0
                figure_counter = 1

                for i in range(0, len(picture_data), 2):
                    image_1 = picture_data[i]
                    image_path_1 = os.path.join(report_folder, image_1['Picture File Name'])
                    description_1 = image_1['Description']
                    caption_1 = image_1['Caption']

                    if i + 1 < len(picture_data):
                        image_2 = picture_data[i + 1]
                        image_path_2 = os.path.join(report_folder, image_2['Picture File Name'])
                        description_2 = image_2['Description']
                        caption_2 = image_2['Caption']
                        num_cols = 2
                    else:
                        image_path_2 = None
                        description_2 = None
                        caption_2 = None
                        num_cols = 1

                    table = add_table_with_images(
                        doc, "Inspection Observations:", table_counter,
                        num_cols, image_path_1, image_path_2,
                    )

                    # Insert each caption as para[1] inside its cell — matching
                    # Word's native InsertCaption behaviour (caption lives in the
                    # same cell as the image, not as a body-level paragraph).
                    bm_1, _ = add_caption_paragraph(
                        table.cell(0, 0).paragraphs[0]._p, caption_1, figure_counter
                    )
                    fig_1 = figure_counter
                    figure_counter += 1

                    bm_2 = fig_2 = None
                    if num_cols == 2:
                        bm_2, _ = add_caption_paragraph(
                            table.cell(0, 1).paragraphs[0]._p, caption_2, figure_counter
                        )
                        fig_2 = figure_counter
                        figure_counter += 1

                    bullets = add_bullets_above_tables(doc, table, num_cols)

                    build_cross_reference_bullet(bullets[0], description_1, bm_1, fig_1)
                    if num_cols == 2 and len(bullets) > 1:
                        build_cross_reference_bullet(bullets[1], description_2, bm_2, fig_2)

                    table_counter += 1

            remove_empty_paragraphs_after_table(doc)
            remove_first_empty_paragraph_above_text(doc, "Inspection Observations:")

            # Single save — document was held in memory the entire time
            doc.save(output_doc_file_path)
            print(f"Report generated successfully: {output_file_name}")
