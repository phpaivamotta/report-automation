import win32com.client

doc_path = r"C:\Users\phpai\OneDrive\Desktop\report-automation\Template Report\Sample Report_DF_PM.docx"

# Open Word application
word = win32com.client.Dispatch("Word.Application")

# Open the document (provide the full path to the document)
doc = word.Documents.Open(doc_path)

# print("Read Properties\n")

# Access core (built-in) properties
core_props = doc.BuiltInDocumentProperties
# print("\nCore Properties:\n")
# print("Title:", core_props("Title"))
# print("Author:", core_props("Author"))
# print("Subject:", core_props("Subject"))
# print("Keywords:", core_props("Keywords"))

# Access custom properties
custom_props = doc.CustomDocumentProperties
# print("\nCustom Properties:\n")
# print("Customer Address:", custom_props("customer address"))
# print("Customer Contact:", custom_props("customer contact"))
# print("Inspection Site:", custom_props("inspection site"))
# print("Customer PO Number:", custom_props("customer po num"))
# print("Customer CCs:", custom_props("customer ccs"))
# print("Inspection Date:", custom_props("inspection date"))
# print("Company:", custom_props("company"))
# print("Maverick Contact Info:", custom_props("maverick contact info cell"))
# print("Maverick Contact Info:", custom_props("maverick contact info email"))
# print("Author Title:", custom_props("author title"))
# print("Maverick CCs:", custom_props("maverick ccs"))
# print("Report Date:", custom_props("report date"))

# Access and modify core properties
core_props("Title").Value = "Clifty"
core_props("Author").Value = "Pedro Motta"
core_props("Subject").Value = "Inspection"
core_props("Keywords").Value = "20243524"

# Access and modify custom properties
custom_props("customer address").Value = "Abbey Road 101"
# custom_props("customer contact").Value = "John Doe"
custom_props("inspection site").Value = "Clifty Creek Plant"
custom_props("customer po num").Value = "34004555"
custom_props("customer ccs").Value = "Jane Doe"
custom_props("inspection date").Value = "10/27/2024"
# custom_props("company").Value = "Maverick"
# custom_props("maverick contact info cell").Value = "696-2424-420"
# custom_props("maverick contact info email").Value = "myemail@gmail.com"
# custom_props("author title").Value = "O Foda"
custom_props("maverick ccs").Value = "Pedro Motta"
custom_props("report date").Value = "11/20/2024"
custom_props("customer contacts").Value = "John Doe"

# # Access and modify core properties
# core_props("Title").Value = report_data['Customer']
# core_props("Author").Value = report_data['From']
# core_props("Subject").Value = report_data['Subject']
# core_props("Keywords").Value = report_data['Maverick Job']
# core_props("Comments").Value = report_data['Customer Contact']

# # Access custom properties
# custom_props = doc.CustomDocumentProperties

# # Access and modify custom properties
# custom_props("customer address").Value = report_data['Customer Address']
# custom_props("inspection site").Value = report_data['Inspection Site']
# custom_props("customer po num").Value = report_data['Customer PO No.']
# custom_props("customer ccs").Value = report_data['Customer CCs']
# custom_props("inspection date").Value = report_data['Inspection Date(s)']
# custom_props("maverick ccs").Value = report_data['Maverick CCs']
# custom_props("report date").Value = report_data['Report Date']

# Access core (built-in) properties
# print("\nCore Properties:\n")
# print("Title:", core_props("Title"))
# print("Author:", core_props("Author"))
# print("Subject:", core_props("Subject"))
# print("Keywords:", core_props("Keywords"))

# # Access custom properties
# print("\nCustom Properties:\n")
# print("Customer Address:", custom_props("customer address"))
# print("Customer Contact:", custom_props("customer contact"))
# print("Inspection Site:", custom_props("inspection site"))
# print("Customer PO Number:", custom_props("customer po num"))
# print("Customer CCs:", custom_props("customer ccs"))
# print("Inspection Date:", custom_props("inspection date"))
# print("Company:", custom_props("company"))
# print("Maverick Contact Info:", custom_props("maverick contact info cell"))
# print("Maverick Contact Info:", custom_props("maverick contact info email"))
# print("Author Title:", custom_props("author title"))
# print("Maverick CCs:", custom_props("maverick ccs"))
# print("Report Date:", custom_props("report date"))

# Close the document and quit Word
# doc.Save()
doc.Close()
word.Quit()