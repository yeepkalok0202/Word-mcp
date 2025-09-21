from spire.doc import *
from spire.doc.common import *

# Create a Document object
document = Document()
# Load a Word DOCX file
document.LoadFromFile("my_local_report.docx")
# Or load a Word DOC file
#document.LoadFromFile("Sample.doc")

# Save the file to a PDF file
document.SaveToFile("my_local_report.pdf", FileFormat.PDF)
document.Close()