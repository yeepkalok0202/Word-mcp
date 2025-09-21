import base64
import os
import tempfile

from fastmcp import FastMCP
from spire.doc import Document
from spire.doc.common import *

# Create an MCP server with a descriptive name
mcp = FastMCP("Word Document Master 📄")

def get_safe_filepath(filename: str) -> str:
    """
    Constructs a full, safe path inside the system's temp directory.
    This prevents directory traversal attacks and ensures the path is writable.
    """
    safe_filename = os.path.basename(filename)
    return os.path.join(tempfile.gettempdir(), safe_filename)

# NOTE: The create_document, add_paragraph, and add_heading tools have been
#       rewritten to use the Spire.Doc library for consistency.

@mcp.tool()
def create_document(filename: str) -> str:
    """
    Creates a new, blank Word document in a safe temporary directory on the server.
    """
    document = Document()
    safe_path = get_safe_filepath(filename)
    document.SaveToFile(safe_path)
    document.Close()
    return f"Document '{filename}' created successfully on the server in a temporary location."

@mcp.tool()
def add_paragraph(filename: str, text: str) -> str:
    """
    Adds a new paragraph to the end of a Word document.
    """
    safe_path = get_safe_filepath(filename)
    document = Document()
    try:
        document.LoadFromFile(safe_path)
        section = document.Sections[0]
        paragraph = section.AddParagraph()
        paragraph.AppendText(text)
        document.SaveToFile(safe_path)
        return f"Paragraph added to '{filename}'."
    except Exception as e:
        return f"Error adding paragraph to '{filename}': {e}"
    finally:
        document.Close()

@mcp.tool()
def add_heading(filename: str, text: str, level: int = 1) -> str:
    """
    Adds a heading to a Word document.
    """
    safe_path = get_safe_filepath(filename)
    document = Document()
    try:
        document.LoadFromFile(safe_path)
        section = document.Sections[0]
        paragraph = section.AddParagraph()
        paragraph.AppendText(text)
        paragraph.ApplyStyle(f"Heading {level}")
        document.SaveToFile(safe_path)
        return f"Heading added to '{filename}'."
    except Exception as e:
        return f"Error adding heading to '{filename}': {e}"
    finally:
        document.Close()

@mcp.tool()
def download_document(filename: str) -> str:
    """
    Reads a document from the server's temp directory and returns its content
    as a base64 encoded string.
    """
    safe_path = get_safe_filepath(filename)
    try:
        with open(safe_path, "rb") as file_handle:
            encoded_string = base64.b64encode(file_handle.read()).decode('utf-8')
        os.remove(safe_path)
        return encoded_string
    except FileNotFoundError:
        return "Error: File not found on the server."
    except Exception as e:
        return f"Error downloading '{filename}': {e}"

@mcp.tool()
def convert_to_pdf(filename: str) -> str:
    """
    Converts a Word document to a PDF file on the server using Spire.Doc.
    """
    docx_path = get_safe_filepath(filename)
    pdf_path = docx_path.replace('.docx', '.pdf')
    if docx_path == pdf_path:
        pdf_path += '.pdf'
    document = Document()
    try:
        document.LoadFromFile(docx_path)
        document.SaveToFile(pdf_path, FileFormat.PDF)
        return f"Document '{filename}' successfully converted to '{os.path.basename(pdf_path)}'."
    except Exception as e:
        return f"Error converting '{filename}' to PDF: {e}"
    finally:
        document.Close()
        
        

if __name__ == "__main__":
    mcp.run()