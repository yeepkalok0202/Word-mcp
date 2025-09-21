import base64
import os
import tempfile

from docx import Document
from docx.shared import Inches
from fastmcp import FastMCP
from spire.doc import Document
from spire.doc.common import FileForm

# Create an MCP server with a descriptive name
mcp = FastMCP("Word Document Master 📄")


# --- Helper Function ---
# This is the key change: it ensures all file operations happen in a writable temporary directory.
def get_safe_filepath(filename: str) -> str:
    """
    Constructs a full, safe path inside the system's temp directory.
    This prevents directory traversal attacks and ensures the path is writable.
    """
    # Sanitize the filename to prevent security issues like ../../etc/passwd
    safe_filename = os.path.basename(filename)
    # Join it with the system's temporary directory (e.g., /tmp)
    return os.path.join(tempfile.gettempdir(), safe_filename)


@mcp.tool()
def create_document(filename: str) -> str:
    """
    Creates a new, blank Word document in a safe temporary directory on the server.

    Args:
        filename: The base name of the file to create (e.g., 'mydocument.docx').

    Returns:
        A success message.
    """
    document = Document()
    safe_path = get_safe_filepath(filename)  # Use the helper to get a writable path
    document.save(safe_path)
    return f"Document '{filename}' created successfully on the server in a temporary location."


@mcp.tool()
def add_paragraph(filename: str, text: str) -> str:
    """
    Adds a new paragraph to the end of a Word document.

    Args:
        filename: The name of the document to edit.
        text: The text to add.

    Returns:
        A success message.
    """
    safe_path = get_safe_filepath(filename)  # Use the helper
    try:
        document = Document(safe_path)
        document.add_paragraph(text)
        document.save(safe_path)
        return f"Paragraph added to '{filename}'."
    except Exception as e:
        return f"Error adding paragraph to '{filename}': {e}"


@mcp.tool()
def add_heading(filename: str, text: str, level: int = 1) -> str:
    """
    Adds a heading to a Word document.

    Args:
        filename: The name of the document.
        text: The heading text.
        level: The heading level (0-9).

    Returns:
        A success message.
    """
    safe_path = get_safe_filepath(filename)  # Use the helper
    try:
        document = Document(safe_path)
        document.add_heading(text, level=level)
        document.save(safe_path)
        return f"Heading added to '{filename}'."
    except Exception as e:
        return f"Error adding heading to '{filename}': {e}"

@mcp.tool()
def download_document(filename: str) -> str:
    """
    Reads a document from the server's temp directory and returns its content
    as a base64 encoded string for the client to download.

    Args:
        filename: The name of the file on the server (e.g., 'report.docx').

    Returns:
        The base64 encoded content of the file, or an error message if not found.
    """
    safe_path = get_safe_filepath(filename)  # Use the helper
    try:
        with open(safe_path, "rb") as docx_file:
            encoded_string = base64.b64encode(docx_file.read()).decode('utf-8')
        # Clean up the temporary file after it has been read
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
    
    Args:
        filename: The name of the Word document to convert (e.g., 'mydocument.docx').
    
    Returns:
        A success message with the new PDF filename, or an error message if conversion fails.
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