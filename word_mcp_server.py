from docx import Document
from docx.shared import Inches
from fastmcp import FastMCP

mcp = FastMCP("Word Document MCP Server")

@mcp.tool()
def create_document(filename: str)-> str:
    document= Document()
    document.save(filename)
    return f"Document `{filename}` created successfully"

@mcp.tool()
def add_paragraph(filename: str, text: str)-> str:
    document= Document(filename)
    document.add_paragraph(text)
    document.save(filename)
    return f"Paragraph added to `{filename}` successfully."

@mcp.tool()
def add_heading(filename: str, text: str, level: int=1 )-> str:
    '''
    level: The heading level (0-9)
    '''
    document= Document(filename)
    document.add_heading(text, level=level)
    document.save(filename)
    return f"Heading added to `{filename}` successfully."

@mcp.tool()
def add_table(filename: str, rows:int, cols:int)-> str:
    document= Document(filename)
    document.add_table(rows=rows, cols=cols)
    document.save(filename)
    return f"Table added to `{filename}` successfully."

@mcp.tool()
def add_picture(filename: str, image_path:str, width_inches:float=2.5)-> str:
    document= Document(filename)
    document.add_picture(image_path, width=Inches(width_inches))
    document.save(filename)
    return f"Paragraph added to `{filename}` successfully."

if __name__== "__main__":
    mcp.run()