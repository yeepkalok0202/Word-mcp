import asyncio
import base64

from fastmcp import Client

# IMPORTANT: Replace with your deployed server's public URL
# If running locally for testing, use "http://127.0.0.1:8000/mcp"
MCP_SERVER_URL = "https://great-ai-word-mcp.fastmcp.app/mcp"

async def main():
    """
    Asynchronously calls the Word MCP server tools to create, edit,
    and then download the final document.
    """
    print(f"🤖 Connecting to FastMCP server at {MCP_SERVER_URL}...")
    client = Client(MCP_SERVER_URL)
    
    # Define filenames for the server and your local machine
    server_filename = "remote_report.docx"
    local_filename = "my_local_report.pdf"  # <-- CHANGED TO PDF

    # The 'async with' block handles connecting and disconnecting gracefully
    async with client:
        try:
            # --- Step 1: Operations to create and modify the file ON THE SERVER ---
            print(f"📞 Creating '{server_filename}' on the server...")
            result = await client.call_tool("create_document", {"filename": server_filename})
            print(f"✅ Server response: {result.content[0].text}\n")

            print(f"📞 Adding heading to '{server_filename}' on the server...")
            heading_args = {"filename": server_filename, "text": "Quarterly Sales Report", "level": 1}
            result = await client.call_tool("add_heading", heading_args)
            print(f"✅ Server response: {result.content[0].text}\n")

            print(f"📞 Adding paragraph to '{server_filename}' on the server...")
            paragraph_args = {
                "filename": server_filename,
                "text": "This document outlines the sales performance for the second quarter of 2025."
            }
            result = await client.call_tool("add_paragraph", paragraph_args)
            print(f"✅ Server response: {result.content[0].text}\n")

            # --- NEW STEP: CONVERT THE DOCX TO PDF ON THE SERVER ---
            print(f"🔄 Converting '{server_filename}' to PDF on the server...")
            convert_args = {"filename": server_filename}
            result = await client.call_tool("convert_to_pdf", convert_args)
            print(f"✅ Server response: {result.content[0].text}\n")

            # --- Step 2: DOWNLOAD the final file from the server to your local machine ---
            # NOTE: We now download the PDF file
            server_pdf_filename = server_filename.replace('.docx', '.pdf')
            print(f"⬇️ Calling 'download_document' for '{server_pdf_filename}'...")
            download_args = {"filename": server_pdf_filename} # <-- CHANGED TO PDF
            result = await client.call_tool("download_document", download_args)
            
            # Access the .text attribute of the first content object
            base64_content = result.content[0].text
            
            if "Error:" not in base64_content:
                print(f"💾 Decoding and saving file to '{local_filename}'...")
                # Decode the base64 string back into binary data
                file_bytes = base64.b64decode(base64_content)
                
                # Write the binary data to a new local file
                with open(local_filename, "wb") as f:
                    f.write(file_bytes)
                print(f"🎉 Success! The final document has been saved locally as '{local_filename}'.")
            else:
                print(f"❌ Server error: {base64_content}")

        except Exception as e:
            print(f"\n❌ An error occurred during the tool call.")
            print(f"Error details: {e}")
            print("Please ensure the server is running and the MCP_SERVER_URL is correct.")
            return

# --- This is the standard way to run an async Python script ---
if __name__ == "__main__":
    asyncio.run(main())