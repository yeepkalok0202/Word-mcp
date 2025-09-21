import asyncio

from fastmcp import Client

MCP_SERVER_URL= "https://great-ai-word-mcp.fastmcp.app/mcp"

async def main():
    print("connecting to fastmcp")
    client= Client(MCP_SERVER_URL)
    
    async with client:
        try:
            # 1. Create a new document named 'report.docx'
            print("📞 Calling 'create_document'...")
            create_args = {"filename": "test.docx"}
            result = await client.call_tool("create_document", create_args)
            # The result from the client is often a structured object, we'll print its content
            print(f"✅ Server response: {result.content}\n")

            # 2. Add a main heading to the document
            print("📞 Calling 'add_heading'...")
            heading_args = {"filename": "test.docx", "text": "Quarterly Sales Report", "level": 1}
            result = await client.call_tool("add_heading", heading_args)
            print(f"✅ Server response: {result.content}\n")

            # 3. Add a paragraph of text
            print("📞 Calling 'add_paragraph'...")
            paragraph_args = {
                "filename": "test.docx",
                "text": "This document outlines the sales performance for the second quarter of 2025."
            }
            result = await client.call_tool("add_paragraph", paragraph_args)
            print(f"✅ Server response: {result.content}\n")

        except Exception as e:
            print(f"\n❌ An error occurred during the tool call.")
            print(f"Error details: {e}")
            print("Please ensure the server is running and accessible.")
            return

    print("🎉 Done! Check for 'report.docx' in your folder.")

# --- This is the standard way to run an async Python script ---
if __name__ == "__main__":
    asyncio.run(main())