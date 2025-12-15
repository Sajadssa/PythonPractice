"""
MCP Server for IDMS Access Database
This server allows Claude to connect to your Access app and SQL Server

Requirements:
- Python 3.8+
- pip install mcp pyodbc win32com flask

Setup:
1. Install this on a Windows machine that has Access installed
2. Configure the paths and connection strings below
3. Run: python mcp_server_idms.py
4. Add MCP server to Claude Desktop settings
"""

import asyncio
import pyodbc
import win32com.client
from pathlib import Path
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# ============================================
# CONFIGURATION - Update these values
# ============================================
ACCESS_DB_PATH = r"C:\Users\SaeeidiAzad\Desktop\IDMS_Rev_2.1.2.accdb"
SQL_CONNECTION = "DRIVER={SQL Server};SERVER=DCC-SAEEDI;DATABASE=IDMS_WRFM;Trusted_Connection=yes"

# ============================================
# Initialize MCP Server
# ============================================
app = Server("idms-access-server")

# ============================================
# Tool 1: Sync Forms to SQL
# ============================================
@app.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="sync_forms",
            description="Sync all Access forms to SQL Server automatically",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        Tool(
            name="query_sql",
            description="Execute SQL query on IDMS database",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "SQL query to execute"
                    }
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="list_forms",
            description="List all forms in Access database",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        Tool(
            name="get_form_code",
            description="Get VBA code from a specific form",
            inputSchema={
                "type": "object",
                "properties": {
                    "form_name": {
                        "type": "string",
                        "description": "Name of the form"
                    }
                },
                "required": ["form_name"]
            }
        ),
        Tool(
            name="execute_vba",
            description="Execute a VBA function in Access",
            inputSchema={
                "type": "object",
                "properties": {
                    "function_name": {
                        "type": "string",
                        "description": "VBA function name to execute"
                    }
                },
                "required": ["function_name"]
            }
        )
    ]

# ============================================
# Tool Implementation: Sync Forms
# ============================================
@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    
    if name == "sync_forms":
        try:
            # Open Access database
            access = win32com.client.Dispatch("Access.Application")
            access.OpenCurrentDatabase(ACCESS_DB_PATH)
            
            # Run the sync function
            result = access.Run("SyncFormsToSQL")
            
            access.CloseCurrentDatabase()
            access.Quit()
            
            return [TextContent(
                type="text",
                text=f"Forms synced successfully! Result: {result}"
            )]
        except Exception as e:
            return [TextContent(
                type="text",
                text=f"Error syncing forms: {str(e)}"
            )]
    
    elif name == "query_sql":
        try:
            query = arguments.get("query", "")
            
            conn = pyodbc.connect(SQL_CONNECTION)
            cursor = conn.cursor()
            cursor.execute(query)
            
            # Fetch results
            if cursor.description:
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                
                # Format results
                result = f"Columns: {', '.join(columns)}\n\n"
                for row in rows[:100]:  # Limit to 100 rows
                    result += f"{row}\n"
                
                if len(rows) > 100:
                    result += f"\n... and {len(rows) - 100} more rows"
            else:
                result = f"Query executed successfully. Rows affected: {cursor.rowcount}"
            
            conn.close()
            
            return [TextContent(type="text", text=result)]
        except Exception as e:
            return [TextContent(type="text", text=f"Error: {str(e)}")]
    
    elif name == "list_forms":
        try:
            access = win32com.client.Dispatch("Access.Application")
            access.OpenCurrentDatabase(ACCESS_DB_PATH)
            
            forms = []
            for i in range(access.CurrentProject.AllForms.Count):
                form_obj = access.CurrentProject.AllForms.Item(i)
                forms.append(form_obj.Name)
            
            access.CloseCurrentDatabase()
            access.Quit()
            
            result = f"Total forms: {len(forms)}\n\n"
            result += "\n".join(forms)
            
            return [TextContent(type="text", text=result)]
        except Exception as e:
            return [TextContent(type="text", text=f"Error: {str(e)}")]
    
    elif name == "get_form_code":
        try:
            form_name = arguments.get("form_name", "")
            
            access = win32com.client.Dispatch("Access.Application")
            access.OpenCurrentDatabase(ACCESS_DB_PATH)
            
            # Get form module code
            form = access.CurrentProject.AllForms(form_name)
            if form.IsLoaded:
                module = access.Forms(form_name).Module
                code = module.Lines(1, module.CountOfLines)
            else:
                code = "Form is not loaded or has no code module"
            
            access.CloseCurrentDatabase()
            access.Quit()
            
            return [TextContent(type="text", text=code)]
        except Exception as e:
            return [TextContent(type="text", text=f"Error: {str(e)}")]
    
    elif name == "execute_vba":
        try:
            function_name = arguments.get("function_name", "")
            
            access = win32com.client.Dispatch("Access.Application")
            access.OpenCurrentDatabase(ACCESS_DB_PATH)
            
            result = access.Run(function_name)
            
            access.CloseCurrentDatabase()
            access.Quit()
            
            return [TextContent(
                type="text",
                text=f"Function '{function_name}' executed. Result: {result}"
            )]
        except Exception as e:
            return [TextContent(type="text", text=f"Error: {str(e)}")]
    
    return [TextContent(type="text", text=f"Unknown tool: {name}")]

# ============================================
# Run Server
# ============================================
async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )

if __name__ == "__main__":
    print("Starting IDMS MCP Server...")
    print(f"Access DB: {ACCESS_DB_PATH}")
    print(f"SQL Server: {SQL_CONNECTION}")
    asyncio.run(main())