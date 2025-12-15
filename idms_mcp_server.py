"""
IDMS MCP Server
Allows Claude to directly interact with your Access database and SQL Server

Installation:
1. pip install mcp pyodbc pywin32
2. Update configuration below
3. Run: python idms_mcp_server.py

Add to Claude Desktop config:
{
  "mcpServers": {
    "idms": {
      "command": "python",
      "args": ["D:\Sepher_Pasargad\works\Maintenace\PythonDataAnalysis\PythonPractice\idms_mcp_server.py"]
    }
  }
}
"""

import asyncio
import json
import pyodbc
import win32com.client
from pathlib import Path
from typing import Any, Sequence
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import (
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource,
)

# ============================================
# CONFIGURATION - UPDATE THESE
# ============================================
CONFIG = {
    "access_db": r"C:\Users\SaeeidiAzad\Desktop\IDMS_Rev_2.1.2.accdb",
    "sql_server": "DCC-SAEEDI",
    "sql_database": "IDMS_WRFM",
    "sql_username": None,  # None for Windows Auth
    "sql_password": None,
}

# Build SQL connection string
if CONFIG["sql_username"]:
    SQL_CONN = f"DRIVER={{SQL Server}};SERVER={CONFIG['sql_server']};DATABASE={CONFIG['sql_database']};UID={CONFIG['sql_username']};PWD={CONFIG['sql_password']}"
else:
    SQL_CONN = f"DRIVER={{SQL Server}};SERVER={CONFIG['sql_server']};DATABASE={CONFIG['sql_database']};Trusted_Connection=yes"

# ============================================
# Initialize MCP Server
# ============================================
app = Server("idms-access-server")

# ============================================
# Helper Functions
# ============================================
def get_sql_connection():
    """Get SQL Server connection"""
    return pyodbc.connect(SQL_CONN)

def get_access_app():
    """Get Access Application object"""
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False
    return access

def format_table_results(columns, rows, max_rows=50):
    """Format SQL results as readable text"""
    if not rows:
        return "No results found"
    
    result = f"Columns: {', '.join(columns)}\n"
    result += "=" * 80 + "\n\n"
    
    for idx, row in enumerate(rows[:max_rows], 1):
        result += f"Row {idx}:\n"
        for col, val in zip(columns, row):
            result += f"  {col}: {val}\n"
        result += "\n"
    
    if len(rows) > max_rows:
        result += f"\n... and {len(rows) - max_rows} more rows\n"
    
    return result

# ============================================
# Define MCP Tools
# ============================================
@app.list_tools()
async def list_tools() -> list[Tool]:
    """List all available tools"""
    return [
        Tool(
            name="sql_query",
            description="Execute SQL query on IDMS database (SELECT, UPDATE, INSERT, DELETE)",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "SQL query to execute"
                    },
                    "max_rows": {
                        "type": "integer",
                        "description": "Maximum rows to return (default: 50)",
                        "default": 50
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
                "properties": {
                    "filter": {
                        "type": "string",
                        "description": "Optional filter (e.g., 'Pi_', 'Cor_')"
                    }
                }
            }
        ),
        
        Tool(
            name="get_form_info",
            description="Get detailed information about a specific form",
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
                        "description": "VBA function name"
                    },
                    "parameters": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Optional parameters"
                    }
                },
                "required": ["function_name"]
            }
        ),
        
        Tool(
            name="sync_forms",
            description="Sync all Access forms to SQL Server",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        
        Tool(
            name="get_user_permissions",
            description="Get permissions for a specific user",
            inputSchema={
                "type": "object",
                "properties": {
                    "username": {
                        "type": "string",
                        "description": "Username"
                    }
                },
                "required": ["username"]
            }
        ),
        
        Tool(
            name="set_user_permission",
            description="Set permission for a user on a specific form",
            inputSchema={
                "type": "object",
                "properties": {
                    "username": {
                        "type": "string",
                        "description": "Username"
                    },
                    "form_name": {
                        "type": "string",
                        "description": "Form name"
                    },
                    "permission_type": {
                        "type": "string",
                        "enum": ["open", "edit", "delete", "add", "show"],
                        "description": "Permission type"
                    },
                    "allow": {
                        "type": "boolean",
                        "description": "True to grant, False to revoke"
                    }
                },
                "required": ["username", "form_name", "permission_type", "allow"]
            }
        ),
        
        Tool(
            name="get_database_stats",
            description="Get statistics about the IDMS database",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
    ]

# ============================================
# Tool Implementations
# ============================================
@app.call_tool()
async def call_tool(name: str, arguments: dict) -> Sequence[TextContent | ImageContent | EmbeddedResource]:
    """Handle tool calls"""
    
    try:
        # SQL Query
        if name == "sql_query":
            query = arguments.get("query", "")
            max_rows = arguments.get("max_rows", 50)
            
            conn = get_sql_connection()
            cursor = conn.cursor()
            cursor.execute(query)
            
            if cursor.description:
                columns = [col[0] for col in cursor.description]
                rows = cursor.fetchall()
                result = format_table_results(columns, rows, max_rows)
                result += f"\nTotal rows: {len(rows)}"
            else:
                result = f"Query executed. Rows affected: {cursor.rowcount}"
            
            conn.close()
            
            return [TextContent(type="text", text=result)]
        
        # List Forms
        elif name == "list_forms":
            filter_str = arguments.get("filter", "")
            
            access = get_access_app()
            access.OpenCurrentDatabase(CONFIG["access_db"])
            
            forms = []
            for i in range(access.CurrentProject.AllForms.Count):
                form_name = access.CurrentProject.AllForms.Item(i).Name
                if not (form_name.startswith('MSys') or form_name.startswith('~')):
                    if not filter_str or form_name.startswith(filter_str):
                        forms.append(form_name)
            
            access.CloseCurrentDatabase()
            access.Quit()
            
            result = f"Found {len(forms)} forms"
            if filter_str:
                result += f" (filtered by '{filter_str}')"
            result += ":\n\n" + "\n".join(f"{i+1}. {f}" for i, f in enumerate(forms))
            
            return [TextContent(type="text", text=result)]
        
        # Get Form Info
        elif name == "get_form_info":
            form_name = arguments.get("form_name", "")
            
            conn = get_sql_connection()
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT FormID, FormName, FormCategory, DisciplineCode, IsActive
                FROM tbl_Forms
                WHERE FormName = ?
            """, form_name)
            
            row = cursor.fetchone()
            
            if row:
                result = f"Form Information:\n"
                result += f"  ID: {row[0]}\n"
                result += f"  Name: {row[1]}\n"
                result += f"  Category: {row[2]}\n"
                result += f"  Discipline: {row[3]}\n"
                result += f"  Active: {row[4]}\n"
            else:
                result = f"Form '{form_name}' not found in database"
            
            conn.close()
            
            return [TextContent(type="text", text=result)]
        
        # Execute VBA
        elif name == "execute_vba":
            function_name = arguments.get("function_name", "")
            
            access = get_access_app()
            access.OpenCurrentDatabase(CONFIG["access_db"])
            
            result_val = access.Run(function_name)
            
            access.CloseCurrentDatabase()
            access.Quit()
            
            return [TextContent(type="text", text=f"VBA function '{function_name}' executed. Result: {result_val}")]
        
        # Sync Forms
        elif name == "sync_forms":
            # Call the sync function we created earlier
            access = get_access_app()
            access.OpenCurrentDatabase(CONFIG["access_db"])
            
            result_val = access.Run("SyncFormsToSQL")
            
            access.CloseCurrentDatabase()
            access.Quit()
            
            return [TextContent(type="text", text=f"Forms synced successfully! Result: {result_val}")]
        
        # Get User Permissions
        elif name == "get_user_permissions":
            username = arguments.get("username", "")
            
            conn = get_sql_connection()
            cursor = conn.cursor()
            
            cursor.execute("""
                SELECT u.UserID
                FROM tbl_Users u
                WHERE u.UserName = ?
            """, username)
            
            user_row = cursor.fetchone()
            
            if not user_row:
                conn.close()
                return [TextContent(type="text", text=f"User '{username}' not found")]
            
            user_id = user_row[0]
            
            cursor.execute("""
                SELECT 
                    f.FormName,
                    p.CanOpen,
                    p.CanEdit,
                    p.CanDelete,
                    p.CanAdd,
                    p.CanShow
                FROM tbl_UserFormPermissions p
                JOIN tbl_Forms f ON p.FormID = f.FormID
                WHERE p.UserID = ?
                ORDER BY f.FormName
            """, user_id)
            
            rows = cursor.fetchall()
            
            if rows:
                result = f"Permissions for user '{username}':\n\n"
                for row in rows:
                    result += f"{row[0]}:\n"
                    result += f"  Open: {row[1]}, Edit: {row[2]}, Delete: {row[3]}, Add: {row[4]}, Show: {row[5]}\n"
            else:
                result = f"No permissions set for user '{username}'"
            
            conn.close()
            
            return [TextContent(type="text", text=result)]
        
        # Set User Permission
        elif name == "set_user_permission":
            username = arguments.get("username", "")
            form_name = arguments.get("form_name", "")
            perm_type = arguments.get("permission_type", "")
            allow = arguments.get("allow", False)
            
            conn = get_sql_connection()
            cursor = conn.cursor()
            
            # Get UserID and FormID
            cursor.execute("SELECT UserID FROM tbl_Users WHERE UserName = ?", username)
            user_row = cursor.fetchone()
            
            cursor.execute("SELECT FormID FROM tbl_Forms WHERE FormName = ?", form_name)
            form_row = cursor.fetchone()
            
            if not user_row or not form_row:
                conn.close()
                return [TextContent(type="text", text="User or Form not found")]
            
            user_id = user_row[0]
            form_id = form_row[0]
            
            # Map permission type to column
            perm_map = {
                "open": "CanOpen",
                "edit": "CanEdit",
                "delete": "CanDelete",
                "add": "CanAdd",
                "show": "CanShow"
            }
            
            col_name = perm_map.get(perm_type)
            
            if not col_name:
                conn.close()
                return [TextContent(type="text", text=f"Invalid permission type: {perm_type}")]
            
            # Update or Insert permission
            cursor.execute(f"""
                MERGE INTO tbl_UserFormPermissions AS target
                USING (SELECT ? AS UserID, ? AS FormID) AS source
                ON target.UserID = source.UserID AND target.FormID = source.FormID
                WHEN MATCHED THEN
                    UPDATE SET {col_name} = ?
                WHEN NOT MATCHED THEN
                    INSERT (UserID, FormID, {col_name})
                    VALUES (?, ?, ?);
            """, user_id, form_id, int(allow), user_id, form_id, int(allow))
            
            conn.commit()
            conn.close()
            
            action = "granted" if allow else "revoked"
            return [TextContent(type="text", text=f"Permission '{perm_type}' {action} for user '{username}' on form '{form_name}'")]
        
        # Get Database Stats
        elif name == "get_database_stats":
            conn = get_sql_connection()
            cursor = conn.cursor()
            
            # Get counts
            cursor.execute("SELECT COUNT(*) FROM tbl_Users")
            user_count = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM tbl_Forms")
            form_count = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM tbl_UserFormPermissions")
            perm_count = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM tbl_Disciplines")
            disc_count = cursor.fetchone()[0]
            
            # Get forms by discipline
            cursor.execute("""
                SELECT DisciplineCode, COUNT(*) as cnt
                FROM tbl_Forms
                GROUP BY DisciplineCode
                ORDER BY cnt DESC
            """)
            disc_stats = cursor.fetchall()
            
            conn.close()
            
            result = "IDMS Database Statistics:\n"
            result += "=" * 50 + "\n\n"
            result += f"Total Users:       {user_count}\n"
            result += f"Total Forms:       {form_count}\n"
            result += f"Total Permissions: {perm_count}\n"
            result += f"Total Disciplines: {disc_count}\n\n"
            result += "Forms by Discipline:\n"
            for disc, cnt in disc_stats:
                result += f"  {disc:<8}: {cnt:>3} forms\n"
            
            return [TextContent(type="text", text=result)]
        
        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]
    
    except Exception as e:
        error_msg = f"Error executing {name}: {str(e)}"
        return [TextContent(type="text", text=error_msg)]

# ============================================
# Run Server
# ============================================
async def main():
    """Main entry point"""
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )

if __name__ == "__main__":
    print("Starting IDMS MCP Server...")
    print(f"Access DB: {CONFIG['access_db']}")
    print(f"SQL Server: {CONFIG['sql_server']}")
    print("Ready for connections...")
    asyncio.run(main())
