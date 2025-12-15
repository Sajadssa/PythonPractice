# تست در Python
import sys
print("Python version:", sys.version)

try:
    import mcp
    print("✓ MCP installed")
except:
    print("✗ MCP not installed")

try:
    import pyodbc
    print("✓ pyodbc installed")
    print("  ODBC Drivers:", pyodbc.drivers())
except:
    print("✗ pyodbc not installed")

try:
    import win32com.client
    print("✓ pywin32 installed")
except:
    print("✗ pywin32 not installed")