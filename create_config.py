import os
import json

# مسیر فایل config
config_dir = os.path.join(os.getenv('APPDATA'), 'Claude')
config_file = os.path.join(config_dir, 'claude_desktop_config.json')

# ایجاد پوشه
os.makedirs(config_dir, exist_ok=True)

# محتوای config
config = {
    "mcpServers": {
        "idms": {
            "command": "python",
            "args": [
                r"D:\Sepher_Pasargad\works\Maintenace\PythonDataAnalysis\PythonPractice\idms_mcp_server.py"
            ]
        }
    }
}

# نوشتن فایل
with open(config_file, 'w', encoding='utf-8') as f:
    json.dump(config, f, indent=2)

print(f"✓ Config file created: {config_file}")
print("\nContent:")
print(json.dumps(config, indent=2))