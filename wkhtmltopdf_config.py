import os

# Detect if running in Docker
if os.path.exists('/usr/bin/wkhtmltopdf'):
    WKHTMLTOPDF_CMD = '/usr/bin/wkhtmltopdf'
else:
    WKHTMLTOPDF_CMD = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
