import os
import json
from http.server import BaseHTTPRequestHandler

# Directory where templates are stored
TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "../templates")
TEMPLATE_DIR = os.path.abspath(TEMPLATE_DIR)

class handler(BaseHTTPRequestHandler):

    def do_GET(self):
        try:
                # List all files in the template directory
                files = os.listdir(TEMPLATE_DIR)
                
                # Filter for .xlsx files
                excel_files = [file for file in files if file.endswith('.xlsx')]
                
                # Return the list of templates as JSON
                response = json.dumps({"templates": excel_files})
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(response.encode('utf-8'))
        except Exception as e:
            print(f"Error reading template directory: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            error_response = json.dumps({"error": "Unable to load templates"})
            self.wfile.write(error_response.encode('utf-8'))