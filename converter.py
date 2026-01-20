from pyhwpx import Hwp
import os
import time
import threading

class HwpConverter:
    def __init__(self, logger=None):
        self.logger = logger
        self.hwp = None

    def log(self, message, level="info"):
        if self.logger:
            getattr(self.logger, level)(message)
        else:
            print(f"[{level.upper()}] {message}")

    def convert_to_hwpx(self, input_path):
        input_path = os.path.abspath(input_path)
        if not os.path.exists(input_path):
            self.log(f"File not found: {input_path}", "error")
            return False

        file_name, ext = os.path.splitext(input_path)
        if ext.lower() != '.hwp':
            self.log(f"Skipping non-HWP file: {input_path}", "warning")
            return False

        output_path = f"{file_name}.hwpx"
        
        try:
            self.log(f"Initializing Hancom Office...")
            # pyhwpx Hwp() automatically handles initialization and visible=True/False
            self.hwp = Hwp(visible=False)
            
            self.log(f"Opening file: {input_path}")
            self.hwp.open(input_path)
            
            self.log(f"Saving as: {output_path}")
            
            # Try saving with "HWPX" format
            formats_to_try = ["HWPX", "KB_HWPX", "HWPML2X"] 
            success = False
            
            for fmt in formats_to_try:
                self.log(f"Attempting conversion with format: {fmt}")
                if self.hwp.save_as(output_path, format=fmt):
                    # Verify signature
                    if self.is_zip_file(output_path):
                        self.log(f"Conversion successful with format {fmt}: {output_path}")
                        success = True
                        break
                    else:
                        self.log(f"Warning: save_as returned True but file is not a valid ZIP (HWPX). It might be legacy binary. Retrying...", "warning")
                else:
                     self.log(f"save_as({fmt}) returned False.")
            
            self.hwp.quit()
            return success

        except Exception as e:
            self.log(f"Error during conversion: {str(e)}", "error")
            if self.hwp:
                try:
                    self.hwp.quit()
                except:
                    pass
            return False

    def is_zip_file(self, filepath):
        try:
            with open(filepath, 'rb') as f:
                header = f.read(4)
            return header == b'\x50\x4b\x03\x04'
        except Exception:
            return False

    def convert_directory(self, input_dir):
        input_dir = os.path.abspath(input_dir)
        if not os.path.exists(input_dir):
            self.log(f"Directory not found: {input_dir}", "error")
            return

        self.log(f"Scanning directory: {input_dir}")
        count = 0
        for root, dirs, files in os.walk(input_dir):
            for file in files:
                if file.lower().endswith('.hwp'):
                    full_path = os.path.join(root, file)
                    # Skip if hwpx already exists? Optional.
                    if self.convert_to_hwpx(full_path):
                        count += 1
                        
        self.log(f"Batch conversion complete. Converted {count} files.")
