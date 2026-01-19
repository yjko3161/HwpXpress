import os
import time
import win32com.client
import pythoncom
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
        
        # Initialize COM in this thread
        pythoncom.CoInitialize()
        
        try:
            self.log(f"Initializing Hancom Office...")
            self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") # Security module bypass if needed/possible
            
            self.log(f"Opening file: {input_path}")
            self.hwp.Open(input_path)
            
            self.log(f"Saving as: {output_path}")
            # Format "HWPX" is usually "HWPX" or equivalent in HWP automation API. 
            # Often it's HWP format "download". 
            # For HWPX specifically, we verify the format string.
            # Using SaveAs with format="HWPX"
            success = self.hwp.SaveAs(output_path, "HWPX")
            
            if success:
                self.log(f"Conversion successful: {output_path}")
                return True
            else:
                self.log(f"Conversion failed (SaveAs returned False).", "error")
                return False

        except Exception as e:
            self.log(f"Error during conversion: {str(e)}", "error")
            return False
        finally:
            if self.hwp:
                self.hwp.Quit()
                self.hwp = None
            pythoncom.CoUninitialize()

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
