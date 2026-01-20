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

    def initialize_hwp(self):
        if self.hwp is None:
            try:
                self.log("Initializing Hancom Office instance...")
                self.hwp = Hwp(visible=False)
                return True
            except Exception as e:
                self.log(f"Failed to initialize HWP: {str(e)}", "error")
                return False
        return True

    def quit_hwp(self):
        if self.hwp:
            try:
                self.hwp.quit()
            except Exception:
                pass
            self.hwp = None

    def _convert_file_internal(self, input_path):
        """Internal method to convert a single file using the existing self.hwp instance."""
        input_path = os.path.abspath(input_path)
        file_name, ext = os.path.splitext(input_path)
        output_path = f"{file_name}.hwpx"
        
        try:
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
                        self.log(f"Warning: save_as returned True but file is not a valid ZIP (HWPX). Retrying...", "warning")
                else:
                     self.log(f"save_as({fmt}) returned False.")
            
            # Close the current document to free memory, but keep HWP instance open
            self.hwp.clear() 
            return success

        except Exception as e:
            self.log(f"Error during conversion of {input_path}: {str(e)}", "error")
            return False

    def convert_to_hwpx(self, input_path):
        """Convert a single file. Manages HWP lifecycle (Init -> Convert -> Quit)."""
        if not os.path.exists(input_path):
            self.log(f"File not found: {input_path}", "error")
            return False
            
        try:
            if self.initialize_hwp():
                result = self._convert_file_internal(input_path)
                return result
            return False
        finally:
            self.quit_hwp()

    def is_zip_file(self, filepath):
        try:
            with open(filepath, 'rb') as f:
                header = f.read(4)
            return header == b'\x50\x4b\x03\x04'
        except Exception:
            return False

    def convert_directory(self, input_dir, progress_callback=None):
        """Convert a directory of files. Manages HWP lifecycle (Init -> Loop -> Quit)."""
        input_dir = os.path.abspath(input_dir)
        if not os.path.exists(input_dir):
            self.log(f"Directory not found: {input_dir}", "error")
            return

        self.log(f"Scanning directory: {input_dir}")
        
        # 1. Count total files first
        total_files = 0
        hwp_files = []
        for root, dirs, files in os.walk(input_dir):
            for file in files:
                if file.lower().endswith('.hwp'):
                    hwp_files.append(os.path.join(root, file))
        
        total_files = len(hwp_files)
        if total_files == 0:
            self.log("No HWP files found.")
            return

        self.log(f"Found {total_files} HWP files.")

        # Initialize Once
        if not self.initialize_hwp():
            return

        count = 0
        try:
            for idx, full_path in enumerate(hwp_files, 1):
                # Use internal method that reuses self.hwp
                if self._convert_file_internal(full_path):
                    count += 1
                
                # Update progress
                if progress_callback:
                    progress_callback(idx, total_files)
                    
        finally:
            # Quit Once
            self.quit_hwp()
                        
        self.log(f"Batch conversion complete. Converted {count} files.")
