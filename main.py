import zipfile
import os
import json
from datetime import datetime

def extract_pbix_contents(pbix_file_path, output_log_path="output.log"):
    """
    Reads a PBIX file as a ZIP archive and lists all files with their paths.
    Writes the output to a log file.
    
    Args:
        pbix_file_path: Path to the PBIX file
        output_log_path: Path to the output log file (default: output.log)
    """
    try:
        # Check if PBIX file exists
        if not os.path.exists(pbix_file_path):
            print(f"Error: PBIX file not found at {pbix_file_path}")
            return
        
        # Open log file for writing
        with open(output_log_path, 'w', encoding='utf-8') as log_file:
            log_file.write(f"PBIX File Analysis Report\n")
            log_file.write(f"{'='*80}\n")
            log_file.write(f"File: {pbix_file_path}\n")
            log_file.write(f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            log_file.write(f"{'='*80}\n\n")
            
            # Open PBIX file as ZIP
            with zipfile.ZipFile(pbix_file_path, 'r') as zip_ref:
                # Get list of all files in the archive
                file_list = zip_ref.namelist()
                
                log_file.write(f"Total files found: {len(file_list)}\n\n")
                log_file.write(f"Files and Folders:\n")
                log_file.write(f"{'-'*80}\n")
                
                # Separate folders and files
                folders = set()
                files = []
                layout_file_path = None
                
                for file_path in file_list:
                    # Check if it's a directory (ends with /)
                    if file_path.endswith('/'):
                        folders.add(file_path)
                        log_file.write(f"[FOLDER] {file_path}\n")
                    else:
                        files.append(file_path)
                        # Extract folder path from file path
                        folder_path = os.path.dirname(file_path)
                        if folder_path:
                            folders.add(folder_path + '/')
                        
                        # Get file info
                        file_info = zip_ref.getinfo(file_path)
                        file_size = file_info.file_size
                        
                        log_file.write(f"[FILE]   {file_path} (Size: {file_size:,} bytes)\n")
                        
                        # Check if this is the Layout file
                        if file_path == "Report/Layout":
                            layout_file_path = file_path
                
                # Summary
                log_file.write(f"\n{'='*80}\n")
                log_file.write(f"Summary:\n")
                log_file.write(f"  Total Folders: {len(folders)}\n")
                log_file.write(f"  Total Files: {len(files)}\n")
                log_file.write(f"{'='*80}\n\n")
                
                # Extract and write Layout file contents
                if layout_file_path:
                    log_file.write(f"\n{'='*80}\n")
                    log_file.write(f"LAYOUT FILE CONTENTS (Report/Layout)\n")
                    log_file.write(f"{'='*80}\n\n")
                    
                    try:
                        # Read the Layout file content
                        with zip_ref.open(layout_file_path) as layout_file:
                            layout_bytes = layout_file.read()
                            
                            # Try different encodings
                            layout_content = None
                            encodings_to_try = ['utf-16-le', 'utf-16', 'utf-8']
                            
                            for encoding in encodings_to_try:
                                try:
                                    layout_content = layout_bytes.decode(encoding)
                                    # If successful, break
                                    break
                                except (UnicodeDecodeError, UnicodeError):
                                    continue
                            
                            if layout_content is None:
                                # Fallback to utf-8 with error handling
                                layout_content = layout_bytes.decode('utf-8', errors='replace')
                            
                            # Try to parse as JSON and pretty print
                            try:
                                layout_json = json.loads(layout_content)
                                formatted_json = json.dumps(layout_json, indent=2, ensure_ascii=False)
                                log_file.write(formatted_json)
                            except json.JSONDecodeError:
                                # If not valid JSON, write as-is
                                log_file.write(layout_content)
                        
                        log_file.write(f"\n\n{'='*80}\n")
                        log_file.write(f"END OF LAYOUT FILE\n")
                        log_file.write(f"{'='*80}\n")
                        
                    except Exception as e:
                        log_file.write(f"Error reading Layout file: {str(e)}\n")
                else:
                    log_file.write(f"\nNote: Layout file not found in PBIX archive.\n")
        
        print(f"Analysis complete! Output written to: {output_log_path}")
        
    except zipfile.BadZipFile:
        print(f"Error: {pbix_file_path} is not a valid ZIP/PBIX file")
    except Exception as e:
        print(f"Error occurred: {str(e)}")

# Example usage
if __name__ == "__main__":
    # Specify your PBIX file path
    pbix_path = "./data/Components requirements next 12 weeks by ORDERING plant.pbix"  # Change this to your PBIX file path
    output_path = "output.log"
    
    extract_pbix_contents(pbix_path, output_path)