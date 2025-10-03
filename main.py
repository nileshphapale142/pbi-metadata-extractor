import zipfile
import os
import json
import csv
from datetime import datetime

# Power BI data type codes mapping
POWERBI_TYPE_CODES = {
    1: "Text",
    2: "Whole Number",
    3: "Date/Time",
    259: "Decimal Number",
    519: "Date",
    520: "Time",
    2048: "Text (Category)",
    260: "Currency",
    261: "Boolean",
    262: "Binary"
}

def get_type_name(type_code):
    """
    Convert Power BI type code to readable type name.
    
    Args:
        type_code: Numeric type code from Power BI
    
    Returns:
        String representation of the type
    """
    return POWERBI_TYPE_CODES.get(type_code, f"Type Code {type_code}")

def parse_visual_containers(layout_json, log_file, csv_file_path):
    """
    Parse visual containers from layout JSON and extract columns and measures used in visuals.
    
    Args:
        layout_json: Parsed JSON object containing layout information
        log_file: File handle to write output
        csv_file_path: Path to CSV file for visual data
    """
    try:
        sections = layout_json.get("sections", [])
        
        log_file.write(f"\n{'='*80}\n")
        log_file.write(f"VISUAL ANALYSIS - COLUMNS AND MEASURES\n")
        log_file.write(f"{'='*80}\n\n")
        
        total_visuals = 0
        
        # Prepare CSV data
        csv_data = []
        
        for section_idx, section in enumerate(sections):
            section_name = section.get("displayName", f"Section {section_idx + 1}")
            visual_containers = section.get("visualContainers", [])
            
            # Count visuals with projections first
            visuals_with_projections = []
            for visual in visual_containers:
                config_str = visual.get("config", "{}")
                try:
                    config = json.loads(config_str)
                    single_visual = config.get("singleVisual", {})
                    projections = single_visual.get("projections", {})
                    
                    # Check if projections exist and are not empty
                    has_projections = projections and any(projections.values())
                    if has_projections:
                        visuals_with_projections.append(visual)
                except json.JSONDecodeError:
                    continue
            
            # Only write section header if there are visuals with projections
            if not visuals_with_projections:
                continue
                
            log_file.write(f"\nSection: {section_name}\n")
            log_file.write(f"{'-'*80}\n")
            log_file.write(f"Total Visuals with Projections: {len(visuals_with_projections)}\n\n")
            
            for visual_idx, visual in enumerate(visuals_with_projections):
                total_visuals += 1
                visual_id = visual.get("id", "Unknown")
                
                # Parse the config JSON string
                config_str = visual.get("config", "{}")
                try:
                    config = json.loads(config_str)
                except json.JSONDecodeError:
                    continue
                
                visual_name = config.get("name", "Unnamed")
                single_visual = config.get("singleVisual", {})
                visual_type = single_visual.get("visualType", "Unknown")
                projections = single_visual.get("projections", {})
                
                log_file.write(f"Visual #{visual_idx + 1}\n")
                log_file.write(f"  ID: {visual_id}\n")
                log_file.write(f"  Name: {visual_name}\n")
                log_file.write(f"  Type: {visual_type}\n")
                
                # Extract columns and measures from projections
                log_file.write(f"  Projections:\n")
                
                # Collect projection data
                projection_details = []
                for proj_type, proj_items in projections.items():
                    if proj_items:
                        log_file.write(f"    {proj_type}:\n")
                        for item in proj_items:
                            query_ref = item.get("queryRef", "N/A")
                            active = item.get("active", True)
                            log_file.write(f"      - {query_ref} (Active: {active})\n")
                            projection_details.append({
                                "projection_type": proj_type,
                                "query_ref": query_ref,
                                "active": active
                            })
                
                # Parse dataTransforms if available for additional metadata
                data_transforms_str = visual.get("dataTransforms", "{}")
                field_details = []
                try:
                    data_transforms = json.loads(data_transforms_str)
                    selects = data_transforms.get("selects", [])
                    
                    if selects:
                        log_file.write(f"  Fields Details:\n")
                        for select in selects:
                            display_name = select.get("displayName", "N/A")
                            query_name = select.get("queryName", "N/A")
                            field_type = select.get("type", {})
                            underlying_type = field_type.get("underlyingType", "N/A")
                            format_info = select.get("format", "N/A")
                            
                            # Convert type code to type name
                            if underlying_type != "N/A":
                                type_name = get_type_name(underlying_type)
                            else:
                                type_name = "N/A"
                            
                            # Check if it's a measure (aggregation)
                            expr = select.get("expr", {})
                            aggregation_func = ""
                            is_measure = False
                            if "Aggregation" in expr:
                                agg_func = expr["Aggregation"].get("Function", "Unknown")
                                aggregation_func = f"Function {agg_func}"
                                is_measure = True
                            
                            log_file.write(f"    - Display Name: {display_name}\n")
                            log_file.write(f"      Query Name: {query_name}\n")
                            log_file.write(f"      Type: {type_name}\n")
                            if format_info != "N/A":
                                log_file.write(f"      Format: {format_info}\n")
                            if aggregation_func:
                                log_file.write(f"      Aggregation: {aggregation_func}\n")
                            
                            field_details.append({
                                "display_name": display_name,
                                "query_name": query_name,
                                "type": type_name,
                                "format": format_info,
                                "aggregation": aggregation_func,
                                "is_measure": is_measure
                            })
                            
                except json.JSONDecodeError:
                    pass
                
                # Add data to CSV
                # For each field in the visual, create a CSV row
                if field_details:
                    for field in field_details:
                        # Find matching projection
                        matching_projection = None
                        for proj in projection_details:
                            if field["query_name"] in proj["query_ref"]:
                                matching_projection = proj
                                break
                        
                        csv_row = {
                            "Section Name": section_name,
                            "Visual ID": visual_id,
                            "Visual Name": visual_name,
                            "Visual Type": visual_type,
                            "Field Display Name": field["display_name"],
                            "Field Query Name": field["query_name"],
                            "Field Type": field["type"],
                            "Field Format": field["format"] if field["format"] != "N/A" else "",
                            "Is Measure": "Yes" if field["is_measure"] else "No",
                            "Aggregation": field["aggregation"],
                            "Projection Type": matching_projection["projection_type"] if matching_projection else "",
                            "Active": matching_projection["active"] if matching_projection else ""
                        }
                        csv_data.append(csv_row)
                else:
                    # If no field details, add basic visual info
                    for proj in projection_details:
                        csv_row = {
                            "Section Name": section_name,
                            "Visual ID": visual_id,
                            "Visual Name": visual_name,
                            "Visual Type": visual_type,
                            "Field Display Name": "",
                            "Field Query Name": proj["query_ref"],
                            "Field Type": "",
                            "Field Format": "",
                            "Is Measure": "",
                            "Aggregation": "",
                            "Projection Type": proj["projection_type"],
                            "Active": proj["active"]
                        }
                        csv_data.append(csv_row)
                
                log_file.write(f"\n")
        
        log_file.write(f"\n{'='*80}\n")
        log_file.write(f"SUMMARY\n")
        log_file.write(f"{'='*80}\n")
        log_file.write(f"Total Sections: {len(sections)}\n")
        log_file.write(f"Total Visuals with Projections: {total_visuals}\n")
        log_file.write(f"{'='*80}\n")
        
        # Write CSV file
        if csv_data:
            with open(csv_file_path, 'w', newline='', encoding='utf-8') as csv_file:
                fieldnames = [
                    "Section Name", "Visual ID", "Visual Name", "Visual Type",
                    "Field Display Name", "Field Query Name", "Field Type", "Field Format",
                    "Is Measure", "Aggregation", "Projection Type", "Active"
                ]
                writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(csv_data)
            
            print(f"CSV file created: {csv_file_path}")
            log_file.write(f"\nCSV file created: {csv_file_path}\n")
        
    except Exception as e:
        log_file.write(f"\nError parsing visual containers: {str(e)}\n")


def extract_pbix_contents(pbix_file_path, output_log_path="output.log", output_csv_path="visuals_data.csv"):
    """
    Reads a PBIX file as a ZIP archive and lists all files with their paths.
    Writes the output to a log file.
    
    Args:
        pbix_file_path: Path to the PBIX file
        output_log_path: Path to the output log file (default: output.log)
        output_csv_path: Path to the output CSV file (default: visuals_data.csv)
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
                layout_json = None
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
                
                # Parse visual containers if layout JSON was successfully loaded
                if layout_json:
                    parse_visual_containers(layout_json, log_file, output_csv_path)
        
        print(f"Analysis complete! Output written to: {output_log_path}")
        
    except zipfile.BadZipFile:
        print(f"Error: {pbix_file_path} is not a valid ZIP/PBIX file")
    except Exception as e:
        print(f"Error occurred: {str(e)}")

# Example usage
if __name__ == "__main__":
    # Specify your PBIX file path
    pbix_path = "./data/Components requirements next 12 weeks by ORDERING plant.pbix"
    output_log = "output.log"
    output_csv = "visuals_data.csv"
    
    extract_pbix_contents(pbix_path, output_log, output_csv)