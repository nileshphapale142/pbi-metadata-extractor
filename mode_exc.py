import zipfile
import os
import json
from datetime import datetime


def extract_pbix_file_list(pbix_file_path, output_log_path="mode_exc_output.log"):
    """
    Reads a PBIX file as a ZIP archive and lists all files with their paths.
    Uses pbi-tools/PBIXRay to extract DataModel contents.
    Writes the output to a log file.

    Args:
        pbix_file_path: Path to the PBIX file or file-like object
        output_log_path: Path to the output log file (default: mode_exc_output.log)
    
    Returns:
        dict: Dictionary containing file information
    """
    try:
        # Check if PBIX file exists
        if isinstance(pbix_file_path, str) and not os.path.exists(pbix_file_path):
            print(f"Error: PBIX file not found at {pbix_file_path}")
            return None

        result = {
            'total_files': 0,
            'total_folders': 0,
            'files': [],
            'folders': set(),
            'file_details': []
        }

        # Open log file for writing
        with open(output_log_path, "w", encoding="utf-8") as log_file:
            log_file.write(f"PBIX File Analysis Report\n")
            log_file.write(f"{'=' * 80}\n")
            log_file.write(f"File: {pbix_file_path}\n")
            log_file.write(
                f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            )
            log_file.write(f"{'=' * 80}\n\n")

            # Open PBIX file as ZIP
            with zipfile.ZipFile(pbix_file_path, "r") as zip_ref:
                # Get list of all files in the archive
                file_list = zip_ref.namelist()

                log_file.write(f"Total files found: {len(file_list)}\n\n")
                log_file.write(f"Files and Folders:\n")
                log_file.write(f"{'-' * 80}\n")

                # Separate folders and files
                folders = set()
                files = []
                datamodel_path = None

                for file_path in file_list:
                    # Check if it's a directory (ends with /)
                    if file_path.endswith("/"):
                        folders.add(file_path)
                        log_file.write(f"[FOLDER] {file_path}\n")
                    else:
                        files.append(file_path)
                        # Extract folder path from file path
                        folder_path = os.path.dirname(file_path)
                        if folder_path:
                            folders.add(folder_path + "/")

                        # Get file info
                        file_info = zip_ref.getinfo(file_path)
                        file_size = file_info.file_size

                        log_file.write(
                            f"[FILE]   {file_path} (Size: {file_size:,} bytes)\n"
                        )

                        # Store file details
                        result['file_details'].append({
                            'path': file_path,
                            'size': file_size,
                            'compressed_size': file_info.compress_size,
                            'folder': folder_path
                        })

                        # Check if this is the DataModel file
                        if file_path == "DataModel":
                            datamodel_path = file_path

                # Update result dictionary
                result['total_files'] = len(files)
                result['total_folders'] = len(folders)
                result['files'] = files
                result['folders'] = folders

                # Summary
                log_file.write(f"\n{'=' * 80}\n")
                log_file.write(f"Summary:\n")
                log_file.write(f"  Total Folders: {len(folders)}\n")
                log_file.write(f"  Total Files: {len(files)}\n")
                log_file.write(f"{'=' * 80}\n")

            # Extract DataModel using PBIXRay
            if datamodel_path or os.path.exists(pbix_file_path):
                extract_datamodel_with_pbixray(pbix_file_path, log_file)

        print(f"File list extraction complete! Output written to: {output_log_path}")
        return result

    except zipfile.BadZipFile:
        print(f"Error: {pbix_file_path} is not a valid ZIP/PBIX file")
        return None
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        return None


def extract_datamodel_with_pbixray(pbix_file_path, log_file):
    """
    Extract DataModel contents using PBIXRay library.

    Args:
        pbix_file_path: Path to the PBIX file
        log_file: File handle to write output
    """
    log_file.write(f"\n\n{'=' * 80}\n")
    log_file.write(f"DATAMODEL EXTRACTION USING PBIXRAY\n")
    log_file.write(f"{'=' * 80}\n\n")

    try:
        # Try to import pbi-tools (PBIXRay)
        try:
            from pbixray import PBIXRay
            
            log_file.write("✓ PBIXRay library loaded successfully\n\n")
            
            # Load the PBIX file
            log_file.write(f"Loading PBIX file: {pbix_file_path}\n")
            model = PBIXRay(pbix_file_path)
            
            log_file.write("✓ PBIX file loaded successfully\n\n")
            
            # Extract tables
            log_file.write(f"{'=' * 80}\n")
            log_file.write(f"TABLES\n")
            log_file.write(f"{'=' * 80}\n\n")
            
            tables = model.tables
            log_file.write(f"Tables found: {model}\n")
            log_file.write(f"Total Tables: {len(tables)}\n\n")

            log_file.write(f"{tables}\n")
            for idx, table in enumerate(tables, 1):
                log_file.write(f"Table {idx}: {table.name}\n")
                log_file.write(f"  Is Hidden: {table.is_hidden}\n")
                
                # Columns
                if hasattr(table, 'columns') and table.columns:
                    log_file.write(f"  Columns ({len(table.columns)}):\n")
                    for col in table.columns:
                        col_name = col.name if hasattr(col, 'name') else str(col)
                        col_type = col.data_type if hasattr(col, 'data_type') else 'Unknown'
                        log_file.write(f"    - {col_name} ({col_type})\n")
                
                # Measures
                if hasattr(table, 'measures') and table.measures:
                    log_file.write(f"  Measures ({len(table.measures)}):\n")
                    for measure in table.measures:
                        measure_name = measure.name if hasattr(measure, 'name') else str(measure)
                        log_file.write(f"    - {measure_name}\n")
                        if hasattr(measure, 'expression'):
                            expr = str(measure.expression)
                            if len(expr) < 200:
                                log_file.write(f"      Expression: {expr}\n")
                            else:
                                log_file.write(f"      Expression: {expr[:200]}...\n")
                
                log_file.write("\n")
            
            # Extract measures
            log_file.write(f"\n{'=' * 80}\n")
            log_file.write(f"ALL MEASURES\n")
            log_file.write(f"{'=' * 80}\n\n")
            
            if hasattr(model, 'dax_measures'):
                measures = model.dax_measures
                log_file.write(f"Total Measures: {len(measures)}\n\n")
                
                for idx, measure in enumerate(measures, 1):
                    measure_name = measure.name if hasattr(measure, 'name') else f"Measure {idx}"
                    table_name = measure.table if hasattr(measure, 'table') else 'Unknown'
                    
                    log_file.write(f"Measure {idx}: {table_name}.{measure_name}\n")
                    
                    if hasattr(measure, 'expression'):
                        log_file.write(f"  Expression:\n")
                        log_file.write(f"    {measure.expression}\n")
                    
                    if hasattr(measure, 'format_string'):
                        log_file.write(f"  Format: {measure.format_string}\n")
                    
                    log_file.write("\n")
            
            # Extract relationships
            log_file.write(f"\n{'=' * 80}\n")
            log_file.write(f"RELATIONSHIPS\n")
            log_file.write(f"{'=' * 80}\n\n")
            
            if hasattr(model, 'relationships'):
                relationships = model.relationships
                log_file.write(f"Total Relationships: {len(relationships)}\n\n")
                
                for idx, rel in enumerate(relationships, 1):
                    from_table = rel.from_table if hasattr(rel, 'from_table') else 'Unknown'
                    from_column = rel.from_column if hasattr(rel, 'from_column') else 'Unknown'
                    to_table = rel.to_table if hasattr(rel, 'to_table') else 'Unknown'
                    to_column = rel.to_column if hasattr(rel, 'to_column') else 'Unknown'
                    
                    log_file.write(f"Relationship {idx}:\n")
                    log_file.write(f"  From: {from_table}.{from_column}\n")
                    log_file.write(f"  To: {to_table}.{to_column}\n")
                    
                    if hasattr(rel, 'cardinality'):
                        log_file.write(f"  Cardinality: {rel.cardinality}\n")
                    
                    if hasattr(rel, 'cross_filter_direction'):
                        log_file.write(f"  Cross Filter Direction: {rel.cross_filter_direction}\n")
                    
                    log_file.write("\n")
            
            # Extract data sources
            log_file.write(f"\n{'=' * 80}\n")
            log_file.write(f"DATA SOURCES\n")
            log_file.write(f"{'=' * 80}\n\n")
            
            if hasattr(model, 'data_sources'):
                data_sources = model.data_sources
                log_file.write(f"Total Data Sources: {len(data_sources)}\n\n")
                
                for idx, ds in enumerate(data_sources, 1):
                    ds_name = ds.name if hasattr(ds, 'name') else f"DataSource {idx}"
                    ds_type = ds.type if hasattr(ds, 'type') else 'Unknown'
                    
                    log_file.write(f"Data Source {idx}: {ds_name}\n")
                    log_file.write(f"  Type: {ds_type}\n")
                    
                    if hasattr(ds, 'connection_string'):
                        conn_str = str(ds.connection_string)
                        if len(conn_str) < 200:
                            log_file.write(f"  Connection String: {conn_str}\n")
                        else:
                            log_file.write(f"  Connection String: {conn_str[:200]}...\n")
                    
                    log_file.write("\n")
            
            # Write full model as JSON if available
            log_file.write(f"\n{'=' * 80}\n")
            log_file.write(f"RAW MODEL DATA (JSON)\n")
            log_file.write(f"{'=' * 80}\n\n")
            
            try:
                if hasattr(model, 'to_dict'):
                    model_dict = model.to_dict()
                    formatted_json = json.dumps(model_dict, indent=2, ensure_ascii=False)
                    log_file.write(formatted_json[:50000])  # Limit to 50KB
                    if len(formatted_json) > 50000:
                        log_file.write("\n\n... (truncated, showing first 50KB)")
                elif hasattr(model, '_model'):
                    # Try to access internal model structure
                    model_str = str(model._model)
                    log_file.write(model_str[:10000])
                    if len(model_str) > 10000:
                        log_file.write("\n\n... (truncated)")
            except Exception as e:
                log_file.write(f"Could not extract raw model data: {e}\n")
            
            log_file.write(f"\n\n{'=' * 80}\n")
            log_file.write(f"END OF DATAMODEL EXTRACTION\n")
            log_file.write(f"{'=' * 80}\n")
            
        except ImportError as e:
            log_file.write(f"✗ PBIXRay library not found\n")
            log_file.write(f"Error: {e}\n\n")
            log_file.write(f"To install PBIXRay, run:\n")
            log_file.write(f"  uv pip install pbi-tools\n")
            log_file.write(f"  or\n")
            log_file.write(f"  pip install pbi-tools\n\n")
            
            # Provide alternative installation commands
            log_file.write(f"Alternative libraries to try:\n")
            log_file.write(f"  uv pip install pbixray\n")
            log_file.write(f"  uv pip install powerbi-tools\n")

    except Exception as e:
        log_file.write(f"Error extracting DataModel with PBIXRay: {str(e)}\n")
        import traceback
        log_file.write(f"Traceback:\n{traceback.format_exc()}\n")


def get_file_structure(pbix_file_path):
    """
    Get PBIX file structure as a hierarchical dictionary.

    Args:
        pbix_file_path: Path to the PBIX file or file-like object
    
    Returns:
        dict: Hierarchical structure of folders and files
    """
    try:
        structure = {}

        with zipfile.ZipFile(pbix_file_path, "r") as zip_ref:
            file_list = zip_ref.namelist()

            for file_path in file_list:
                if file_path.endswith("/"):
                    continue

                parts = file_path.split("/")
                current = structure

                for i, part in enumerate(parts):
                    if i == len(parts) - 1:
                        # It's a file
                        if "_files" not in current:
                            current["_files"] = []
                        
                        file_info = zip_ref.getinfo(file_path)
                        current["_files"].append({
                            'name': part,
                            'size': file_info.file_size,
                            'compressed_size': file_info.compress_size
                        })
                    else:
                        # It's a folder
                        if part not in current:
                            current[part] = {}
                        current = current[part]

        return structure

    except Exception as e:
        print(f"Error getting file structure: {e}")
        return None


# Example usage
if __name__ == "__main__":
    # Specify your PBIX file path
    pbix_path = "./data/sem_model.pbix"
    
    print("Installing PBIXRay library...")
    print("Run: uv pip install pbi-tools")
    print("or: pip install pbi-tools")
    print()
    
    # Extract file list
    result = extract_pbix_file_list(pbix_path, "mode_exc_output.log")
    
    if result:
        print(f"\nSummary:")
        print(f"Total Files: {result['total_files']}")
        print(f"Total Folders: {result['total_folders']}")
        
        # Get hierarchical structure
        structure = get_file_structure(pbix_path)
        if structure:
            print(f"\nFile Structure:")
            print(f"Root folders: {list(structure.keys())}")