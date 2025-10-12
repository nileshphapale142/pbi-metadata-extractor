import zipfile
import os
import json
import csv
import re
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
    262: "Binary",
}

def clean_text(text):
    """
    Remove invisible Unicode characters like Left-to-Right Mark, Right-to-Left Mark, etc.

    Args:
        text: String to clean

    Returns:
        Cleaned string
    """
    if not isinstance(text, str):
        return text

    # Remove common invisible Unicode characters
    invisible_chars = [
        "\u200e",  # Left-to-Right Mark (LRM)
        "\u200f",  # Right-to-Left Mark (RLM)
        "\u202a",  # Left-to-Right Embedding
        "\u202b",  # Right-to-Left Embedding
        "\u202c",  # Pop Directional Formatting
        "\u202d",  # Left-to-Right Override
        "\u202e",  # Right-to-Left Override
        "\ufeff",  # Zero Width No-Break Space (BOM)
        "\u200b",  # Zero Width Space
        "\u200c",  # Zero Width Non-Joiner
        "\u200d",  # Zero Width Joiner
    ]

    cleaned = text
    for char in invisible_chars:
        cleaned = cleaned.replace(char, "")

    return cleaned


def get_type_name(type_code):
    """
    Convert Power BI type code to readable type name.

    Args:
        type_code: Numeric type code from Power BI

    Returns:
        String representation of the type
    """
    return POWERBI_TYPE_CODES.get(type_code, f"Type Code {type_code}")


def parse_filter_item(filter_item):
    """
    Parse individual filter item to extract meaningful information.

    Args:
        filter_item: Filter dictionary

    Returns:
        Dictionary with filter details or None if no meaningful filter
    """
    try:
        # Extract filter conditions - only process if Where clause exists
        filter_data = filter_item.get("filter", {})
        where_clause = filter_data.get("Where", [])

        # Skip filters without Where clause
        if not where_clause:
            return None

        # Extract basic filter info
        filter_type = filter_item.get("type", "Unknown")
        expression = filter_item.get("expression", {})

        # Extract table and column from expression
        table_name = "Unknown"
        column_name = "Unknown"

        if "Column" in expression:
            column_info = expression["Column"]

            # Get table name from Entity
            entity_expr = column_info.get("Expression", {})
            if "SourceRef" in entity_expr:
                source_ref = entity_expr["SourceRef"]
                table_name = source_ref.get("Entity", "Unknown")

            # Get column name from Property
            column_name = column_info.get("Property", "Unknown")

        # Helper function to clean Power BI literal values
        def clean_literal_value(value_str):
            """Remove Power BI type suffixes from literal values"""
            if not isinstance(value_str, str):
                return str(value_str)

            # Remove common Power BI type suffixes
            # D = Decimal, L = Long, M = Money/Currency
            if (
                value_str.endswith("D")
                or value_str.endswith("L")
                or value_str.endswith("M")
            ):
                # Check if it's actually a number with suffix
                base_value = value_str[:-1]
                # Check if base is numeric (including decimals and negatives)
                if (
                    base_value.replace("-", "")
                    .replace(".", "")
                    .replace(",", "")
                    .isdigit()
                ):
                    return base_value

            return value_str

        # Extract filter conditions
        conditions = []
        operator = "Unknown"

        for condition in where_clause:
            if "Condition" in condition:
                cond = condition["Condition"]

                # Handle "In" operator
                if "In" in cond:
                    in_clause = cond["In"]
                    values = in_clause.get("Values", [])

                    if values:
                        operator = "IN"
                        # Extract literal values
                        extracted_values = []
                        for value_array in values:
                            if isinstance(value_array, list):
                                for val in value_array:
                                    if isinstance(val, dict) and "Literal" in val:
                                        literal_val = val["Literal"].get("Value", "")
                                        cleaned_val = clean_literal_value(literal_val)
                                        extracted_values.append(cleaned_val)

                        if extracted_values:
                            # Show all values without truncation
                            conditions.append(f"IN ({', '.join(extracted_values)})")
                        else:
                            conditions.append(f"IN ({len(values)} values)")

                # Handle "Not" operator with "In"
                elif "Not" in cond:
                    not_clause = cond["Not"]
                    if "Expression" in not_clause and "In" in not_clause["Expression"]:
                        in_clause = not_clause["Expression"]["In"]
                        values = in_clause.get("Values", [])
                        operator = "NOT IN"

                        if values:
                            extracted_values = []
                            for value_array in values:
                                if isinstance(value_array, list):
                                    for val in value_array:
                                        if isinstance(val, dict) and "Literal" in val:
                                            literal_val = val["Literal"].get(
                                                "Value", ""
                                            )
                                            cleaned_val = clean_literal_value(
                                                literal_val
                                            )
                                            extracted_values.append(cleaned_val)

                            if extracted_values:
                                # Show all values without truncation
                                conditions.append(
                                    f"NOT IN ({', '.join(extracted_values)})"
                                )
                            else:
                                conditions.append(f"NOT IN ({len(values)} values)")

                # Handle comparison operators
                elif "Comparison" in cond:
                    comparison = cond["Comparison"]
                    comparison_kind = comparison.get("ComparisonKind", 0)

                    # Map comparison kind to operator
                    comparison_map = {0: "=", 1: "<>", 2: ">", 3: ">=", 4: "<", 5: "<="}
                    operator = comparison_map.get(comparison_kind, "Unknown")

                    # Try to get the comparison value
                    right_expr = comparison.get("Right", {})
                    if "Literal" in right_expr:
                        value = right_expr["Literal"].get("Value", "")
                        cleaned_value = clean_literal_value(value)
                        conditions.append(f"{operator} {cleaned_value}")
                    else:
                        conditions.append(operator)

                # Handle "Between" operator
                elif "Between" in cond:
                    operator = "BETWEEN"
                    between_clause = cond["Between"]
                    lower = between_clause.get("Lower", {})
                    upper = between_clause.get("Upper", {})

                    lower_val = ""
                    upper_val = ""

                    if "Literal" in lower:
                        lower_val = clean_literal_value(
                            lower["Literal"].get("Value", "")
                        )
                    if "Literal" in upper:
                        upper_val = clean_literal_value(
                            upper["Literal"].get("Value", "")
                        )

                    if lower_val and upper_val:
                        conditions.append(f"BETWEEN {lower_val} AND {upper_val}")
                    else:
                        conditions.append("BETWEEN (values)")

                # Handle "Contains" operator
                elif "Contains" in cond:
                    operator = "CONTAINS"
                    contains_clause = cond["Contains"]
                    right_expr = contains_clause.get("Right", {})

                    if "Literal" in right_expr:
                        value = right_expr["Literal"].get("Value", "")
                        cleaned_value = clean_literal_value(value)
                        conditions.append(f"CONTAINS '{cleaned_value}'")
                    else:
                        conditions.append("CONTAINS (text)")

        # Only return filter if we have meaningful conditions
        if not conditions:
            return None

        return {
            "table": clean_text(table_name),
            "column": clean_text(column_name),
            "type": filter_type,
            "operator": operator,
            "conditions": conditions,
        }

    except Exception as e:
        return None


def extract_filters(filters_json):
    """
    Extract and format filter information from filters JSON.

    Args:
        filters_json: Filters JSON string or object

    Returns:
        List of filter dictionaries with detailed information
    """
    filter_list = []

    try:
        # Parse JSON if string
        if isinstance(filters_json, str):
            if not filters_json or filters_json.strip() == "":
                return []
            filters = json.loads(filters_json)
        else:
            filters = filters_json

        if not filters:
            return []

        # Handle different filter structures
        if isinstance(filters, list):
            for filter_item in filters:
                filter_info = parse_filter_item(filter_item)
                if filter_info:
                    filter_list.append(filter_info)
        elif isinstance(filters, dict):
            filter_info = parse_filter_item(filters)
            if filter_info:
                filter_list.append(filter_info)

    except (json.JSONDecodeError, Exception) as e:
        pass

    return filter_list


def format_filters_for_display(filters):
    """
    Format filter list for display in UI.

    Args:
        filters: List of filter dictionaries

    Returns:
        Formatted string for display
    """
    if not filters:
        return ""

    formatted_filters = []
    for f in filters:
        table = f.get("table", "Unknown")
        column = f.get("column", "Unknown")
        conditions = f.get("conditions", [])

        if conditions:
            condition_str = " ".join(conditions)
            formatted_filters.append(f"{table}.{column}: {condition_str}")
        else:
            formatted_filters.append(f"{table}.{column}")

    return " | ".join(formatted_filters)


def extract_report_metadata(pbix_file_path):
    """
    Extract comprehensive report metadata including pages, visuals, and filters.

    Args:
        pbix_file_path: Path to the PBIX file or file-like object

    Returns:
        Dictionary containing report metadata
    """
    try:
        report_data = {"summary": {}, "pages": [], "visuals": []}

        # Open PBIX file as ZIP
        with zipfile.ZipFile(pbix_file_path, "r") as zip_ref:
            # Find Layout file
            layout_file_path = None
            for file_path in zip_ref.namelist():
                if file_path == "Report/Layout":
                    layout_file_path = file_path
                    break

            if not layout_file_path:
                return report_data

            # Read and parse Layout file
            with zip_ref.open(layout_file_path) as layout_file:
                layout_bytes = layout_file.read()

                # Try different encodings
                layout_content = None
                encodings_to_try = ["utf-16-le", "utf-16", "utf-8"]

                for encoding in encodings_to_try:
                    try:
                        layout_content = layout_bytes.decode(encoding)
                        break
                    except (UnicodeDecodeError, UnicodeError):
                        continue

                if layout_content is None:
                    layout_content = layout_bytes.decode("utf-8", errors="replace")

                # Parse JSON
                layout_json = json.loads(layout_content)
                sections = layout_json.get("sections", [])

                total_visuals = 0

                # Process each page/section
                for section_idx, section in enumerate(sections):
                    page_name = clean_text(
                        section.get("displayName", f"Page {section_idx + 1}")
                    )
                    visual_containers = section.get("visualContainers", [])

                    # Extract page-level filters (parse JSON string)
                    page_filters_str = section.get("filters", "[]")
                    page_filters = extract_filters(page_filters_str)
                    page_filters_display = format_filters_for_display(page_filters)

                    page_visuals_count = 0
                    page_visuals_with_data = 0

                    # Process each visual
                    for visual_idx, visual in enumerate(visual_containers):
                        config_str = visual.get("config", "{}")
                        try:
                            config = json.loads(config_str)
                            single_visual = config.get("singleVisual", {})
                            projections = single_visual.get("projections", {})

                            # Check if projections exist and are not empty
                            has_projections = projections and any(projections.values())

                            if not has_projections:
                                continue

                            page_visuals_with_data += 1
                            total_visuals += 1

                            visual_id = str(visual.get("id", f"visual_{visual_idx}"))
                            visual_type = clean_text(
                                single_visual.get("visualType", "Unknown")
                            )

                            # Extract visual title from vcObjects
                            visual_title = "[No Title]"
                            vc_objects = single_visual.get("vcObjects", {})
                            if vc_objects and "title" in vc_objects:
                                title_settings = vc_objects["title"]
                                if (
                                    isinstance(title_settings, list)
                                    and len(title_settings) > 0
                                ):
                                    for title_obj in title_settings:
                                        properties = title_obj.get("properties", {})
                                        if "text" in properties:
                                            text_expr = properties["text"].get(
                                                "expr", {}
                                            )
                                            if "Literal" in text_expr:
                                                visual_title = clean_text(
                                                    text_expr["Literal"]
                                                    .get("Value", "[No Title]")
                                                    .strip("'")
                                                )
                                                break

                            # Extract visual-level filters (parse JSON string)
                            visual_filters_str = visual.get("filters", "[]")
                            visual_filters = extract_filters(visual_filters_str)
                            visual_filters_display = format_filters_for_display(
                                visual_filters
                            )

                            # Collect projection data
                            projection_details = []
                            for proj_type, proj_items in projections.items():
                                if proj_items:
                                    for item in proj_items:
                                        query_ref = clean_text(
                                            item.get("queryRef", "N/A")
                                        )
                                        active = item.get("active", True)
                                        projection_details.append(
                                            {
                                                "projection_type": proj_type,
                                                "query_ref": query_ref,
                                                "active": active,
                                            }
                                        )

                            # Parse dataTransforms
                            data_transforms_str = visual.get("dataTransforms", "{}")
                            field_details = []
                            try:
                                data_transforms = json.loads(data_transforms_str)
                                selects = data_transforms.get("selects", [])

                                for select in selects:
                                    display_name = clean_text(
                                        select.get("displayName", "N/A")
                                    )
                                    query_name = clean_text(
                                        select.get("queryName", "N/A")
                                    )
                                    field_type = select.get("type", {})
                                    underlying_type = field_type.get(
                                        "underlyingType", "N/A"
                                    )
                                    format_info = select.get("format", "N/A")

                                    # Convert type code to type name
                                    if underlying_type != "N/A":
                                        type_name = get_type_name(underlying_type)
                                    else:
                                        type_name = "N/A"

                                    # Check if it's a measure
                                    expr = select.get("expr", {})
                                    aggregation_func = ""
                                    is_measure = False
                                    if "Aggregation" in expr:
                                        agg_func = expr["Aggregation"].get(
                                            "Function", "Unknown"
                                        )
                                        aggregation_func = f"Function {agg_func}"
                                        is_measure = True

                                    field_details.append(
                                        {
                                            "display_name": display_name,
                                            "query_name": query_name,
                                            "type": type_name,
                                            "format": format_info,
                                            "aggregation": aggregation_func,
                                            "is_measure": is_measure,
                                        }
                                    )

                            except json.JSONDecodeError:
                                pass

                            # Create visual records
                            if field_details:
                                for field in field_details:
                                    # Find matching projection
                                    matching_projection = None
                                    for proj in projection_details:
                                        if field["query_name"] in proj["query_ref"]:
                                            matching_projection = proj
                                            break

                                    visual_row = {
                                        "Page Name": page_name,
                                        "Visual ID": visual_id,
                                        "Visual Title": visual_title,
                                        "Visual Type": visual_type,
                                        "Visual Filters": visual_filters_display,
                                        "Field Display Name": field["display_name"],
                                        "Field Query Name": field["query_name"],
                                        "Field Type": field["type"],
                                        "Field Format": field["format"]
                                        if field["format"] != "N/A"
                                        else "",
                                        "Is Measure": "Yes"
                                        if field["is_measure"]
                                        else "No",
                                        "Aggregation": field["aggregation"],
                                        "Projection Type": matching_projection[
                                            "projection_type"
                                        ]
                                        if matching_projection
                                        else "",
                                        "Active": str(matching_projection["active"])
                                        if matching_projection
                                        else "",
                                    }
                                    report_data["visuals"].append(visual_row)
                            else:
                                # If no field details, add basic visual info
                                for proj in projection_details:
                                    visual_row = {
                                        "Page Name": page_name,
                                        "Visual ID": visual_id,
                                        "Visual Title": visual_title,
                                        "Visual Type": visual_type,
                                        "Visual Filters": visual_filters_display,
                                        "Field Display Name": "",
                                        "Field Query Name": proj["query_ref"],
                                        "Field Type": "",
                                        "Field Format": "",
                                        "Is Measure": "",
                                        "Aggregation": "",
                                        "Projection Type": proj["projection_type"],
                                        "Active": str(proj["active"]),
                                    }
                                    report_data["visuals"].append(visual_row)

                        except json.JSONDecodeError:
                            continue

                    # Add page summary
                    page_data = {
                        "Page Name": page_name,
                        "Visual Count": page_visuals_with_data,
                        "Page Filters": page_filters_display
                        if page_filters_display
                        else "None",
                    }
                    report_data["pages"].append(page_data)

                # Add report summary
                report_data["summary"] = {
                    "Total Pages": len(sections),
                    "Total Visuals": total_visuals,
                }

        return report_data

    except Exception as e:
        raise Exception(f"Error extracting report metadata: {str(e)}")


def extract_visuals_data(pbix_file_path):
    """
    Extract visual data from PBIX file and return as a list of dictionaries.
    This function is designed for use with Streamlit.

    Args:
        pbix_file_path: Path to the PBIX file or file-like object

    Returns:
        Dictionary containing report metadata
    """
    metadata = extract_report_metadata(pbix_file_path)
    return metadata


# ... (keep all existing functions for backward compatibility)
# ... (parse_visual_containers and extract_pbix_contents remain the same)


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

        log_file.write(f"\n{'=' * 80}\n")
        log_file.write(f"VISUAL ANALYSIS - COLUMNS AND MEASURES\n")
        log_file.write(f"{'=' * 80}\n\n")

        total_visuals = 0

        # Prepare CSV data
        csv_data = []

        for section_idx, section in enumerate(sections):
            page_name = clean_text(
                section.get("displayName", f"Page {section_idx + 1}")
            )
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

            log_file.write(f"\nPage: {page_name}\n")
            log_file.write(f"{'-' * 80}\n")
            log_file.write(
                f"Total Visuals with Projections: {len(visuals_with_projections)}\n\n"
            )

            for visual_idx, visual in enumerate(visuals_with_projections):
                total_visuals += 1

                # Parse the config JSON string
                config_str = visual.get("config", "{}")
                try:
                    config = json.loads(config_str)
                except json.JSONDecodeError:
                    continue

                single_visual = config.get("singleVisual", {})
                visual_type = clean_text(single_visual.get("visualType", "Unknown"))
                projections = single_visual.get("projections", {})

                log_file.write(f"Visual #{visual_idx + 1}\n")
                log_file.write(f"  Type: {visual_type}\n")

                # Extract columns and measures from projections
                log_file.write(f"  Projections:\n")

                # Collect projection data
                projection_details = []
                for proj_type, proj_items in projections.items():
                    if proj_items:
                        log_file.write(f"    {proj_type}:\n")
                        for item in proj_items:
                            query_ref = clean_text(item.get("queryRef", "N/A"))
                            active = item.get("active", True)
                            log_file.write(f"      - {query_ref} (Active: {active})\n")
                            projection_details.append(
                                {
                                    "projection_type": proj_type,
                                    "query_ref": query_ref,
                                    "active": active,
                                }
                            )

                # Parse dataTransforms if available for additional metadata
                data_transforms_str = visual.get("dataTransforms", "{}")
                field_details = []
                try:
                    data_transforms = json.loads(data_transforms_str)
                    selects = data_transforms.get("selects", [])

                    if selects:
                        log_file.write(f"  Fields Details:\n")
                        for select in selects:
                            display_name = clean_text(select.get("displayName", "N/A"))
                            query_name = clean_text(select.get("queryName", "N/A"))
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
                                agg_func = expr["Aggregation"].get(
                                    "Function", "Unknown"
                                )
                                aggregation_func = f"Function {agg_func}"
                                is_measure = True

                            log_file.write(f"    - Display Name: {display_name}\n")
                            log_file.write(f"      Query Name: {query_name}\n")
                            log_file.write(f"      Type: {type_name}\n")
                            if format_info != "N/A":
                                log_file.write(f"      Format: {format_info}\n")
                            if aggregation_func:
                                log_file.write(
                                    f"      Aggregation: {aggregation_func}\n"
                                )

                            field_details.append(
                                {
                                    "display_name": display_name,
                                    "query_name": query_name,
                                    "type": type_name,
                                    "format": format_info,
                                    "aggregation": aggregation_func,
                                    "is_measure": is_measure,
                                }
                            )

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
                            "Page Name": page_name,
                            "Visual Type": visual_type,
                            "Field Display Name": field["display_name"],
                            "Field Query Name": field["query_name"],
                            "Field Type": field["type"],
                            "Field Format": field["format"]
                            if field["format"] != "N/A"
                            else "",
                            "Is Measure": "Yes" if field["is_measure"] else "No",
                            "Aggregation": field["aggregation"],
                            "Projection Type": matching_projection["projection_type"]
                            if matching_projection
                            else "",
                            "Active": matching_projection["active"]
                            if matching_projection
                            else "",
                        }
                        csv_data.append(csv_row)
                else:
                    # If no field details, add basic visual info
                    for proj in projection_details:
                        csv_row = {
                            "Page Name": page_name,
                            "Visual Type": visual_type,
                            "Field Display Name": "",
                            "Field Query Name": proj["query_ref"],
                            "Field Type": "",
                            "Field Format": "",
                            "Is Measure": "",
                            "Aggregation": "",
                            "Projection Type": proj["projection_type"],
                            "Active": proj["active"],
                        }
                        csv_data.append(csv_row)

                log_file.write(f"\n")

        log_file.write(f"\n{'=' * 80}\n")
        log_file.write(f"SUMMARY\n")
        log_file.write(f"{'=' * 80}\n")
        log_file.write(f"Total Pages: {len(sections)}\n")
        log_file.write(f"Total Visuals with Projections: {total_visuals}\n")
        log_file.write(f"{'=' * 80}\n")

        # Write CSV file
        if csv_data:
            with open(csv_file_path, "w", newline="", encoding="utf-8") as csv_file:
                fieldnames = [
                    "Page Name",
                    "Visual Type",
                    "Field Display Name",
                    "Field Query Name",
                    "Field Type",
                    "Field Format",
                    "Is Measure",
                    "Aggregation",
                    "Projection Type",
                    "Active",
                ]
                writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(csv_data)

            print(f"CSV file created: {csv_file_path}")
            log_file.write(f"\nCSV file created: {csv_file_path}\n")

    except Exception as e:
        log_file.write(f"\nError parsing visual containers: {str(e)}\n")


def extract_pbix_contents(
    pbix_file_path, output_log_path="output.log", output_csv_path="visuals_data.csv"
):
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
                layout_file_path = None

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

                        # Check if this is the Layout file
                        if file_path == "Report/Layout":
                            layout_file_path = file_path

                # Summary
                log_file.write(f"\n{'=' * 80}\n")
                log_file.write(f"Summary:\n")
                log_file.write(f"  Total Folders: {len(folders)}\n")
                log_file.write(f"  Total Files: {len(files)}\n")
                log_file.write(f"{'=' * 80}\n\n")

                # Extract and write Layout file contents
                layout_json = None
                if layout_file_path:
                    log_file.write(f"\n{'=' * 80}\n")
                    log_file.write(f"LAYOUT FILE CONTENTS (Report/Layout)\n")
                    log_file.write(f"{'=' * 80}\n\n")

                    try:
                        # Read the Layout file content
                        with zip_ref.open(layout_file_path) as layout_file:
                            layout_bytes = layout_file.read()

                            # Try different encodings
                            layout_content = None
                            encodings_to_try = ["utf-16-le", "utf-16", "utf-8"]

                            for encoding in encodings_to_try:
                                try:
                                    layout_content = layout_bytes.decode(encoding)
                                    # If successful, break
                                    break
                                except (UnicodeDecodeError, UnicodeError):
                                    continue

                            if layout_content is None:
                                # Fallback to utf-8 with error handling
                                layout_content = layout_bytes.decode(
                                    "utf-8", errors="replace"
                                )

                            # Try to parse as JSON and pretty print
                            try:
                                layout_json = json.loads(layout_content)
                                formatted_json = json.dumps(
                                    layout_json, indent=2, ensure_ascii=False
                                )
                                log_file.write(formatted_json)
                            except json.JSONDecodeError:
                                # If not valid JSON, write as-is
                                log_file.write(layout_content)

                        log_file.write(f"\n\n{'=' * 80}\n")
                        log_file.write(f"END OF LAYOUT FILE\n")
                        log_file.write(f"{'=' * 80}\n")

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
