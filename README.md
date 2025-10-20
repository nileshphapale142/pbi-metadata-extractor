# 📊 Power BI Metadata Extractor

A comprehensive Streamlit-based web application for extracting and analyzing metadata from Power BI (.pbix) files. This tool helps you understand report structure, visualizations, data sources, filters, and complexity metrics.

## ✨ Features

### 📄 Single File Analysis
- **Report Summary**: Overview of pages and visuals count
- **Page Details**: Visual count and filters per page
- **Visual Details**: Comprehensive field-level information including:
  - Visual titles and types
  - Fields and measures with data types
  - Aggregations and projections
  - Format information
- **Filter Analysis**: Dedicated view of all page-level and visual-level filters
- **Export Options**: Download results as CSV or Excel

### 📚 Multiple Files Comparison
- Upload and compare multiple PBIX files side-by-side
- Complexity scoring system with visual indicators
- Batch processing capabilities
- Cross-report analysis metrics including:
  - Total pages and visuals
  - Field and measure counts
  - Table usage statistics
  - Filter counts
  - Complexity scores (Low/Medium/High/Very High)

## 🚀 Getting Started

### Prerequisites

- Python 3.13 or higher
- pip or uv package manager
- Power BI (.pbix) files to analyze

### Installation

#### Option 1: Using `uv` (Recommended - Fast)

1. **Clone or download this repository**
   ```bash
   git clone <your-repo-url>
   cd pbi-metadata-extractor
   ```

2. **Install uv** (if not already installed)
   ```bash
   pip install uv
   ```

3. **Install dependencies using uv**
   ```bash
   uv sync
   ```

#### Option 2: Using `pip`

1. **Clone or download this repository**
   ```bash
   git clone <your-repo-url>
   cd pbi-metadata-extractor
   ```

2. **Create a virtual environment** (recommended)
   ```bash
   python -m venv venv
   
   # On Windows
   venv\Scripts\activate
   
   # On Linux/Mac
   source venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install streamlit pandas openpyxl lz4
   ```

## 🏃 Running the Application

### Using Streamlit

Run the following command in your terminal:

```bash
streamlit run app.py
```

The application will automatically open in your default web browser at `http://localhost:8501`

### Alternative Port

If port 8501 is already in use, specify a different port:

```bash
streamlit run app.py --server.port 8502
```

## 📖 How to Use

### Single File Analysis Mode

1. **Select Processing Mode**: Choose "📄 Single File Analysis" from the sidebar
2. **Upload File**: Click "Browse files" and select a single .pbix file
3. **Wait for Processing**: The app will extract metadata from the file
4. **Explore Results**: Navigate through four tabs:
   - **Page Overview**: Summary of pages and their visuals with drill-down capability
   - **Visual Details**: Detailed field-level information with filtering options
   - **Filters**: Comprehensive filter analysis with table-level insights
   - **Export Data**: Download results in CSV or Excel format

### Multiple Files Comparison Mode

1. **Select Processing Mode**: Choose "📚 Multiple Files Comparison" from the sidebar
2. **Upload Files**: Select multiple .pbix files (hold Ctrl/Cmd for multiple selection)
3. **View Comparison**: Review the comparison summary table showing:
   - Report metrics (pages, visuals, fields, measures, tables, filters)
   - Complexity scores and levels
4. **Drill Down**: Select any report to view detailed analysis
5. **Export**: Download comparison summary as CSV

## 📊 Understanding Complexity Scores

The app calculates a complexity score for each report using the following formula:

```
Complexity Score = (Pages × 10) + (Visuals × 5) + (Fields × 2) + (Measures × 3) + (Tables × 8) + (Filters × 4)
```

**Complexity Levels:**
- 🟢 **Low**: Score < 100 (Simple reports)
- 🟡 **Medium**: Score 100-499 (Moderate complexity)
- 🟠 **High**: Score 500-999 (Complex reports)
- 🔴 **Very High**: Score ≥ 1000 (Highly complex reports)

## 🗂️ Project Structure

```
pbi-metadata-extractor/
├── app.py                  # Main Streamlit application
├── pyproject.toml          # Project dependencies and metadata
├── uv.lock                 # Dependency lock file (uv)
├── README.md               # This file
├── src/                    # Source modules
│   ├── constants.py        # Application constants
│   ├── filters.py          # Filter extraction logic
│   ├── report.py           # Report metadata extraction
│   ├── utils.py            # Utility functions
│   └── visuals.py          # Visual extraction logic
```

## 🛠️ Dependencies

Core dependencies (automatically installed):
- **streamlit** (≥1.50.0): Web application framework
- **pandas**: Data manipulation and analysis
- **openpyxl** (≥3.1.5): Excel file generation
- **lz4** (≥4.4.4): Compression library for PBIX file extraction

## 🎯 Use Cases

- **Report Documentation**: Automatically document report structure and contents
- **Quality Assurance**: Verify report consistency across environments
- **Migration Planning**: Understand report complexity before migration
- **Development vs Production**: Compare different versions of reports
- **Team Collaboration**: Analyze reports created by different teams
- **Audit & Compliance**: Track filter usage and data sources

## 🔧 Troubleshooting

### Port Already in Use
```bash
streamlit run app.py --server.port 8502
```

### Module Not Found Error
Make sure all dependencies are installed:
```bash
pip install -r requirements.txt
# or
uv sync
```

### PBIX File Processing Error
- Ensure the .pbix file is not corrupted
- Verify the file is not password-protected
- Check that the file contains valid Power BI report structure

### Memory Issues with Large Files
If processing large PBIX files causes memory issues:
- Process files one at a time in Single File mode
- Close other applications to free up memory
- Consider increasing available system memory

## 📝 Notes

- The application processes PBIX files locally - no data is sent to external servers
- Extracted metadata includes only structure, not actual data values
- Processing time varies based on report complexity and size
- The app supports standard Power BI report formats

## 👨‍💻 Author

Built with Streamlit & ❤️ by Nilesh

## 📄 License

[Add your license information here]

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!

## 🔮 Future Enhancements

- DAX measure extraction
- Data model diagram visualization
- Relationship analysis
- Performance optimization suggestions
- Custom theme extraction
- Bookmark and button action analysis
