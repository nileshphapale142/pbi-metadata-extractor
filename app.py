import streamlit as st
import pandas as pd
from main import extract_visuals_data
import io
import os
from pathlib import Path

# Page configuration
st.set_page_config(
    page_title="Power BI Metadata Extractor",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .metric-card {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #ff6b6b;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .sidebar .sidebar-content {
        background-color: #f0f2f6;
    }
    </style>
    """, unsafe_allow_html=True)

# Sidebar for processing mode selection
st.sidebar.title("⚙️ Processing Mode")
st.sidebar.markdown("---")

processing_mode = st.sidebar.radio(
    "Select how you want to process PBIX files:",
    options=[
        "📄 Single File",
        "📚 Multiple Files",
        "📁 Directory (Batch)"
    ],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
### Processing Modes:
- **Single File**: Upload and analyze one PBIX file
- **Multiple Files**: Upload and compare multiple PBIX files
- **Directory**: Process all PBIX files in a folder
""")

# Title and description
st.title("📊 Power BI Metadata Extractor")

def process_single_pbix(uploaded_file, file_name="Report"):
    """Process a single PBIX file and return the data"""
    try:
        with st.spinner(f'🔄 Processing {file_name}...'):
            report_data = extract_visuals_data(uploaded_file)
        return report_data
    except Exception as e:
        st.error(f"❌ Error processing {file_name}: {str(e)}")
        return None

def display_report_data(report_data, report_name="Report"):
    """Display report data in tabs"""
    
    if not report_data or not report_data.get("visuals"):
        st.warning(f"⚠️ No visuals with data projections found in {report_name}.")
        return
    
    summary = report_data.get("summary", {})
    pages = report_data.get("pages", [])
    visuals = report_data.get("visuals", [])
    
    # Convert to DataFrames
    df_pages = pd.DataFrame(pages)
    df_visuals = pd.DataFrame(visuals)
    
    # ========== REPORT SUMMARY ==========
    st.header(f"📈 Report Summary - {report_name}")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="📄 Total Pages",
            value=summary.get("Total Pages", 0)
        )
    with col2:
        st.metric(
            label="📊 Total Visuals",
            value=summary.get("Total Visuals", 0)
        )
    with col3:
        st.metric(
            label="🔢 Total Fields",
            value=len(df_visuals)
        )
    
    st.divider()
    
    # ========== TABS FOR DIFFERENT VIEWS ==========
    tab1, tab2, tab3, tab4 = st.tabs(["📄 Page Overview", "📊 Visual Details", "🔍 Filters", "📥 Export Data"])
    
    # TAB 1: PAGE OVERVIEW
    with tab1:
        st.subheader("📄 Pages Summary")
        
        if not df_pages.empty:
            st.dataframe(
                df_pages,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Page Name": st.column_config.TextColumn("Page Name", width="medium"),
                    "Visual Count": st.column_config.NumberColumn("Visual Count", width="small"),
                    "Page Filters": st.column_config.TextColumn("Page Filters", width="large")
                }
            )
            
            st.divider()
            st.subheader("🔍 Page Drill-Down")
            
            selected_page = st.selectbox(
                "Select a page to view its visuals:",
                options=df_pages['Page Name'].unique(),
                key=f"page_select_{report_name}"
            )
            
            if selected_page:
                page_visuals = df_visuals[df_visuals['Page Name'] == selected_page]
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Visuals on Page", page_visuals['Visual ID'].nunique())
                with col2:
                    st.metric("Total Fields", len(page_visuals))
                with col3:
                    measures_count = len(page_visuals[page_visuals['Is Measure'] == 'Yes'])
                    st.metric("Measures Used", measures_count)
                
                st.markdown("**Visuals on this page:**")
                page_visual_types = page_visuals.groupby('Visual Type')['Visual ID'].nunique().reset_index()
                page_visual_types.columns = ['Visual Type', 'Count']
                
                col1, col2 = st.columns([1, 2])
                with col1:
                    st.dataframe(page_visual_types, use_container_width=True, hide_index=True)
                with col2:
                    st.bar_chart(page_visual_types.set_index('Visual Type'))
        else:
            st.info("No page data available")
    
    # TAB 2: VISUAL DETAILS
    with tab2:
        st.subheader("📊 Visual Details")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            selected_pages = st.multiselect(
                "Filter by Page",
                options=df_visuals['Page Name'].unique(),
                default=df_visuals['Page Name'].unique(),
                key=f"pages_{report_name}"
            )
        
        with col2:
            selected_visual_types = st.multiselect(
                "Filter by Visual Type",
                options=df_visuals['Visual Type'].unique(),
                default=df_visuals['Visual Type'].unique(),
                key=f"types_{report_name}"
            )
        
        with col3:
            selected_titles = st.multiselect(
                "Filter by Title",
                options=df_visuals['Visual Title'].unique(),
                default=df_visuals['Visual Title'].unique(),
                key=f"titles_{report_name}"
            )
        
        with col4:
            measure_filter = st.selectbox(
                "Filter by Type",
                options=["All", "Measures Only", "Columns Only"],
                key=f"measure_{report_name}"
            )
        
        with col5:
            has_filters = st.selectbox(
                "Has Visual Filters",
                options=["All", "With Filters", "Without Filters"],
                key=f"filters_{report_name}"
            )
        
        filtered_df = df_visuals[
            (df_visuals['Page Name'].isin(selected_pages)) &
            (df_visuals['Visual Type'].isin(selected_visual_types)) &
            (df_visuals['Visual Title'].isin(selected_titles))
        ]
        
        if measure_filter == "Measures Only":
            filtered_df = filtered_df[filtered_df['Is Measure'] == 'Yes']
        elif measure_filter == "Columns Only":
            filtered_df = filtered_df[filtered_df['Is Measure'] == 'No']
        
        if has_filters == "With Filters":
            filtered_df = filtered_df[filtered_df['Visual Filters'].str.len() > 0]
        elif has_filters == "Without Filters":
            filtered_df = filtered_df[filtered_df['Visual Filters'].str.len() == 0]
        
        st.divider()
        st.markdown(f"**Showing {len(filtered_df)} records**")
        
        if st.checkbox("Group by Visual", value=True, key=f"group_{report_name}"):
            unique_visuals = filtered_df.groupby(['Page Name', 'Visual ID', 'Visual Title', 'Visual Type']).size().reset_index(name='Field Count')
            
            for idx, visual in unique_visuals.iterrows():
                title_display = f"{visual['Visual Title']} ({visual['Visual Type']})" 
                with st.expander(f"🔹 {title_display} on {visual['Page Name']} ({visual['Field Count']} fields)"):
                    visual_data = filtered_df[
                        (filtered_df['Page Name'] == visual['Page Name']) &
                        (filtered_df['Visual ID'] == visual['Visual ID'])
                    ]
                    
                    st.dataframe(
                        visual_data[[
                            'Field Display Name', 'Field Query Name', 'Field Type', 
                            'Field Format', 'Is Measure', 'Aggregation', 'Projection Type'
                        ]],
                        use_container_width=True,
                        hide_index=True
                    )
        else:
            st.dataframe(
                filtered_df.drop(columns=['Visual ID', 'Visual Filters'], errors='ignore'),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Page Name": st.column_config.TextColumn("Page", width="small"),
                    "Visual Title": st.column_config.TextColumn("Title", width="medium"),
                    "Visual Type": st.column_config.TextColumn("Visual Type", width="small"),
                    "Field Display Name": st.column_config.TextColumn("Field", width="medium"),
                    "Field Type": st.column_config.TextColumn("Type", width="small"),
                    "Is Measure": st.column_config.TextColumn("Measure?", width="small"),
                }
            )
        
        st.divider()
        st.subheader("📊 Visual Type Distribution")
        visual_type_counts = filtered_df.groupby('Visual Type')['Visual ID'].nunique()
        st.bar_chart(visual_type_counts)
    
    # TAB 3: FILTERS
    with tab3:
        st.subheader("🔍 Filter Analysis")
        
        filters_data = []
        
        for _, page in df_pages.iterrows():
            page_filters_str = page['Page Filters']
            if page_filters_str and page_filters_str != "None":
                filter_items = page_filters_str.split(' | ')
                for filter_item in filter_items:
                    if ':' in filter_item:
                        field_part, condition_part = filter_item.split(':', 1)
                        if '.' in field_part:
                            table, column = field_part.rsplit('.', 1)
                            filters_data.append({
                                'Filter Level': 'Page',
                                'Page Name': page['Page Name'],
                                'Visual Title': 'N/A',
                                'Visual Type': 'N/A',
                                'Table': table.strip(),
                                'Column': column.strip(),
                                'Condition': condition_part.strip()
                            })
        
        unique_visual_filters = df_visuals[['Page Name', 'Visual ID', 'Visual Title', 'Visual Type', 'Visual Filters']].drop_duplicates()
        
        for _, visual in unique_visual_filters.iterrows():
            visual_filters_str = visual['Visual Filters']
            if visual_filters_str:
                filter_items = visual_filters_str.split(' | ')
                for filter_item in filter_items:
                    if ':' in filter_item:
                        field_part, condition_part = filter_item.split(':', 1)
                        if '.' in field_part:
                            table, column = field_part.rsplit('.', 1)
                            filters_data.append({
                                'Filter Level': 'Visual',
                                'Page Name': visual['Page Name'],
                                'Visual Title': visual['Visual Title'],
                                'Visual Type': visual['Visual Type'],
                                'Table': table.strip(),
                                'Column': column.strip(),
                                'Condition': condition_part.strip()
                            })
        
        if filters_data:
            df_filters = pd.DataFrame(filters_data)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Filters", len(df_filters))
            with col2:
                page_filters_count = len(df_filters[df_filters['Filter Level'] == 'Page'])
                st.metric("Page Filters", page_filters_count)
            with col3:
                visual_filters_count = len(df_filters[df_filters['Filter Level'] == 'Visual'])
                st.metric("Visual Filters", visual_filters_count)
            with col4:
                unique_tables = df_filters['Table'].nunique()
                st.metric("Tables Filtered", unique_tables)
            
            st.divider()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                filter_level_options = st.multiselect(
                    "Filter Level",
                    options=df_filters['Filter Level'].unique(),
                    default=df_filters['Filter Level'].unique(),
                    key=f"filter_level_{report_name}"
                )
            with col2:
                page_filter_options = st.multiselect(
                    "Page",
                    options=df_filters['Page Name'].unique(),
                    default=df_filters['Page Name'].unique(),
                    key=f"page_filter_{report_name}"
                )
            with col3:
                table_filter_options = st.multiselect(
                    "Table",
                    options=sorted(df_filters['Table'].unique()),
                    default=df_filters['Table'].unique(),
                    key=f"table_filter_{report_name}"
                )
            
            filtered_filters_df = df_filters[
                (df_filters['Filter Level'].isin(filter_level_options)) &
                (df_filters['Page Name'].isin(page_filter_options)) &
                (df_filters['Table'].isin(table_filter_options))
            ]
            
            st.divider()
            st.markdown(f"**Showing {len(filtered_filters_df)} filters**")
            
            st.dataframe(
                filtered_filters_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Filter Level": st.column_config.TextColumn("Level", width="small"),
                    "Page Name": st.column_config.TextColumn("Page", width="medium"),
                    "Visual Title": st.column_config.TextColumn("Title", width="medium"),
                    "Visual Type": st.column_config.TextColumn("Visual", width="small"),
                    "Table": st.column_config.TextColumn("Table", width="medium"),
                    "Column": st.column_config.TextColumn("Column", width="medium"),
                    "Condition": st.column_config.TextColumn("Condition", width="large")
                }
            )
            
            st.divider()
            st.subheader("📊 Filters by Table")
            filters_by_table = filtered_filters_df['Table'].value_counts()
            st.bar_chart(filters_by_table)
            
        else:
            st.info("No filters found in the report.")
    
    # TAB 4: EXPORT DATA
    with tab4:
        st.subheader("💾 Export Data")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 📄 Page Summary")
            csv_pages = df_pages.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Page Summary (CSV)",
                data=csv_pages,
                file_name=f"{report_name}_pages_summary.csv",
                mime="text/csv",
                key=f"export_pages_{report_name}"
            )
        
        with col2:
            st.markdown("### 📊 Visual Details")
            csv_visuals = df_visuals.drop(columns=['Visual ID', 'Visual Filters'], errors='ignore').to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Visual Details (CSV)",
                data=csv_visuals,
                file_name=f"{report_name}_visuals_details.csv",
                mime="text/csv",
                key=f"export_visuals_{report_name}"
            )
        
        st.divider()
        
        if filters_data:
            st.markdown("### 🔍 Filters")
            csv_filters = df_filters.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download Filters (CSV)",
                data=csv_filters,
                file_name=f"{report_name}_filters.csv",
                mime="text/csv",
                key=f"export_filters_{report_name}"
            )
            
            st.divider()
        
        st.markdown("### 📑 Complete Report (Excel)")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame([summary]).to_excel(writer, index=False, sheet_name='Summary')
            df_pages.to_excel(writer, index=False, sheet_name='Pages')
            df_visuals.drop(columns=['Visual ID', 'Visual Filters'], errors='ignore').to_excel(writer, index=False, sheet_name='Visual Details')
            if filters_data:
                df_filters.to_excel(writer, index=False, sheet_name='Filters')
        
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 Download Complete Report (Excel)",
            data=excel_data,
            file_name=f"{report_name}_complete_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"export_excel_{report_name}"
        )

# ========== MAIN APPLICATION LOGIC ==========

if processing_mode == "📄 Single File":
    st.markdown("""
    Upload a single Power BI (.pbix) file to extract:
    - **Report Summary**: Pages and visuals count
    - **Page Details**: Visual count and filters per page  
    - **Visual Details**: Fields, measures, data types, and filters
    """)
    
    uploaded_file = st.file_uploader("Upload a Power BI (.pbix) file", type=['pbix'])
    
    if uploaded_file is not None:
        report_data = process_single_pbix(uploaded_file, uploaded_file.name)
        if report_data:
            display_report_data(report_data, uploaded_file.name.replace('.pbix', ''))

elif processing_mode == "📚 Multiple Files":
    st.markdown("""
    Upload multiple Power BI (.pbix) files to compare and analyze them side-by-side.
    """)
    
    uploaded_files = st.file_uploader(
        "Upload Power BI (.pbix) files", 
        type=['pbix'], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
        
        # Process all files
        all_reports_data = {}
        for uploaded_file in uploaded_files:
            report_name = uploaded_file.name.replace('.pbix', '')
            report_data = process_single_pbix(uploaded_file, uploaded_file.name)
            if report_data:
                all_reports_data[report_name] = report_data
        
        if all_reports_data:
            # Comparison Summary
            st.header("📊 Comparison Summary")
            
            comparison_data = []
            for report_name, report_data in all_reports_data.items():
                summary = report_data.get("summary", {})
                visuals = report_data.get("visuals", [])
                comparison_data.append({
                    "Report Name": report_name,
                    "Total Pages": summary.get("Total Pages", 0),
                    "Total Visuals": summary.get("Total Visuals", 0),
                    "Total Fields": len(visuals)
                })
            
            df_comparison = pd.DataFrame(comparison_data)
            st.dataframe(df_comparison, use_container_width=True, hide_index=True)
            
            st.divider()
            
            # Report Selection
            st.subheader("📑 Select Report to View Details")
            selected_report = st.selectbox(
                "Choose a report:",
                options=list(all_reports_data.keys())
            )
            
            if selected_report:
                st.divider()
                display_report_data(all_reports_data[selected_report], selected_report)

elif processing_mode == "📁 Directory (Batch)":
    st.markdown("""
    Process all Power BI (.pbix) files from a directory.
    
    **Note**: Enter the full path to a directory containing PBIX files.
    """)
    
    directory_path = st.text_input(
        "Enter directory path:",
        placeholder="C:/Users/YourName/Documents/PowerBI Reports"
    )
    
    if directory_path and st.button("🔍 Process Directory"):
        if os.path.isdir(directory_path):
            pbix_files = list(Path(directory_path).glob("*.pbix"))
            
            if pbix_files:
                st.success(f"✅ Found {len(pbix_files)} PBIX file(s) in the directory!")
                
                # Process all files
                all_reports_data = {}
                progress_bar = st.progress(0)
                
                for idx, pbix_file in enumerate(pbix_files):
                    report_name = pbix_file.stem
                    with open(pbix_file, 'rb') as f:
                        report_data = process_single_pbix(f, pbix_file.name)
                        if report_data:
                            all_reports_data[report_name] = report_data
                    
                    progress_bar.progress((idx + 1) / len(pbix_files))
                
                progress_bar.empty()
                
                if all_reports_data:
                    # Comparison Summary
                    st.header("📊 Batch Processing Summary")
                    
                    comparison_data = []
                    for report_name, report_data in all_reports_data.items():
                        summary = report_data.get("summary", {})
                        visuals = report_data.get("visuals", [])
                        comparison_data.append({
                            "Report Name": report_name,
                            "Total Pages": summary.get("Total Pages", 0),
                            "Total Visuals": summary.get("Total Visuals", 0),
                            "Total Fields": len(visuals)
                        })
                    
                    df_comparison = pd.DataFrame(comparison_data)
                    st.dataframe(df_comparison, use_container_width=True, hide_index=True)
                    
                    # Export batch summary
                    st.divider()
                    st.subheader("📥 Export Batch Summary")
                    csv_batch = df_comparison.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Batch Summary (CSV)",
                        data=csv_batch,
                        file_name="batch_processing_summary.csv",
                        mime="text/csv",
                    )
                    
                    st.divider()
                    
                    # Report Selection
                    st.subheader("📑 Select Report to View Details")
                    selected_report = st.selectbox(
                        "Choose a report:",
                        options=list(all_reports_data.keys()),
                        key="batch_report_select"
                    )
                    
                    if selected_report:
                        st.divider()
                        display_report_data(all_reports_data[selected_report], selected_report)
            else:
                st.warning("⚠️ No PBIX files found in the specified directory.")
        else:
            st.error("❌ Invalid directory path. Please enter a valid path.")

# Show instructions when no file is uploaded
if not uploaded_file if processing_mode == "📄 Single File" else (not uploaded_files if processing_mode == "📚 Multiple Files" else not directory_path):
    with st.expander("ℹ️ How to use this application"):
        st.markdown("""
        ### Getting Started
        
        **Choose a processing mode from the sidebar:**
        
        1. **📄 Single File**: Upload and analyze one PBIX file at a time
        2. **📚 Multiple Files**: Upload multiple files to compare and analyze
        3. **📁 Directory (Batch)**: Process all PBIX files in a folder automatically
        
        ### Features
        - 📄 **Page Overview**: See all pages with visual counts and page-level filters
        - 📊 **Visual Details**: Detailed information about each visual with title support
        - 🔍 **Filters Tab**: Dedicated view showing all filters (page & visual level) - one per row
        - 📥 **Export**: Download data in CSV or Excel format
        - 🔄 **Comparison**: Compare multiple reports side-by-side
        
        ### What's Extracted
        - Report summary (total pages and visuals)
        - Visual titles (with "[No Title]" indicator when absent)
        - Page-level and visual-level filters (parsed and displayed clearly)
        - Fields and measures used in each visual
        - Data types, formats, and aggregations
        - Projection types (Rows, Columns, Values, etc.)
        
        **Note:** Only visuals with data projections are included. Decorative elements are excluded.
        """)

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: gray; padding: 20px;'>
    <small>Power BI Metadata Extractor | Built with Streamlit & ❤️ by Nilesh Phapale</small>
</div>
""", unsafe_allow_html=True)