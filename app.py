import streamlit as st
import pandas as pd
from main import extract_visuals_data
import io

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
    </style>
    """, unsafe_allow_html=True)

# Title and description
st.title("📊 Power BI Metadata Extractor")
st.markdown("""
Extract comprehensive metadata from Power BI (.pbix) files including:
- **Report Summary**: Pages and visuals count
- **Page Details**: Visual count and filters per page  
- **Visual Details**: Fields, measures, data types, and filters
""")

# File uploader
uploaded_file = st.file_uploader("Upload a Power BI (.pbix) file", type=['pbix'])

if uploaded_file is not None:
    try:
        # Show processing message
        with st.spinner('🔄 Processing PBIX file...'):
            # Extract comprehensive report metadata
            report_data = extract_visuals_data(uploaded_file)
        
        if report_data and report_data.get("visuals"):
            summary = report_data.get("summary", {})
            pages = report_data.get("pages", [])
            visuals = report_data.get("visuals", [])
            
            # Convert to DataFrames
            df_pages = pd.DataFrame(pages)
            df_visuals = pd.DataFrame(visuals)
            
            # ========== REPORT SUMMARY ==========
            st.header("📈 Report Summary")
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
                    # Display pages table
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
                    
                    # Page selection for drill-down
                    st.divider()
                    st.subheader("🔍 Page Drill-Down")
                    
                    selected_page = st.selectbox(
                        "Select a page to view its visuals:",
                        options=df_pages['Page Name'].unique()
                    )
                    
                    if selected_page:
                        page_visuals = df_visuals[df_visuals['Page Name'] == selected_page]
                        
                        # Show page stats
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Visuals on Page", page_visuals['Visual ID'].nunique())
                        with col2:
                            st.metric("Total Fields", len(page_visuals))
                        with col3:
                            measures_count = len(page_visuals[page_visuals['Is Measure'] == 'Yes'])
                            st.metric("Measures Used", measures_count)
                        
                        # Show visuals on this page
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
                
                # Filters
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    selected_pages = st.multiselect(
                        "Filter by Page",
                        options=df_visuals['Page Name'].unique(),
                        default=df_visuals['Page Name'].unique()
                    )
                
                with col2:
                    selected_visual_types = st.multiselect(
                        "Filter by Visual Type",
                        options=df_visuals['Visual Type'].unique(),
                        default=df_visuals['Visual Type'].unique()
                    )
                
                with col3:
                    selected_titles = st.multiselect(
                        "Filter by Title",
                        options=df_visuals['Visual Title'].unique(),
                        default=df_visuals['Visual Title'].unique()
                    )
                
                with col4:
                    measure_filter = st.selectbox(
                        "Filter by Type",
                        options=["All", "Measures Only", "Columns Only"]
                    )
                
                with col5:
                    has_filters = st.selectbox(
                        "Has Visual Filters",
                        options=["All", "With Filters", "Without Filters"]
                    )
                
                # Apply filters - Update this section (around line 166-175)
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
                
                # Display filtered data
                st.markdown(f"**Showing {len(filtered_df)} records**")
                
                # Group by visual for better readability - Update the grouping (around line 190)
                if st.checkbox("Group by Visual", value=True):
                    unique_visuals = filtered_df.groupby(['Page Name', 'Visual ID', 'Visual Title', 'Visual Type']).size().reset_index(name='Field Count')
                    
                    for idx, visual in unique_visuals.iterrows():
                        title_display = f"{visual['Visual Title']} ({visual['Visual Type']})" 
                        with st.expander(f"🔹 {title_display} on {visual['Page Name']} ({visual['Field Count']} fields)"):
                            visual_data = filtered_df[
                                (filtered_df['Page Name'] == visual['Page Name']) &
                                (filtered_df['Visual ID'] == visual['Visual ID'])
                            ]
                            
                            # Display fields (without filters column)
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
                
                # Visual Type Distribution
                st.divider()
                st.subheader("📊 Visual Type Distribution")
                visual_type_counts = filtered_df.groupby('Visual Type')['Visual ID'].nunique()
                st.bar_chart(visual_type_counts)
            
            # TAB 3: FILTERS
            with tab3:
                st.subheader("🔍 Filter Analysis")
                
                # Create filters breakdown table
                filters_data = []
                
                # Process page filters
                for _, page in df_pages.iterrows():
                    page_filters_str = page['Page Filters']
                    if page_filters_str and page_filters_str != "None":
                        # Split by pipe separator
                        filter_items = page_filters_str.split(' | ')
                        for filter_item in filter_items:
                            # Parse filter: "Table.Column: Condition"
                            if ':' in filter_item:
                                field_part, condition_part = filter_item.split(':', 1)
                                if '.' in field_part:
                                    table, column = field_part.rsplit('.', 1)
                                    filters_data.append({
                                        'Filter Level': 'Page',
                                        'Page Name': page['Page Name'],
                                        'Visual Type': 'N/A',
                                        'Table': table.strip(),
                                        'Column': column.strip(),
                                        'Condition': condition_part.strip()
                                    })
                
                # Process visual filters
                unique_visual_filters = df_visuals[['Page Name', 'Visual ID', 'Visual Title', 'Visual Type', 'Visual Filters']].drop_duplicates()
                
                for _, visual in unique_visual_filters.iterrows():
                    visual_filters_str = visual['Visual Filters']
                    if visual_filters_str:
                        # Split by pipe separator
                        filter_items = visual_filters_str.split(' | ')
                        for filter_item in filter_items:
                            # Parse filter: "Table.Column: Condition"
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
                    
                    # Summary metrics
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
                    
                    # Filter controls
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        filter_level_options = st.multiselect(
                            "Filter Level",
                            options=df_filters['Filter Level'].unique(),
                            default=df_filters['Filter Level'].unique()
                        )
                    with col2:
                        page_filter_options = st.multiselect(
                            "Page",
                            options=df_filters['Page Name'].unique(),
                            default=df_filters['Page Name'].unique()
                        )
                    with col3:
                        table_filter_options = st.multiselect(
                            "Table",
                            options=sorted(df_filters['Table'].unique()),
                            default=df_filters['Table'].unique()
                        )
                    
                    # Apply filter selections
                    filtered_filters_df = df_filters[
                        (df_filters['Filter Level'].isin(filter_level_options)) &
                        (df_filters['Page Name'].isin(page_filter_options)) &
                        (df_filters['Table'].isin(table_filter_options))
                    ]
                    
                    st.divider()
                    st.markdown(f"**Showing {len(filtered_filters_df)} filters**")
                    
                    # Display filters table
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
                    
                    # Filters by table chart
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
                
                # Export Page Summary
                with col1:
                    st.markdown("### 📄 Page Summary")
                    csv_pages = df_pages.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Page Summary (CSV)",
                        data=csv_pages,
                        file_name="pbix_pages_summary.csv",
                        mime="text/csv",
                    )
                
                # Export Visual Details
                with col2:
                    st.markdown("### 📊 Visual Details")
                    csv_visuals = df_visuals.drop(columns=['Visual ID', 'Visual Filters'], errors='ignore').to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Visual Details (CSV)",
                        data=csv_visuals,
                        file_name="pbix_visuals_details.csv",
                        mime="text/csv",
                    )
                
                st.divider()
                
                # Export Filters
                if filters_data:
                    st.markdown("### 🔍 Filters")
                    csv_filters = df_filters.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Filters (CSV)",
                        data=csv_filters,
                        file_name="pbix_filters.csv",
                        mime="text/csv",
                    )
                    
                    st.divider()
                
                # Export as Excel (all sheets)
                st.markdown("### 📑 Complete Report (Excel)")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Summary sheet
                    pd.DataFrame([summary]).to_excel(writer, index=False, sheet_name='Summary')
                    # Pages sheet
                    df_pages.to_excel(writer, index=False, sheet_name='Pages')
                    # Visuals sheet
                    df_visuals.drop(columns=['Visual ID', 'Visual Filters'], errors='ignore').to_excel(writer, index=False, sheet_name='Visual Details')
                    # Filters sheet
                    if filters_data:
                        df_filters.to_excel(writer, index=False, sheet_name='Filters')
                
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Download Complete Report (Excel)",
                    data=excel_data,
                    file_name="pbix_complete_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            
        else:
            st.warning("⚠️ No visuals with data projections found in the PBIX file.")
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        with st.expander("Show Error Details"):
            st.exception(e)
else:
    # Show instructions when no file is uploaded
    st.info("👆 Please upload a Power BI (.pbix) file to get started.")
    
    with st.expander("ℹ️ How to use this application"):
        st.markdown("""
        ### Getting Started
        1. **Upload a PBIX file** using the file uploader above
        2. **View Report Summary** to see overall statistics
        3. **Explore Page Overview** to see page-level details and filters
        4. **Drill into Visual Details** to see fields, measures, and visual-level filters
        5. **Analyze Filters** in the dedicated Filters tab with one filter per row
        6. **Export data** as CSV or Excel for further analysis
        
        ### Features
        - 📄 **Page Overview**: See all pages with visual counts and page-level filters
        - 📊 **Visual Details**: Detailed information about each visual
        - 🔍 **Filters Tab**: Dedicated view showing all filters (page & visual level) - one per row
        - 📥 **Export**: Download data in CSV or Excel format
        
        ### What's Extracted
        - Report summary (total pages and visuals)
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