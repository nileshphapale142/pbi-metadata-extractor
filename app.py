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

# Title and description
st.title("📊 Power BI Metadata Extractor")
st.markdown("""
This application extracts visual metadata from Power BI (.pbix) files, including:
- Visual types and names
- Fields and measures used
- Data types and formats
- Aggregations and projections
""")

# File uploader
uploaded_file = st.file_uploader("Upload a Power BI (.pbix) file", type=['pbix'])

if uploaded_file is not None:
    try:
        # Show processing message
        with st.spinner('Processing PBIX file...'):
            # Extract visual data
            visuals_data = extract_visuals_data(uploaded_file)
        
        if visuals_data:
            # Convert to DataFrame
            df = pd.DataFrame(visuals_data)
            
            # Display summary metrics
            st.subheader("📈 Summary")
            col1, col2, col3= st.columns(3)
            
            with col1:
                st.metric("Total Records", len(df))
            with col3:
                st.metric("Sections", df['Page Name'].nunique())
            with col3:
                st.metric("Visual Types", df['Visual Type'].nunique())
            
            st.divider()
            
            # Filters
            st.subheader("🔍 Filters")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                selected_sections = st.multiselect(
                    "Select Sections",
                    options=df['Page Name'].unique(),
                    default=df['Page Name'].unique()
                )
            
            with col2:
                selected_visual_types = st.multiselect(
                    "Select Visual Types",
                    options=df['Visual Type'].unique(),
                    default=df['Visual Type'].unique()
                )
            
            with col3:
                measure_filter = st.selectbox(
                    "Filter by Measure",
                    options=["All", "Measures Only", "Columns Only"]
                )
            
            # Apply filters
            filtered_df = df[
                (df['Page Name'].isin(selected_sections)) &
                (df['Visual Type'].isin(selected_visual_types))
            ]
            
            if measure_filter == "Measures Only":
                filtered_df = filtered_df[filtered_df['Is Measure'] == 'Yes']
            elif measure_filter == "Columns Only":
                filtered_df = filtered_df[filtered_df['Is Measure'] == 'No']
            
            st.divider()
            
            # Display data table
            st.subheader("📋 Visual Metadata Table")
            st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Page Name": st.column_config.TextColumn("Page Name", width="small"),
                    "Visual Type": st.column_config.TextColumn("Visual Type", width="small"),
                    "Field Display Name": st.column_config.TextColumn("Field Name", width="medium"),
                    "Field Query Name": st.column_config.TextColumn("Column Expression", width="medium"),
                    "Field Type": st.column_config.TextColumn("Data Type", width="small"),
                    "Is Measure": st.column_config.TextColumn("Measure?", width="small"),
                    "Projection Type": st.column_config.TextColumn("Projection", width="small"),
                }
            )
            
            st.divider()
            
            # Download options
            st.subheader("💾 Download Data")
            col1, col2 = st.columns(2)
            
            with col1:
                # Download as CSV
                csv = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download as CSV",
                    data=csv,
                    file_name="pbix_visuals_metadata.csv",
                    mime="text/csv",
                )
            
            with col2:
                # Download as Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Visuals Metadata')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Download as Excel",
                    data=excel_data,
                    file_name="pbix_visuals_metadata.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            
            # Visual Type Distribution
            st.divider()
            st.subheader("📊 Visual Type Distribution")
            visual_type_counts = filtered_df['Visual Type'].value_counts()
            st.bar_chart(visual_type_counts)
            
        else:
            st.warning("⚠️ No visuals with projections found in the PBIX file.")
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.exception(e)
else:
    # Show instructions when no file is uploaded
    st.info("👆 Please upload a Power BI (.pbix) file to get started.")
    
    with st.expander("ℹ️ How to use this application"):
        st.markdown("""
        1. **Upload a PBIX file** using the file uploader above
        2. **Review the summary metrics** to see an overview of the visuals
        3. **Use filters** to narrow down the data you want to see
        4. **Browse the table** to see detailed information about each visual
        5. **Download the data** as CSV or Excel for further analysis
        
        **Note:** This tool extracts metadata from visuals that have data projections (fields/measures).
        Decorative visuals like text boxes, images, and shapes are excluded.
        """)

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: gray; padding: 20px;'>
    <small>Power BI Metadata Extractor | Built with Streamlit</small>
</div>
""", unsafe_allow_html=True)