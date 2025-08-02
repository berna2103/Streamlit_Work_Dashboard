import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import datetime
from fpdf import FPDF
import base64
import plotly.io as pio
import io
import os
import urllib.parse

# Use kaleido for static image export
pio.kaleido.scope.default_format = "png"

TITLE = 'FSE Inventory Dashboard'
st.set_page_config(layout="wide")

st.title(TITLE)
st.subheader('Inventory')

# --- CSS for the red button ---
st.markdown('<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">', unsafe_allow_html=True)
st.markdown("""
<style>
.red-button {
    background-color: #FF4B4B; /* A nice red */
    color: white; /* White text for contrast */
    padding: 10px 20px;
    border: none;
    border-radius: 5px;
    cursor: pointer;
    font-size: 16px;
    font-weight: bold;
    text-align: center;
    text-decoration: none; /* Remove underline from <a> if button is wrapped */
    display: inline-block; /* Allows padding and width */
    margin-top: 10px; /* Some spacing */
    transition: background-color 0.3s ease; /* Smooth transition on hover */
    /* ... existing styles ... */
    display: inline-flex; /* Use inline-flex to align icon and text */
    align-items: center; /* Vertically align icon and text */
    gap: 8px; /* Space between icon and text */
    /* ... rest of styles ... */
}
.red-button:hover {
    background-color: #CC0000; /* Darker red on hover */
}
</style>
""", unsafe_allow_html=True)


st.sidebar.title('Settings')
uploaded_file = st.sidebar.file_uploader(label='Load inventory file:', key='inventory', type=['xlsx', 'xls', 'csv'])

st.sidebar.divider()

# Function to highlight cells based on conditions for Streamlit display
def highlight_days(val):
    if val > 90:
        return 'background-color: rgb(220,20,60)'
    elif 60 < val <= 90:
        return 'background-color: orange'
    elif 30 < val <= 60:
        return 'rgb(255,215,0)'
    elif val <= 30:
        return 'background-color: green'
    return ''

if uploaded_file is not None:
    # Read the Excel file
    excel_file = pd.ExcelFile(uploaded_file)

    sheet_names = pd.ExcelFile(uploaded_file).sheet_names

    selected_sheet = st.sidebar.selectbox('Select Sheet:', sheet_names)

    df = pd.read_excel(excel_file, sheet_name=selected_sheet)

    # Remove specified columns that are never needed
    columns_to_remove_initial = ['Age Of Inventory', 'Warehouse', 'Warehouse Location', 'Batch Number', 'Mandatory Return?']
    # Ensure 'Item Status' is correctly handled - rename it then drop it if it exists
    if 'Item Status' in df.columns:
        df.rename(columns={'Item Status': 'Days_Placeholder'}, inplace=True)
    
    # Remove columns not needed in any downstream processing or display
    columns_to_drop_after_initial = [col for col in columns_to_remove_initial if col in df.columns]
    if 'Days_Placeholder' in df.columns:
        columns_to_drop_after_initial.append('Days_Placeholder')

    df.drop(columns=columns_to_drop_after_initial, inplace=True)

    # Remove the last row (assuming it's a summary row)
    df = df.iloc[:-1]

    # Rename 'Age of inventory' if it still exists and relevant (from previous versions of code)
    if 'Age of inventory' in df.columns:
        df.rename(columns={'Age of inventory': 'age'}, inplace=True)


    # Convert Excel serial date to datetime and format as mm/dd/yyyy
    df['Receipt Date'] = pd.to_datetime(df['Receipt Date'], origin='1899-12-30', unit='D').dt.strftime('%m/%d/%Y')

    # Calculate the number of days from today's date minus Receipt Date
    df['Days'] = (pd.to_datetime(datetime.datetime.now().strftime('%m/%d/%Y')) - pd.to_datetime(df['Receipt Date'])).dt.days

  

    
    # Streamlit sidebar for selecting Stock Location (FSE)
    # --- NEW FILTERING LOGIC ---
    # 1. Filter by CLMmanagername first
    if 'CLMmanagername' in df.columns:
        all_managers = ['All'] + sorted(df['CLMmanagername'].unique().tolist()) # Add 'All' option
        selected_manager = st.sidebar.selectbox('Select CLM Manager:', all_managers)

        if selected_manager != 'All':
            df_filtered_by_manager = df[df['CLMmanagername'] == selected_manager].copy()
        else:
            df_filtered_by_manager = df.copy() # If 'All' is selected, use the entire df
    else:
        st.sidebar.warning("CLMmanagername column not found in the loaded file.")
        df_filtered_by_manager = df.copy() # Fallback if column is missing
        selected_manager = "N/A" # Set a default for consistency

    # 2. Then, filter by Stock Location (FSE), based on the manager-filtered data
    stock_locations_for_selection = ['All'] + sorted(df_filtered_by_manager['Stock Location'].unique().tolist()) # Add 'All' option
    selected_location = st.sidebar.selectbox('Select FSE (Stock Location):', stock_locations_for_selection)

    if selected_location != 'All':
        filtered_df = df_filtered_by_manager[df_filtered_by_manager['Stock Location'] == selected_location].copy()
    else:
        filtered_df = df_filtered_by_manager.copy() # If 'All' is selected, use the manager-filtered df

      # Display the modified dataframe
    # if st.sidebar.checkbox('Show Central 1', key='Main'):
    #     st.dataframe(selected_location.style.hide(axis='index'))

    # --- END NEW FILTERING LOGIC ---

    # Categorize 'Days' into 0-30, 31-60, 61-90, and over 90 for the dashboard
    age_bins = [0, 31, 61, 91, float('inf')]
    age_labels = ['0-30 days', '31-60 days', '61-90 days', 'Over 90 days']
    age_colors = ['green', 'rgb(255,215,0)', 'orange', 'rgb(220,20,60)']

    filtered_df['Age Category'] = pd.cut(filtered_df['Days'],
                                         bins=age_bins,
                                         labels=age_labels,
                                         right=False)

    # Aggregate data based on these categories for graphs
    quantity_by_category = filtered_df.groupby('Age Category')['Quantity'].sum().reindex(age_labels, fill_value=0)
    stock_value_by_category = filtered_df.groupby('Age Category')['Stock Value (Transfer Cost)'].sum().reindex(age_labels, fill_value=0)


    # --- Dashboard Graphs ---
    st.subheader(f"{selected_location}'s Inventory Overview")

    # Overall Donut Graph (existing)
    labels_donut = ['0-30', '31-60', '61-90', 'Over 90']
    fig_donut = make_subplots(rows=1, cols=2, specs=[[{'type': 'domain'}, {'type': 'domain'}]],
                        subplot_titles=[f'{int(quantity_by_category.sum())} parts', f'${round(stock_value_by_category.sum(), 2)}'])

    fig_donut.add_trace(go.Pie(labels=labels_donut, values=quantity_by_category, name="Quantity"),
                  1, 1)
    fig_donut.add_trace(go.Pie(labels=labels_donut, values=stock_value_by_category, name="Value"),
                  1, 2)

    fig_donut.update_traces(hole=0.5, hoverinfo="label+value+name", marker=dict(colors=age_colors))
    fig_donut.update_layout(annotations=[dict(y=1.05, font_size=20, showarrow=False),
                                     dict(y=1.05, font_size=19, showarrow=False)])
    st.plotly_chart(fig_donut, use_container_width=True)


    # New Bar Chart for Age Categories
    st.subheader(f"Inventory Distribution by Age Category for {selected_location}")
    fig_bar = make_subplots(rows=1, cols=2, subplot_titles=['Quantity by Age Category', 'Stock Value by Age Category'])

    # Quantity Bar Chart
    fig_bar.add_trace(go.Bar(
        x=quantity_by_category.index,
        y=quantity_by_category.values,
        marker_color=age_colors,
        name='Quantity'
    ), row=1, col=1)

    # Stock Value Bar Chart
    fig_bar.add_trace(go.Bar(
        x=stock_value_by_category.index,
        y=stock_value_by_category.values,
        marker_color=age_colors,
        name='Stock Value'
    ), row=1, col=2)

    fig_bar.update_layout(
        xaxis_title="Age Category",
        yaxis_title="Count / Value",
        showlegend=False,
    )
    st.plotly_chart(fig_bar, use_container_width=True)


    # Save the plotly figures as images for Excel and PDF generation
    temp_dir = "/tmp"
    if os.name == 'nt':
        temp_dir = os.path.join(os.environ.get('TEMP', 'C:\\Temp'))
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

    donut_fig_path = os.path.join(temp_dir, "donut_chart.png")
    bar_fig_path = os.path.join(temp_dir, "bar_chart.png")

    fig_donut.write_image(donut_fig_path)
    fig_bar.write_image(bar_fig_path)


    # --- Email Automation Section ---
    st.sidebar.divider()
    st.sidebar.subheader("Email FSE Inventory")

    # --- DYNAMICALLY GENERATE FSE Contact Data Mapping ---
    fse_contact_data = {}
    for location_name in df['Stock Location'].unique():
        parts = location_name.split(' ')
        if len(parts) >= 2:
            # Assume first word is first name, rest is last name
            first_name = parts[0]
            last_name = ' '.join(parts[1:]) # Handles multi-word last names
            fse_contact_data[location_name] = {"first_name": first_name, "last_name": last_name}
        else:
            # If not parsable as "First Last", use a generic name and warn the user
            fse_contact_data[location_name] = {"first_name": location_name, "last_name": "FSE"}
            st.sidebar.warning(
                f"Cannot parse a full name from Stock Location: '{location_name}'. "
                f"Using '{location_name}' as first name and 'FSE' as last name. "
                "Email will be formatted as '{location_name}.fse@elekta.com'."
            )
    # --- END DYNAMIC GENERATION ---

    # Get the FSE's details from the dynamically generated mapping
    # Using .get() provides a default if somehow a selected_location isn't in fse_contact_data (unlikely now)
    fse_details = fse_contact_data.get(selected_location, {"first_name": "Unknown", "last_name": "User"})
    fse_first_name = fse_details["first_name"]
    fse_last_name = fse_details["last_name"]

    # Construct the personalized email address (first.last@co.com)
    # Ensure names are lowercased for email
    fse_email = f"{fse_first_name.lower()}.{fse_last_name.lower()}@elekta.com"

    # Construct the personalized greeting (Hi Firstname)
    fse_greeting_name = fse_first_name # The name used in the greeting
    subject = f"Inventory Report"
    body = f"Hi {fse_greeting_name},\n\nFind attached a copy of your inventory. Please return parts over 60 days. If you notice a discrepancy let me know.\n\nThank you,\nBernardo"


    # Define columns to remove from the Excel/PDF table displays
    columns_to_remove_from_table = ['[Country Description]', 'Region', 'Stock Location', 'Business Lines', 'Age Category']
    
    # Create the DataFrame for Excel/PDF table by dropping unwanted columns
    # and then SORT by 'Days' column higher to lower
    table_display_df = filtered_df.drop(columns=columns_to_remove_from_table, errors='ignore').copy()
    if 'Days' in table_display_df.columns:
        table_display_df.sort_values(by='Days', ascending=False, inplace=True)


    # Function to create an Excel file in memory with conditional formatting, autofit, and embedded graphs
    def to_excel_in_memory_with_graphs(dataframe_for_excel, donut_img_path, bar_img_path, age_col_name='Days'):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        dataframe_for_excel.to_excel(writer, index=False, sheet_name='Inventory')

        workbook = writer.book
        worksheet = writer.sheets['Inventory']

        # Define formats for conditional formatting
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        orange_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
        yellow_format = workbook.add_format({'bg_color': '#FFFCB0', 'font_color': '#8B8000'})
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})

        # Get the column letter for 'Days'
        def get_excel_column_letter(col_idx):
            letter = ''
            while col_idx >= 0:
                letter = chr(65 + (col_idx % 26)) + letter
                col_idx = (col_idx // 26) - 1
            return letter

        if age_col_name in dataframe_for_excel.columns:
            days_col_idx = dataframe_for_excel.columns.get_loc(age_col_name)
            days_col_letter = get_excel_column_letter(days_col_idx)
            max_data_row = len(dataframe_for_excel) + 1
            days_range = f'{days_col_letter}2:{days_col_letter}{max_data_row}'

            # Apply conditional formatting rules
            worksheet.conditional_format(days_range, {'type': 'cell', 'criteria': '>', 'value': 90, 'format': red_format})
            worksheet.conditional_format(days_range, {'type': 'cell', 'criteria': 'between', 'minimum': 61, 'maximum': 90, 'format': orange_format})
            worksheet.conditional_format(days_range, {'type': 'cell', 'criteria': 'between', 'minimum': 31, 'maximum': 60, 'format': yellow_format})
            worksheet.conditional_format(days_range, {'type': 'cell', 'criteria': '<=', 'value': 30, 'format': green_format})

        # --- Autofit Columns ---
        for i, col in enumerate(dataframe_for_excel.columns):
            max_len = 0
            for row_val in dataframe_for_excel[col].values:
                cell_value_str = str(row_val) if pd.notna(row_val) else ""
                max_len = max(max_len, len(cell_value_str))

            header_len = len(str(col))
            calculated_width = max(max_len, header_len)
            final_width = min(calculated_width + 2, 80)
            worksheet.set_column(i, i, final_width)

        # --- Insert Graphs into Excel ---
        insert_row = len(dataframe_for_excel) + 2 + 5
        insert_col = 0

        worksheet.insert_image(insert_row, insert_col, donut_img_path,
                               {'x_scale': 0.7, 'y_scale': 0.7})

        insert_col_bar = insert_col + 4
        worksheet.insert_image(insert_row, insert_col_bar, bar_img_path,
                               {'x_scale': 0.7, 'y_scale': 0.7})


        writer.close()
        processed_data = output.getvalue()
        return processed_data

    # Generate Excel in memory for download with conditional formatting and autofit
    excel_download_data = to_excel_in_memory_with_graphs(table_display_df, donut_fig_path, bar_fig_path, age_col_name='Days')

    # Create a download button for the Excel file
    st.sidebar.download_button(
        label=f"Download {fse_greeting_name}'s Inventory Excel",
        data=excel_download_data,
        file_name=f"inventory_{fse_first_name.lower()}_{fse_last_name.lower()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Create a mailto link for subject and body (user will attach manually)
    mailto_link_for_body = (
        f"mailto:{fse_email}?"
        f"subject={urllib.parse.quote(subject)}&"
        f"body={urllib.parse.quote(body)}"
    )
    st.sidebar.markdown(
        f'<a href="{mailto_link_for_body}" target="_blank"><button button class="red-button"><i class="fa fa-envelope"></i>Open Email Draft for {fse_greeting_name}</button></a>',
        unsafe_allow_html=True
    )
    # st.sidebar.info("Please download the Excel file and attach it to the email manually. The Excel includes conditional formatting, auto-adjusted column widths, and embedded graphs, sorted by 'Days' (highest first).")

    # Function to generate PDF using fpdf
    def create_pdf_with_graphs(dataframe_for_pdf_table):
        pdf = FPDF(orientation='L')  # 'L' for Landscape orientation
        pdf.add_page()

        # Title
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 10, txt=TITLE, ln=True, align="C")

        # Add the donut chart image
        pdf.image(donut_fig_path, x=10, y=20, w=130)

        # Add the bar chart image
        pdf.image(bar_fig_path, x=145, y=20, w=130)

        # Add a space between the chart and table
        pdf.ln(100)

        # Add a small header for the table
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 10, txt="Inventory Data:", ln=True, align="L")

        # Add DataFrame content as a table
        pdf.set_font("Arial", size=8)
        pdf.set_fill_color(200, 220, 255)

        # Add table header
        headers = list(dataframe_for_pdf_table.columns)
        page_width = pdf.w - 2 * pdf.l_margin
        col_width = page_width / len(headers)
        for header in headers:
            pdf.cell(col_width, 10, str(header), border=1, align='C', fill=True)
        pdf.ln()

        # Add table rows
        for i, row in dataframe_for_pdf_table.iterrows():
            for val in row:
                pdf.cell(col_width, 10, str(val), border=1, align='C')
            pdf.ln()

        # Save the PDF to a file
        pdf_path = os.path.join(temp_dir, f"{fse_first_name} {fse_last_name} dashboard_report.pdf")
        pdf.output(pdf_path)
        return pdf_path


    # Function to get the download link for the PDF
    def get_pdf_download_link(pdf_path):
        with open(pdf_path, "rb") as pdf_file:
            base64_pdf = base64.b64encode(pdf_file.read()).decode('utf-8')
        href = f'<a href="data:application/octet-stream;base64,{base64_pdf}" download="{fse_first_name} {fse_first_name} dashboard_report.pdf">Download Report as PDF</a>'
        return href

    # Add a button in the sidebar to generate and download the PDF
    st.sidebar.divider()
    if st.sidebar.button("Create PDF of Dashboard"):
        pdf_path = create_pdf_with_graphs(table_display_df)
        st.sidebar.markdown(get_pdf_download_link(pdf_path), unsafe_allow_html=True)

st.divider()