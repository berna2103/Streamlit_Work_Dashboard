import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import os
import random
from io import BytesIO

# Import pptx libraries
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN # Import PP_ALIGN for text alignment

# --- Define Color Scheme ---
PRIMARY_COLOR = 'rgb(43, 101, 124)'
SECONDARY_COLOR = 'rgb(54, 164, 179)'
ELEKTA_FONT_COLOR = RGBColor(43,101,125)
RGB_CUSTOM_COLORS = [
    'rgb(43,101,125)',  # Base color (teal-blue)
    'rgb(85,130,145)',  # Lighter and more saturated
    'rgb(25,85,105)',   # Darker and less saturated
    'rgb(135,175,195)', # Much lighter
    'rgb(0,60,80)',     # Very dark teal
    'rgb(58,121,150)',  # Brighter, with more blue
    'rgb(90,140,160)',  # Muted teal-blue
    'rgb(10,70,90)',    # Dark, almost navy
    'rgb(100,180,200)', # Light cyan-blue
    'rgb(30,90,110)'    # Deep teal
]
COLOR_SEQUENCE = RGB_CUSTOM_COLORS

# Set title and configure Streamlit page layout
TITLE = 'Service Agreement Report'
DOWNTIME_URL = 'https://elekta.lightning.force.com/lightning/r/Report/00OKf000000Z339MAC/view?queryScope=userFolders'

# --- Helper functions for PPTX ---
def add_custom_textbox(slide, left:Inches, top:Inches, width: Inches, height: Inches, font_name: str, font_size:Pt, font_color: RGBColor, bold: bool, text: str, text_align: PP_ALIGN = PP_ALIGN.LEFT):
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.text = text
    p = text_frame.paragraphs[0]
    p.font.name = font_name
    p.font.size = font_size
    p.font.bold = bold
    p.font.color.rgb = font_color
    p.alignment = text_align

def add_formatted_text_line(slide, left:Inches, top:Inches, width: Inches, height: Inches, font_name: str, font_size:Pt, font_color: RGBColor, text_parts: list[tuple[str, bool]], text_align: PP_ALIGN = PP_ALIGN.LEFT):
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.paragraphs[0]
    p.alignment = text_align
    for text, bold_status in text_parts:
        run = p.add_run()
        run.text = text
        font = run.font
        font.name = font_name
        font.size = font_size
        font.bold = bold_status
        font.color.rgb = font_color

def add_rectangle_background(slide, left:Inches, top:Inches, width:Inches, height:Inches, BGcolor:RGBColor, border:int):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = BGcolor
    shape.shadow.inherit = False
    line = shape.line
    if border == 0:
        line.fill.background()
    else:
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 0, 0)
        line.width = Pt(border)

def add_table_to_slide(slide, df_table, left, top, width, height, font_name, font_size):
    rows, cols = df_table.shape
    table = slide.shapes.add_table(rows + 1, cols, left, top, width, height).table
    for i in range(cols):
        table.columns[i].width = int(width / cols)
    for col_idx, col_name in enumerate(df_table.columns):
        cell = table.cell(0, col_idx)
        text_frame = cell.text_frame
        text_frame.text = str(col_name)
        p = text_frame.paragraphs[0]
        p.font.name = font_name
        p.font.size = Pt(font_size)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(230, 230, 230)
    for row_idx, row_data in df_table.iterrows():
        for col_idx, cell_value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            text_frame = cell.text_frame
            text_frame.text = str(cell_value)
            p = text_frame.paragraphs[0]
            p.font.name = font_name
            p.font.size = Pt(font_size)
            p.alignment = PP_ALIGN.CENTER

def get_sanitized_image_path(product_name_str: str, base_dir: str = "images/Cards") -> str:
    """
    Generates a sanitized, consistent image filename from a product name string.
    Example: 'Versa HD / 123' -> 'images/Cards/Versa_HD.png'
    """
    if not isinstance(product_name_str, str) or not product_name_str:
        return "images/Cards/Unknown.png" # Return a path to a default image
        
    first_part = product_name_str.split('/')[0].strip()
    # Replace non-alphanumeric characters with underscores, then clean up any extra underscores
    sanitized_name = "".join(c if c.isalnum() else '_' for c in first_part)
    sanitized_name = sanitized_name.replace('__', '_').strip('_')
    
    image_filename = f"{sanitized_name}.png"
    return os.path.join(base_dir, image_filename)

def generate_service_contract_slides(df_data: pd.DataFrame, ppt_title: str):
    """
    Generates a PowerPoint presentation from the dataframe and returns it as an in-memory BytesIO buffer.
    Graphs are converted to images in memory to avoid writing to disk.
    """
    prs = Presentation()
    prs.slide_width = Inches(26.66)
    prs.slide_height = Inches(15)
    font_name = 'Calibri'
    image_folder_ = './images/'
    image_folder = './images/Cards'
    slide_layout = prs.slide_layouts[6]

    # --- 1. Title Slide ---
    slide = prs.slides.add_slide(slide_layout)
    try:
        if os.path.exists(image_folder_) and os.listdir(image_folder_):
            images_in_folder = [f for f in os.listdir(image_folder_) if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif'))]
            if images_in_folder:
                random_image = random.choice(images_in_folder)
                image_path = os.path.join(image_folder_, random_image)
                slide.shapes.add_picture(image_path, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    except Exception as e:
        st.error(f"Error adding background image to title slide: {e}")

    add_custom_textbox(slide=slide, left=Inches(1), top=Inches(5.47), width=Inches(11), height=Inches(3),
                       font_name=font_name, font_color=ELEKTA_FONT_COLOR, font_size=Pt(90), bold=True, text=ppt_title)

    # --- 2. KPI Summary Slide ---
    slide = prs.slides.add_slide(slide_layout)
    add_rectangle_background(slide, Inches(0.81), Inches(2.7), Inches(24.75), Inches(11.86), RGBColor(248,248,248), 0)
    add_custom_textbox(slide, Inches(0.8), Inches(1.08), Inches(24), Inches(2.9), font_name, Pt(80), ELEKTA_FONT_COLOR, True, "Service Contract Key Metrics")

    total_devices = df_data.shape[0]
    contracts_expiring_soon = df_data[df_data['Contract Status'] == 'Expiring Soon'].shape[0]
    expired_contracts = df_data[df_data['Contract Status'] == 'Expired'].shape[0]
    total_contract_value = df_data['Contract Price'].sum() if 'Contract Price' in df_data.columns else 0

    add_custom_textbox(slide, Inches(2), Inches(4), Inches(5), Inches(1), font_name, Pt(40), ELEKTA_FONT_COLOR, False, "Total Devices Installed")
    add_custom_textbox(slide, Inches(3), Inches(4.7), Inches(5), Inches(1), font_name, Pt(100), ELEKTA_FONT_COLOR, True, f"{total_devices}")
    add_custom_textbox(slide, Inches(9), Inches(4), Inches(5), Inches(1), font_name, Pt(40), ELEKTA_FONT_COLOR, False, "Contracts Expiring (<12 weeks)")
    add_custom_textbox(slide, Inches(10), Inches(4.7), Inches(5), Inches(1), font_name, Pt(100), ELEKTA_FONT_COLOR, True, f"{contracts_expiring_soon}")
    add_custom_textbox(slide, Inches(16), Inches(4), Inches(5), Inches(1), font_name, Pt(40), ELEKTA_FONT_COLOR, False, "Expired Contracts")
    add_custom_textbox(slide, Inches(17), Inches(4.7), Inches(5), Inches(1), font_name, Pt(100), ELEKTA_FONT_COLOR, True, f"{expired_contracts}")
    add_custom_textbox(slide, Inches(2), Inches(9), Inches(5), Inches(1), font_name, Pt(60), ELEKTA_FONT_COLOR, False, "Total Contract Annual Value")
    add_custom_textbox(slide, Inches(2), Inches(10.2), Inches(5), Inches(1), font_name, Pt(120), ELEKTA_FONT_COLOR, True, f"${total_contract_value:,.0f}")

    # --- 3. Contract Status Distribution Slide ---
    slide = prs.slides.add_slide(slide_layout)
    add_rectangle_background(slide, Inches(0.81), Inches(2.7), Inches(24.75), Inches(11.86), RGBColor(248,248,248), 0)
    add_custom_textbox(slide, Inches(1.27), Inches(1.08), Inches(24), Inches(2.9), font_name, Pt(70), ELEKTA_FONT_COLOR, True, "Contract Status Distribution")
    contract_status_counts = df_data['Contract Status'].value_counts().rename_axis('Status').reset_index(name='Count')
    fig_contract_status = px.pie(contract_status_counts, values='Count', names='Status', title='Overall Contract Status Distribution', color_discrete_sequence=COLOR_SEQUENCE, template="plotly_white")
    img_bytes = fig_contract_status.to_image(format="png", width=1200, height=800)
    slide.shapes.add_picture(BytesIO(img_bytes), Inches(5), Inches(3.5), width=Inches(16))

    # --- 4. Device Lifecycle & Risk (Age Distribution) Slide ---
    if 'Device Age Group' in df_data.columns:
        slide = prs.slides.add_slide(slide_layout)
        add_rectangle_background(slide, Inches(0.81), Inches(2.7), Inches(24.75), Inches(11.86), RGBColor(248,248,248), 0)
        add_custom_textbox(slide, Inches(1.27), Inches(1.08), Inches(24), Inches(2.9), font_name, Pt(70), ELEKTA_FONT_COLOR, True, "Device Lifecycle & Risk")
        fig_age = px.histogram(df_data, x='Device Age Group', title='Distribution of Device Ages', labels={'Device Age Group': 'Device Age (Years)'}, category_orders={"Device Age Group": ["0-5 years", "5-10 years", ">10 years"]}, color_discrete_sequence=COLOR_SEQUENCE, template="plotly_white")
        fig_age.update_layout(bargap=0.8, showlegend=True)
        img_bytes = fig_age.to_image(format="png", width=1800, height=900)
        slide.shapes.add_picture(BytesIO(img_bytes), Inches(2), Inches(4), width=Inches(20))

    # --- 5. Upcoming Renewals & Expirations Slide ---
    if 'Weeks To Renewal' in df_data.columns:
        slide = prs.slides.add_slide(slide_layout)
        add_rectangle_background(slide, Inches(0.81), Inches(2.7), Inches(24.75), Inches(11.86), RGBColor(248,248,248), 0)
        add_custom_textbox(slide, Inches(0.8), Inches(1.08), Inches(24), Inches(2.9), font_name, Pt(80), ELEKTA_FONT_COLOR, True, "Upcoming Renewals & Expirations")
        upcoming_renewals = df_data[df_data['Weeks To Renewal'] <= 52].copy()
        if not upcoming_renewals.empty:
            upcoming_renewals['Renewal Period'] = pd.cut(upcoming_renewals['Weeks To Renewal'], bins=[-0.1, 0, 12, 26, 52], labels=['Expired', '0-12 Weeks', '12-26 Weeks', '26-52 Weeks'], right=True, include_lowest=True)
            renewal_counts = upcoming_renewals.groupby(['Renewal Period','Location']).size().unstack(fill_value=0)
            fig_renewals = px.bar(renewal_counts, x=renewal_counts.index, y=renewal_counts.columns, title='Number of Contracts by Upcoming Renewal Period', labels={'value': 'Number of Contracts', 'Location': 'Location'}, color_discrete_sequence=COLOR_SEQUENCE, template="plotly_white")
            fig_renewals.update_layout(barmode='stack')
            img_bytes = fig_renewals.to_image(format="png", width=1800, height=900)
            slide.shapes.add_picture(BytesIO(img_bytes), Inches(2), Inches(4), width=Inches(20))
        else:
            add_custom_textbox(slide, Inches(2), Inches(5), Inches(20), Inches(2), font_name, Pt(30), RGBColor(100,100,100), False, "No upcoming renewals in the next 52 weeks.", text_align=PP_ALIGN.CENTER)

    # --- 6. Financial Summary Slide ---
    if 'Contract Price' in df_data.columns:
        slide = prs.slides.add_slide(slide_layout)
        add_rectangle_background(slide, Inches(0.81), Inches(2.7), Inches(24.75), Inches(11.86), RGBColor(248,248,248), 0)
        add_custom_textbox(slide, Inches(1.27), Inches(1.08), Inches(24), Inches(2.9), font_name, Pt(70), ELEKTA_FONT_COLOR, True, "Contract Value by Location")
        contract_value_by_location = df_data.groupby('Location')['Contract Price'].sum().reset_index().sort_values(by='Contract Price', ascending=False)
        fig_financial = px.bar(contract_value_by_location, y='Location', x='Contract Price', title='Total Contract Value by Location', labels={'Contract Price': 'Contract Value'}, color='Location', color_discrete_sequence=COLOR_SEQUENCE, template="plotly_white")
        fig_financial.update_layout(showlegend=False)
        img_bytes = fig_financial.to_image(format="png", width=1800, height=900)
        slide.shapes.add_picture(BytesIO(img_bytes), Inches(2), Inches(4), width=Inches(20))

    # --- 7. Individual Device Cards Slides for PowerPoint ---
    card_width_ppt, card_height_ppt = Inches(7.5), Inches(9)
    card_starts_x_ppt = [Inches(1), Inches(9.5), Inches(18)]
    card_start_y_ppt = Inches(4) 

    for i in range(0, len(df_data), 3):
        slide = prs.slides.add_slide(slide_layout)
        add_rectangle_background(slide, Inches(0.3), Inches(2.7), Inches(26.00), Inches(11.86), RGBColor(248,248,248), 0)
        add_custom_textbox(slide, Inches(0.8), Inches(1.08), Inches(24), Inches(1.5), font_name, Pt(80), ELEKTA_FONT_COLOR, True, "Machine Fleet Overview")
        devices_on_this_slide = df_data.iloc[i : i + 3]
        for j, (idx, device) in enumerate(devices_on_this_slide.iterrows()):
            current_card_left_ppt = card_starts_x_ppt[j]
            add_rectangle_background(slide, current_card_left_ppt, card_start_y_ppt, card_width_ppt, card_height_ppt, RGBColor(255,255,255), 1)
            
            # --- Image Handling ---
            full_image_path_on_disk = get_sanitized_image_path(device.get('Installed Product'))
            img_width_card_ppt = Inches(3)
            img_left_card_ppt = current_card_left_ppt + (card_width_ppt - img_width_card_ppt) / 2
            img_top_card_ppt = card_start_y_ppt + Inches(0.5)
            if os.path.exists(full_image_path_on_disk):
                slide.shapes.add_picture(full_image_path_on_disk, img_left_card_ppt, img_top_card_ppt, width=img_width_card_ppt)
            else:
                add_custom_textbox(slide, img_left_card_ppt, img_top_card_ppt, img_width_card_ppt, Inches(1), font_name, Pt(10), RGBColor(150,150,150), False, "Image N/A", text_align=PP_ALIGN.CENTER)

            # --- Text Content ---
            add_custom_textbox(slide, left=current_card_left_ppt + Inches(0.5), top=img_top_card_ppt + img_width_card_ppt + Inches(0.2), width=card_width_ppt - Inches(1), height=Inches(0.8), font_name=font_name, font_size=Pt(40), font_color=ELEKTA_FONT_COLOR, bold=True, text=f"{device.get('Display Product Name', 'N/A')}", text_align=PP_ALIGN.CENTER)
            
            detail_start_top_ppt = img_top_card_ppt + img_width_card_ppt + Inches(1.2)
            line_height_ppt = Inches(0.5)
            text_box_left = current_card_left_ppt + Inches(0.5)
            text_box_width = card_width_ppt - Inches(1)

            end_date_str = device.get('Contract End Date', pd.NaT).strftime('%m/%d/%Y') if pd.notna(device.get('Contract End Date')) else "N/A"
            add_formatted_text_line(slide, text_box_left, detail_start_top_ppt, text_box_width, line_height_ppt, font_name, Pt(25), RGBColor(50,50,50), [("Contract Expires: ", True), (end_date_str, False)], text_align=PP_ALIGN.CENTER)
            add_formatted_text_line(slide, text_box_left, detail_start_top_ppt + line_height_ppt, text_box_width, line_height_ppt, font_name, Pt(25), RGBColor(50,50,50), [("Age: ", True), (f"{device.get('Device Age', 'N/A')} years", False)], text_align=PP_ALIGN.CENTER)
            add_formatted_text_line(slide, text_box_left, detail_start_top_ppt + 2*line_height_ppt, text_box_width, line_height_ppt, font_name, Pt(25), RGBColor(50,50,50), [("Renew In: ", True), (f"{device.get('Weeks To Renewal', 'N/A')} weeks", False)], text_align=PP_ALIGN.CENTER)
            
            # --- MODIFIED: Draw a thin rectangle to act as a line ---
            line_top_ppt = detail_start_top_ppt + 3 * line_height_ppt + Inches(0.4)
            line_shape = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, text_box_left, line_top_ppt, text_box_width, Pt(1)
            )
            line_fill = line_shape.fill
            line_fill.solid()
            line_fill.fore_color.rgb = RGBColor(200, 200, 200)
            line_shape.line.fill.background() # No border for the rectangle

            customs_acceptance_date_str = device.get('Customs Acceptance Date', pd.NaT).strftime('%m/%d/%Y') if pd.notna(device.get('Customs Acceptance Date')) else "N/A"
            warranty_end_str = device.get('Warranty End Date', pd.NaT).strftime('%m/%d/%Y') if pd.notna(device.get('Warranty End Date')) else 'N/A'
            eol_date_str = device.get('EoL Date IP', pd.NaT).strftime('%m/%d/%Y') if pd.notna(device.get('EoL Date IP')) else 'N/A'
            
            add_formatted_text_line(slide, text_box_left, detail_start_top_ppt + 4*line_height_ppt + Inches(0.3), text_box_width, line_height_ppt, font_name, Pt(20), RGBColor(100,100,100), [("CAT: ", True), (customs_acceptance_date_str, False)], text_align=PP_ALIGN.CENTER)
            add_formatted_text_line(slide, text_box_left, detail_start_top_ppt + 5*line_height_ppt + Inches(0.3), text_box_width, line_height_ppt, font_name, Pt(20), RGBColor(100,100,100), [("Warranty End Date: ", True), (warranty_end_str, False)], text_align=PP_ALIGN.CENTER)
            add_formatted_text_line(slide, text_box_left, detail_start_top_ppt + 6*line_height_ppt + Inches(0.3), text_box_width, line_height_ppt, font_name, Pt(20), RGBColor(100,100,100), [("EoL Date IP: ", True), (eol_date_str, False)], text_align=PP_ALIGN.CENTER)

    # Save the presentation to an in-memory buffer
    ppt_buffer = BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

# --- Page Configuration ---
st.set_page_config(page_title="Service Contracts Dashboard", page_icon="🏥", layout="wide", initial_sidebar_state="expanded")

# --- Custom CSS ---
st.markdown(f"""
    <style>
    .main {{ background-color: #f0f2f6; }}
    .stButton>button {{ background-color: {PRIMARY_COLOR}; color: white; border-radius: 8px; border: none; }}
    .stButton>button:hover {{ background-color: {SECONDARY_COLOR}; }}
    .st-emotion-cache-1r6dm1s {{ background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05); padding: 15px; }}
    </style>
""", unsafe_allow_html=True)

# --- Title and Description ---
st.markdown(f"<h1 style='color: {PRIMARY_COLOR};'>🏥 Service Contracts Dashboard</h1>", unsafe_allow_html=True)
st.markdown("Upload your CSV/Excel file to visualize key insights about your equipment service contracts.")

# --- File Uploader ---
st.sidebar.title('Data Upload & Filters')
if 'uploaded_df' not in st.session_state:
    uploaded_file = st.sidebar.file_uploader("Upload your data file", type=["csv", "xlsx", "xls"])
else:
    st.sidebar.success("Data already loaded. Upload a new file to replace.")
    uploaded_file = st.sidebar.file_uploader("Upload new data file", type=["csv", "xlsx", "xls"])
st.sidebar.divider()

df = None
if uploaded_file is not None:
    try:
        file_extension = os.path.splitext(uploaded_file.name)[1]
        df = pd.read_excel(uploaded_file) if file_extension in ['.xlsx', '.xls'] else pd.read_csv(uploaded_file, encoding='utf-8', on_bad_lines='skip')
        st.session_state['uploaded_df'] = df.copy()
        st.sidebar.success("File uploaded successfully!")
    except Exception as e:
        st.error(f"Error processing file: {e}.")
        if 'uploaded_df' in st.session_state: del st.session_state['uploaded_df']

if 'uploaded_df' in st.session_state and not st.session_state['uploaded_df'].empty:
    df = st.session_state['uploaded_df'].copy()

    # --- Data Preprocessing ---
    rename_map = {
        'Installed Product: Installed Product': 'Installed Product', 'Installed Product: Serial/Lot Number': 'Serial Number',
        'Serial/Lot Number': 'Serial Number', 'Installed Product: Warranty End Date': 'Warranty End Date',
        'Installed Product: EoL Date IP': 'EoL Date IP', 'Installed Product: EoGS Date IP': 'EoGS Date IP',
        'Installed Product: Device Age': 'Device Age', 'Installed Product: Customer/Device Acceptance Date': 'Customs Acceptance Date',
        'Service/Maintenance Contract: Contract Name/Number': 'Contract Name/Number', 'Covered Product: Record Number': 'Covered Product Record Number',
        'Current Term Start Date': 'Contract Start Date', 'Current Term End Date': 'Contract End Date'
    }
    df.rename(columns=rename_map, inplace=True)
    df['Display Product Name'] = df['Installed Product'].apply(lambda x: f"{x.split('/')[0]} {x.split('/')[-1]}" if isinstance(x, str) and len(x.split('/')) >= 3 else x) if 'Installed Product' in df.columns else 'N/A'
    date_cols = ['Warranty End Date', 'EoL Date IP', 'EoGS Date IP', 'End Date', 'Start Date', 'Customs Acceptance Date', 'Contract Start Date', 'Contract End Date']
    for col in date_cols:
        if col in df.columns: df[col] = pd.to_datetime(df[col], errors='coerce')
    if 'Device Age' not in df.columns and 'Customs Acceptance Date' in df.columns: df['Device Age'] = ((datetime.now() - df['Customs Acceptance Date']).dt.days / 365.25).round(1)
    if 'Device Age' in df.columns:
        df['Device Age Group'] = pd.cut(df['Device Age'], bins=[0, 5, 10, float('inf')], labels=['0-5 years', '5-10 years', '>10 years'], right=False)
        df['Device Age Group'] = df['Device Age Group'].cat.add_categories('N/A').fillna('N/A')
    if 'Contract Price' in df.columns: df['Contract Price'] = pd.to_numeric(df['Contract Price'], errors='coerce').fillna(0)
    if 'Weeks To Renewal' not in df.columns and 'Contract End Date' in df.columns: df['Weeks To Renewal'] = ((df['Contract End Date'] - datetime.now()).dt.days / 7).apply(lambda x: max(0, x)).round(0)
    elif 'Weeks To Renewal' in df.columns: df['Weeks To Renewal'] = pd.to_numeric(df['Weeks To Renewal'], errors='coerce').fillna(9999)
    else: df['Weeks To Renewal'] = 9999
    df['Contract Status'] = df.apply(lambda row: 'Expired' if row['Weeks To Renewal'] <= 0 else ('Expiring Soon' if row['Weeks To Renewal'] <= 12 else 'Active'), axis=1)

    # --- Sidebar Filters ---
    st.sidebar.header("Filters")
    all_accounts = ['All'] + sorted(df['Account'].unique().tolist())
    selected_accounts = st.sidebar.multiselect("Filter by Account", all_accounts, default='All')
    filtered_df = df if 'All' in selected_accounts else df[df['Account'].isin(selected_accounts)]
    all_statuses = ['All'] + sorted(filtered_df['Contract Status'].unique().tolist())
    selected_statuses = st.sidebar.multiselect("Filter by Contract Status", all_statuses, default='All')
    if 'All' not in selected_statuses: filtered_df = filtered_df[filtered_df['Contract Status'].isin(selected_statuses)]
    max_weeks = int(filtered_df['Weeks To Renewal'].max()) if not filtered_df.empty else 104
    weeks_filter = st.sidebar.slider("Contracts expiring in next X weeks", 0, max_weeks, max_weeks)
    df_display = filtered_df[filtered_df['Weeks To Renewal'] <= weeks_filter].copy()

    if df_display.empty:
        st.warning("No data matches the selected filters.")
    else:
        # --- Dashboard Sections ---
        st.header("Overall Fleet & Contract Summary")
        with st.container(border=True):
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Devices Installed", df_display.shape[0])
            col2.metric("Contracts Expiring Soon (<12 weeks)", df_display[df_display['Contract Status'] == 'Expiring Soon'].shape[0])
            col3.metric("Expired Contracts", df_display[df_display['Contract Status'] == 'Expired'].shape[0])
            col4.metric("Total Contract Value", f"${df_display['Contract Price'].sum():,.0f}")

        st.markdown("---")
        st.header("Analytics")
        tab1, tab2, tab3, tab4 = st.tabs(["Contract Status", "Device Age", "Upcoming Renewals", "Financials"])
        with tab1:
            # --- MODIFIED: Robustly create the dataframe for the pie chart ---
            status_counts = df_display['Contract Status'].value_counts().rename_axis('Status').reset_index(name='Count')
            fig = px.pie(status_counts, values='Count', names='Status', title='Contract Status Distribution', color_discrete_sequence=COLOR_SEQUENCE)
            st.plotly_chart(fig, use_container_width=True)
        with tab2:
            fig = px.histogram(df_display, x='Device Age Group', title='Distribution of Device Ages', color_discrete_sequence=COLOR_SEQUENCE, category_orders={"Device Age Group": ["0-5 years", "5-10 years", ">10 years", "N/A"]})
            st.plotly_chart(fig, use_container_width=True)
        with tab3:
            renewals = df_display[df_display['Weeks To Renewal'] <= 52].copy()
            if not renewals.empty:
                renewals['Renewal Period'] = pd.cut(renewals['Weeks To Renewal'], bins=[-0.1, 0, 12, 26, 52], labels=['Expired', '0-12 Weeks', '12-26 Weeks', '26-52 Weeks'], right=True)
                renewal_counts = renewals.groupby(['Renewal Period','Installed Product']).size().unstack(fill_value=0)
                fig = px.bar(renewal_counts, x=renewal_counts.index, y=renewal_counts.columns, title='Contracts by Upcoming Renewal Period', color_discrete_sequence=COLOR_SEQUENCE)
                fig.update_layout(barmode='stack')
                st.plotly_chart(fig, use_container_width=True)
            else: st.info("No upcoming renewals in the next 52 weeks.")
        with tab4:
            value_by_loc = df_display.groupby('Location')['Contract Price'].sum().reset_index().sort_values('Contract Price', ascending=False)
            fig = px.bar(value_by_loc, y='Location', x='Contract Price', title='Total Contract Value by Location', color='Location', color_discrete_sequence=COLOR_SEQUENCE)
            fig.update_layout(showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.header("Devices At Risk (Expired or Expiring Soon)")
        with st.container(border=True):
            at_risk = df_display[df_display['Contract Status'].isin(['Expired', 'Expiring Soon'])].copy()
            if not at_risk.empty:
                display_cols = ['Account', 'Location', 'Display Product Name', 'Serial Number', 'Contract Status', 'Weeks To Renewal', 'Warranty End Date', 'EoL Date IP']
                at_risk_display = at_risk[display_cols]
                for col in ['Warranty End Date', 'EoL Date IP']: at_risk_display[col] = at_risk_display[col].dt.strftime('%m/%d/%Y').fillna('N/A')
                st.dataframe(at_risk_display, use_container_width=True)
            else: st.info("No devices currently at risk based on filters.")

        st.markdown("---")
        st.header("Machine Fleet Overview")
        # NOTE: For images to appear, create a folder named 'images/Cards' next to your script.
        # Inside, place PNG files named after the sanitized product name, e.g., 'Versa_HD.png'.
        cols_per_row = 3
        for i in range(0, len(df_display), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, (idx, device) in enumerate(df_display.iloc[i:i+cols_per_row].iterrows()):
                with cols[j]:
                    with st.container(border=True):
                        st.markdown(f"<h3 style='text-align: center; color: {PRIMARY_COLOR};'>{device.get('Display Product Name', 'N/A')}</h3>", unsafe_allow_html=True)
                        img_path = get_sanitized_image_path(device.get('Installed Product'))
                        st.image(img_path if os.path.exists(img_path) else "https://placehold.co/150x150/e0e0e0/A0A0A0?text=No+Image", use_column_width='auto')
                        st.markdown(f"""
                        <div style="text-align: center;">
                            <p><strong>Contract Expires:</strong> {device.get('Contract End Date', pd.NaT).strftime('%m/%d/%Y') if pd.notna(device.get('Contract End Date')) else 'N/A'}</p>
                            <p><strong>Renew In:</strong> {device.get('Weeks To Renewal', 'N/A')} weeks</p>
                            <p><strong>Age:</strong> {device.get('Device Age', 'N/A')} years</p>
                            <hr>
                            <p><small><strong>Warranty End:</strong> {device.get('Warranty End Date', pd.NaT).strftime('%m/%d/%Y') if pd.notna(device.get('Warranty End Date')) else 'N/A'}</small></p>
                            <p><small><strong>EoL IP:</strong> {device.get('EoL Date IP', pd.NaT).strftime('%m/%d/%Y') if pd.notna(device.get('EoL Date IP')) else 'N/A'}</small></p>
                        </div>
                        """, unsafe_allow_html=True)
    
    st.sidebar.title('PowerPoint Export')
    if not df_display.empty:
        if st.sidebar.button('Generate Service Contract Slides'):
            with st.spinner('Generating PowerPoint... This may take a moment.'):
                try:
                    # The function now returns an in-memory buffer
                    ppt_buffer = generate_service_contract_slides(df_display, "Radiation Therapy Service Contracts Dashboard")
                    
                    # Define the filename for the download
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    output_ppt_file = f'Service_Contracts_Dashboard_{timestamp}.pptx'

                    # The download button now uses the buffer directly
                    st.sidebar.download_button(
                        label="Download PowerPoint",
                        data=ppt_buffer,
                        file_name=output_ppt_file,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    st.sidebar.success(f"Presentation is ready for download!")
                except Exception as e:
                    st.sidebar.error(f"Error generating PowerPoint: {e}")
                    st.sidebar.info("Please ensure 'kaleido' is installed (`pip install kaleido`) for graph export.")
    else:
        st.sidebar.info("Upload data and apply filters to enable PowerPoint generation.")
else:
    st.info("Please upload a Service Agreement Report to begin.")
    st.markdown(f"1. Go to Reports on CLM and select **Service Agreement Report**. [Click here to open]({DOWNTIME_URL})")
    st.markdown("2. Export the report as **Details Only** in `.csv` or `.xlsx` format.")
    st.markdown("3. Upload the file using the sidebar.")
