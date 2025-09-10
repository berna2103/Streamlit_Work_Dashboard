import streamlit as st
import pandas as pd
import numpy as np

# --- Page Configuration ---
st.set_page_config(
    page_title="Service Operations Dashboard",
    page_icon="🛠️",
    layout="wide"
)

# --- Helper Functions ---

@st.cache_data
def convert_df_to_csv(df):
    """Converts a dataframe to a CSV file for downloading."""
    # IMPORTANT: Cache the conversion to prevent computation on every rerun
    return df.to_csv(index=False).encode('utf-8')

def _deduplicate_columns(columns):
    """Helper function to rename duplicate column names with a suffix."""
    seen = {}
    new_columns = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_columns.append(col)
    return new_columns

def load_data(uploaded_file):
    """Loads data from the uploaded file and performs initial cleaning based on new column names."""
    if uploaded_file is None:
        return None
    try:
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)

        # --- Data Cleaning and Preprocessing ---
        df.columns = df.columns.str.lower().str.replace(':', '').str.replace(' ', '_').str.replace('/', '_')
        df.columns = _deduplicate_columns(df.columns)

        # Filter out rows where 'Consumed From Location' is blank or a hyphen
        if 'consumed_from_location_location_name' in df.columns:
            # Convert column to string to handle various blank types (NaN, None) and strip whitespace
            df['consumed_from_location_location_name'] = df['consumed_from_location_location_name'].astype(str).str.strip()
            # Keep rows where the location is not 'nan', an empty string, or a hyphen
            df = df[~df['consumed_from_location_location_name'].isin(['nan', '', '-'])]

        # Define the new date columns
        date_cols = [
            'start_date_and_time', 'end_date_and_time', 'date_time_opened', 
            'date_time_closed', 'first_assigned_datetime', 
            'acknowledged_by_technician_date_time'
        ]
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        return df
    except Exception as e:
        st.error(f"Error loading or processing file: {e}")
        return None

def calculate_case_times(df):
    """
    Calculates resolution time and response time for each case that has labor.
    Returns a single DataFrame with all time metrics for a consistent set of cases.
    """
    if df is None or df.empty or 'case_number' not in df.columns or 'line_type' not in df.columns:
        return pd.DataFrame()

    # 1. Identify all cases that have at least one 'labor' line type.
    cases_with_labor = df[df['line_type'].str.lower() == 'labor']['case_number'].unique()
    labor_case_data = df[df['case_number'].isin(cases_with_labor)].copy()

    if labor_case_data.empty:
        return pd.DataFrame()

    # 2. Calculate Resolution Time (MTTR) for these cases
    mttr_df = labor_case_data.groupby('case_number').agg(
        first_start=('start_date_and_time', 'min'),
        last_end=('end_date_and_time', 'max')
    ).reset_index()
    mttr_df['resolution_time'] = mttr_df['last_end'] - mttr_df['first_start']
    
    # 3. Calculate Response Time (TTR) for these cases
    labor_lines = labor_case_data[labor_case_data['line_type'].str.lower() == 'labor']
    ttr_df = labor_lines.groupby('case_number').agg(
        first_labor_start=('start_date_and_time', 'min'),
        case_opened=('date_time_opened', 'first')
    ).reset_index()
    ttr_df['response_time'] = ttr_df['first_labor_start'] - ttr_df['case_opened']

    # 4. Merge them into a single DataFrame
    case_times_df = pd.merge(mttr_df, ttr_df, on='case_number', how='inner')
    
    # Filter out negative durations
    case_times_df = case_times_df[
        (case_times_df['resolution_time'] > pd.Timedelta(0)) & 
        (case_times_df['response_time'] > pd.Timedelta(0))
    ]

    return case_times_df


def calculate_ftfr(df, repeat_visit_days=30):
    """
    Calculates a proxy for First-Time Fix Rate (FTFR).
    Logic: A work order is a "repeat visit" if another work order is created
    for the same Case within the specified number of days of the previous one.
    """
    if df is None or df.empty or 'case_number' not in df.columns:
        return 0, 0, pd.DataFrame()

    wo_df = df[['case_number', 'work_order_number', 'date_time_opened']].drop_duplicates()
    wo_df = wo_df.sort_values(by=['case_number', 'date_time_opened'])
    wo_df['time_since_last_wo_in_case'] = wo_df.groupby('case_number')['date_time_opened'].diff()
    wo_df['is_repeat'] = wo_df['time_since_last_wo_in_case'] <= pd.Timedelta(days=repeat_visit_days)

    total_work_orders = wo_df['work_order_number'].nunique()
    repeat_visits = wo_df['is_repeat'].sum()
    
    if total_work_orders == 0:
        return 0, 0, pd.DataFrame()

    ftfr = (total_work_orders - repeat_visits) / total_work_orders
    
    # Safely merge columns that exist in the dataframe to prevent KeyErrors.
    merge_cols = ['work_order_number', 'owner_full_name', 'installed_product_installed_product']
    if 'installed_product_serial_number' in df.columns:
        merge_cols.append('installed_product_serial_number')
    
    wo_df = wo_df.merge(df[merge_cols].drop_duplicates(), on='work_order_number', how='left')
    
    return ftfr, repeat_visits, wo_df

def calculate_delayed_starts(df):
    """
    Calculates delay from the case open time to the first work start time.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    # Use 'acknowledged_by_technician_date_time' if available, otherwise fall back to 'start_date_and_time'
    start_col = 'acknowledged_by_technician_date_time' if 'acknowledged_by_technician_date_time' in df.columns else 'start_date_and_time'

    wo_start_info = df.groupby('work_order_number').agg(
        first_start=(start_col, 'min'),
        creation_date=('date_time_opened', 'first')
    ).reset_index()

    wo_start_info['start_delay'] = wo_start_info['first_start'] - wo_start_info['creation_date']
    
    return wo_start_info

def remove_iqr_outliers(df, column):
    """Removes outliers from a dataframe based on the IQR method."""
    if df.empty or column not in df.columns:
        return df, 0, None, None, None
    
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1

    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR

    initial_count = len(df)
    filtered_df = df[(df[column] >= lower_bound) & (df[column] <= upper_bound)]
    outlier_count = initial_count - len(filtered_df)

    return filtered_df, outlier_count, Q1, Q3, IQR

# --- Main Application UI ---

st.title("🛠️ Service Operations Analysis Dashboard")
st.markdown("""
Welcome! This dashboard analyzes field service reports to provide insights into key operational metrics. 
Please upload your report file (CSV or Excel) and use the sidebar to adjust filters.
""")

# --- Sidebar for File Upload and Filters ---
with st.sidebar:
    st.header("Controls")
    uploaded_file = st.file_uploader("Upload your service report", type=["csv", "xlsx"], key="service_dashboard_uploader")
    
    # Placeholder for the origin filter, will be populated after data load
    origin_filter_placeholder = st.empty()

    st.markdown("---")
    st.header("Analysis Filters")
    
    max_start_delay_days = st.slider(
        "Filter Outlier Starts (Days)", 
        min_value=1, 
        max_value=30, 
        value=7, 
        help="Exclude work orders that started more than this many days after the case was opened."
    )
    
    max_mttr_days = st.slider(
        "Filter Outlier Resolution Times (Days)",
        min_value=1,
        max_value=90,
        value=30,
        help="Exclude cases that took longer than this many days to resolve from the final MTTR calculation."
    )

    repeat_visit_days = st.slider(
        "Repeat Visit Window (Days)", 
        min_value=1, 
        max_value=60, 
        value=30,
        help="A visit is a 'repeat' if another work order is created for the same Case within this window."
    )
    
    st.markdown("---")
    st.header("Display Options")
    remove_mttr_iqr_outliers = st.checkbox(
        "Remove MTTR Outliers (IQR Method)", 
        False, 
        help="Refine MTTR calculation by removing statistical outliers using the Interquartile Range (IQR) method."
    )
    remove_ttr_iqr_outliers = st.checkbox(
        "Remove TTR Outliers (IQR Method)", 
        False, 
        help="Refine Time to Respond calculation by removing statistical outliers using the Interquartile Range (IQR) method."
    )
    show_data_tables = st.checkbox("Show Data Tables", False, help="Toggle the visibility of the data tables below the charts in each tab.")

    st.markdown("---")
    st.info("""
    **Metric Definitions:**
    - **MTTR:** Average time from the first start time to the last end time for a single **Case**.
    - **Time to Respond:** Average time from case creation to the start of the first **Labor** activity.
    - **FTFR (Proxy):** A work order is a 'repeat' if another is created for the same **Case** within the defined window.
    - **Delayed Starts:** Work orders where the first activity started more than 24 hours after the case was opened.
    """)

# --- Main Content Area ---
if uploaded_file is not None:
    data = load_data(uploaded_file)

    if data is not None and not data.empty:
        
        # --- Origin Filter ---
        selected_origins = []
        if 'origin' in data.columns:
            all_origins = data['origin'].dropna().unique().tolist()
            selected_origins = origin_filter_placeholder.multiselect(
                "Filter by Case Origin",
                options=all_origins,
                default=all_origins
            )
            data = data[data['origin'].isin(selected_origins)]

        # Create a separate dataframe for time-based analysis, which requires valid dates
        time_based_data = data.dropna(subset=['start_date_and_time', 'end_date_and_time', 'work_order_number', 'case_number', 'date_time_opened']).copy()
        
        # --- Pre-calculate all time metrics for a consistent base ---
        all_case_times = calculate_case_times(time_based_data)
        all_start_delays = calculate_delayed_starts(time_based_data)

        # --- Apply Filters Sequentially ---
        # Filter 1: Start Delay on Work Orders
        start_outlier_wos = all_start_delays[all_start_delays['start_delay'] > pd.Timedelta(days=max_start_delay_days)]['work_order_number']
        cases_after_start_filter = time_based_data[~time_based_data['work_order_number'].isin(start_outlier_wos)]['case_number'].unique()
        
        # Apply this case filter to our master time metrics table
        times_after_start_filter = all_case_times[all_case_times['case_number'].isin(cases_after_start_filter)]
        
        # Filter 2: Max MTTR on Cases
        times_after_mttr_filter = times_after_start_filter[times_after_start_filter['resolution_time'] <= pd.Timedelta(days=max_mttr_days)]
        
        # Filter 3 & 4: IQR-based Outlier Removal
        final_times_data = times_after_mttr_filter.copy()
        median_mttr_outlier_count, mttr_q1, mttr_q3, mttr_iqr = 0, None, None, None
        if remove_mttr_iqr_outliers:
            final_times_data, median_mttr_outlier_count, mttr_q1, mttr_q3, mttr_iqr = remove_iqr_outliers(final_times_data, 'resolution_time')

        median_ttr_outlier_count, ttr_q1, ttr_q3, ttr_iqr = 0, None, None, None
        if remove_ttr_iqr_outliers:
            final_times_data, median_ttr_outlier_count, ttr_q1, ttr_q3, ttr_iqr = remove_iqr_outliers(final_times_data, 'response_time')

        # This is the final, clean dataset for all time-based KPIs
        final_analyzed_cases = final_times_data['case_number'].unique()
        final_analyzed_data = time_based_data[time_based_data['case_number'].isin(final_analyzed_cases)]
        final_wo_count = final_analyzed_data['work_order_number'].nunique()

        st.success(f"Successfully loaded **{uploaded_file.name}**.")
        
        total_wo_in_file = data['work_order_number'].nunique() if 'work_order_number' in data.columns else 0
        initial_wo_count_time_data = time_based_data['work_order_number'].nunique()

        st.info(f"""
        **Filtering Summary:**
        - Total Work Orders in File: **{total_wo_in_file}**
        - WOs with valid time data: **{initial_wo_count_time_data}**
        - WOs removed by filters: **{initial_wo_count_time_data - final_wo_count}**
        - **Final Work Orders Analyzed: {final_wo_count}**
        """)

        # --- Calculate Metrics with Final, Consistent Datasets ---
        mttr_duration = final_times_data['resolution_time'].mean()
        ttr_duration = final_times_data['response_time'].mean()

        ftfr_rate, repeat_count, ftfr_data = calculate_ftfr(final_analyzed_data, repeat_visit_days)
        delayed_start_info_filtered = calculate_delayed_starts(final_analyzed_data)
        delayed_count = len(delayed_start_info_filtered[delayed_start_info_filtered['start_delay'] > pd.Timedelta(days=1)])
        
        def format_timedelta(td):
            if pd.notna(td):
                days = td.days
                hours, remainder = divmod(td.seconds, 3600)
                minutes, _ = divmod(remainder, 60)
                return f"{days}d {hours}h {minutes}m"
            return "N/A"

        mttr_str = format_timedelta(mttr_duration)
        ttr_str = format_timedelta(ttr_duration)

        # --- Display Key Metrics ---
        st.header("High-Level KPIs")
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Work Orders Analyzed", value=final_wo_count)
        col2.metric("Mean Time to Resolution", value=mttr_str)
        col3.metric("Mean Time to Respond", value=ttr_str)
        col4.metric("First-Time Fix Rate (Proxy)", value=f"{ftfr_rate:.1%}")
        col5.metric("Delayed Starts (>24h)", value=delayed_count)


        # --- Create Tabs for Detailed Analysis ---
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📊 Overview", "⏱️ MTTR & Respond Time", "🔧 FTFR Analysis", "🔩 Component Analysis", "💰 Parts & Cost Analysis", "📄 Raw Data Explorer"])

        with tab1:
            st.subheader("Work Order Overview")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### Activity by Line Type")
                if 'line_type' in data.columns:
                    wo_type_counts = final_analyzed_data['line_type'].value_counts()
                    st.bar_chart(wo_type_counts)
                    if show_data_tables:
                        st.dataframe(wo_type_counts)


            with col2:
                st.markdown("#### Activity by Technician")
                if 'owner_full_name' in data.columns:
                    tech_activity = final_analyzed_data['owner_full_name'].value_counts()
                    st.bar_chart(tech_activity)
                    if show_data_tables:
                        st.dataframe(tech_activity)

            col3, col4 = st.columns(2)
            with col3:
                st.markdown("#### Total Labor Hours by Technician")
                if 'owner_full_name' in final_analyzed_data.columns and 'line_type' in final_analyzed_data.columns:
                    labor_hours_df = final_analyzed_data[final_analyzed_data['line_type'].str.lower() == 'labor'].copy()
                    if not labor_hours_df.empty:
                        labor_hours_df['duration_hours'] = (labor_hours_df['end_date_and_time'] - labor_hours_df['start_date_and_time']) / np.timedelta64(1, 'h')
                        tech_hours = labor_hours_df.groupby('owner_full_name')['duration_hours'].sum().nlargest(20)
                        st.bar_chart(tech_hours)
                        if show_data_tables:
                            st.dataframe(tech_hours)
            with col4:
                st.markdown("#### Cases by Origin")
                if 'origin' in final_analyzed_data.columns:
                    # **FIX**: Count unique cases per origin, not every row.
                    origin_case_data = final_analyzed_data[['case_number', 'origin']].drop_duplicates()
                    origin_counts = origin_case_data['origin'].value_counts()
                    st.bar_chart(origin_counts)
                    if show_data_tables:
                        st.dataframe(origin_counts)

            st.markdown("#### Work Orders Over Time")
            if 'date_time_opened' in data.columns:
                wo_over_time = final_analyzed_data.set_index('date_time_opened').resample('M')['work_order_number'].nunique()
                st.line_chart(wo_over_time)
                if show_data_tables:
                    st.dataframe(wo_over_time)

        with tab2:
            st.subheader("MTTR & Respond Time Analysis")
            st.markdown("These metrics are calculated *only* for cases containing at least one 'Labor' line item to ensure a fair comparison.")
            
            if not final_times_data.empty:
                # Prepare data for download
                csv_to_download = convert_df_to_csv(final_times_data[['case_number', 'resolution_time', 'response_time']])
                st.download_button(
                   label="Download Time Metrics as CSV",
                   data=csv_to_download,
                   file_name='case_time_metrics.csv',
                   mime='text/csv',
                )

            with st.expander("Show Statistical Details for Outlier Removal"):
                st.markdown("#### MTTR Statistics")
                if mttr_q1 is not None:
                    st.write(f"- **Q1 (25th Percentile):** {format_timedelta(mttr_q1)}")
                    st.write(f"- **Q3 (75th Percentile):** {format_timedelta(mttr_q3)}")
                    st.write(f"- **IQR (Interquartile Range):** {format_timedelta(mttr_iqr)}")
                else:
                    st.write("MTTR IQR filter not active.")

                st.markdown("#### TTR Statistics")
                if ttr_q1 is not None:
                    st.write(f"- **Q1 (25th Percentile):** {format_timedelta(ttr_q1)}")
                    st.write(f"- **Q3 (75th Percentile):** {format_timedelta(ttr_q3)}")
                    st.write(f"- **IQR (Interquartile Range):** {format_timedelta(ttr_iqr)}")
                else:
                    st.write("TTR IQR filter not active.")

            if not final_times_data.empty:
                final_times_data['resolution_hours'] = final_times_data['resolution_time'] / np.timedelta64(1, 'h')
                final_times_data['response_hours'] = final_times_data['response_time'] / np.timedelta64(1, 'h')

                st.markdown("#### Frequency Distribution of Resolution Time (MTTR)")
                st.bar_chart(final_times_data['resolution_hours'])
                if show_data_tables:
                    st.dataframe(final_times_data[['case_number', 'resolution_time', 'resolution_hours']].sort_values('resolution_time', ascending=False))

                st.markdown("---")
                st.markdown("#### Frequency Distribution of Response Time (TTR)")
                st.bar_chart(final_times_data['response_hours'])
                if show_data_tables:
                    st.dataframe(final_times_data[['case_number', 'response_time', 'response_hours']].sort_values('response_time', ascending=False))

        with tab3:
            st.subheader("First-Time Fix Rate Deep Dive")
            st.markdown(f"Based on our proxy, we identified **{int(repeat_count)}** repeat visits.")

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### Repeat Visits by Technician")
                repeats_df = ftfr_data[ftfr_data['is_repeat'] == True]
                if not repeats_df.empty and 'owner_full_name' in repeats_df.columns:
                    tech_repeats = repeats_df['owner_full_name'].value_counts()
                    st.bar_chart(tech_repeats)
                    if show_data_tables:
                        st.dataframe(tech_repeats)

            with col2:
                st.markdown("#### Cases with Most Repeat Visits")
                if not repeats_df.empty and 'case_number' in repeats_df.columns:
                    case_repeats = repeats_df['case_number'].value_counts().head(10)
                    st.bar_chart(case_repeats)
                    if show_data_tables:
                        st.dataframe(case_repeats)
        
        with tab4:
            st.subheader("Component Analysis")
            component_cols = [col for col in ['product_group', 'product_area', 'sub_assembly'] if col in data.columns]
            
            if component_cols:
                grouping_col = st.selectbox("Group Analysis By:", component_cols)

                st.markdown(f"#### Case Count by {grouping_col.replace('_', ' ').title()}")
                case_counts = data.groupby(grouping_col)['case_number'].nunique().nlargest(20)
                st.bar_chart(case_counts)
                if show_data_tables:
                    st.dataframe(case_counts)

                st.markdown(f"#### Work Order Count by {grouping_col.replace('_', ' ').title()}")
                wo_counts = data.groupby(grouping_col)['work_order_number'].nunique().nlargest(20)
                st.bar_chart(wo_counts)
                if show_data_tables:
                    st.dataframe(wo_counts)

                parts_data = data[data['line_type'].str.lower() == 'parts'].copy()
                if not parts_data.empty:
                    parts_data['total_line_price'] = pd.to_numeric(parts_data['total_line_price'], errors='coerce').fillna(0)
                    st.markdown(f"#### Total Parts Cost by {grouping_col.replace('_', ' ').title()}")
                    cost_by_group = parts_data.groupby(grouping_col)['total_line_price'].sum().nlargest(20)
                    st.bar_chart(cost_by_group)
                    if show_data_tables:
                        st.dataframe(cost_by_group)
            else:
                st.warning("No component columns (product_group, product_area, sub_assembly) found in the uploaded file.")

        with tab5:
            st.subheader("Parts & Cost Analysis")
            # Filter the original dataframe to only include parts lines
            parts_data = data[data['line_type'].str.lower() == 'parts'].copy()

            # Intelligently choose the best column for grouping parts data
            if 'installed_product_serial_number' in parts_data.columns:
                grouping_col = 'installed_product_serial_number'
                grouping_title = "Serial Numbers"
            elif 'installed_product_installed_product' in parts_data.columns:
                grouping_col = 'installed_product_installed_product'
                grouping_title = "Products"
            else:
                grouping_col = None

            if not parts_data.empty and grouping_col:
                # Clean up cost and quantity columns
                parts_data['total_line_price'] = pd.to_numeric(parts_data['total_line_price'], errors='coerce').fillna(0)
                parts_data['line_qty'] = pd.to_numeric(parts_data['line_qty'], errors='coerce').fillna(0)

                total_parts_cost = parts_data['total_line_price'].sum()
                total_parts_qty = parts_data['line_qty'].sum()

                st.markdown("#### Parts Consumption KPIs")
                col1, col2 = st.columns(2)
                col1.metric("Total Cost of Parts Used", value=f"${total_parts_cost:,.2f}")
                col2.metric("Total Quantity of Parts Used", value=f"{total_parts_qty:,.0f}")

                st.markdown("---")
                col1, col2 = st.columns(2)

                with col1:
                    st.markdown(f"#### Top 15 {grouping_title} by Parts Cost")
                    parts_by_cost = parts_data.groupby(grouping_col)['total_line_price'].sum().nlargest(15)
                    st.bar_chart(parts_by_cost)
                    if show_data_tables:
                        st.dataframe(parts_by_cost)
                
                with col2:
                    st.markdown(f"#### Top 15 {grouping_title} by Parts Quantity")
                    parts_by_qty = parts_data.groupby(grouping_col)['line_qty'].sum().nlargest(15)
                    st.bar_chart(parts_by_qty)
                    if show_data_tables:
                        st.dataframe(parts_by_qty)

            else:
                st.warning("No data with 'Line Type' of 'Parts' found or required product/serial number columns are missing.")


        with tab6:
            st.subheader("Explore the Full Raw Data")
            st.markdown("This shows the complete dataset before any filtering.")
            st.dataframe(data)

    else:
        st.info("No data to display. Please check the uploaded file.")

else:
    st.info("Awaiting your report file. Please upload a file in the sidebar to begin analysis.")
