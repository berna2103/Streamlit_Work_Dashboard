# Import necessary libraries
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from Create_Power_Point import add_slide_with_chart_and_text
import warnings





# Set title and configure Streamlit page layout
TITLE = 'Downtime Report'
DOWNTIME_URL = 'https://elekta.lightning.force.com/lightning/r/Report/00O6g000006RPtf/view'
QLIK_URL = 'https://qliksense.elekta.com/sense/app/6b74e786-8876-4f7a-8444-478355cc7b84/sheet/401b848c-b981-4694-abae-b1e347a1dfd8/state/analysis'
st.set_page_config(layout="wide")

# Suppress specific openpyxl warning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


total_hours = 3276 # default hours for uptime calculation 260 weekdays * 13 (8am - 9pm)
selected_service_agreement_uptime = '97%' 



# Styling the metrics for better UI presentation
st.markdown(
    """
    <style>
    [data-testid="stMetricValue"] {
        font-size: 45px;
        color: rgb(43, 101, 124);
        justify-content: center;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Hide metric delta change arrows (not relevant for the use case)
st.write(
    """
    <style>
    [data-testid="stMetricDelta"] svg {
        display: none;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Sidebar settings for file upload and other configurations
st.sidebar.title('Settings')

# Upload file section in the sidebar
uploaded_file = st.sidebar.file_uploader(label='Load data file:', key='downtime', type=['xlsx', 'xls','csv'])

# Add a visual divider in the sidebar for better separation of sections
st.sidebar.divider()

if uploaded_file is None:
# Customizing the title with color
    st.markdown("<h1 style='color: rgb(43, 101, 124);'>Load Downtime Report</h1>", unsafe_allow_html=True)
    st.write('1. Go to Reports on CLM and select Downtime Matrix (or {location} Downtime Report Last 120 Days).')
    st.markdown("CLM Downtime Report: (%s) " % DOWNTIME_URL)
    st.write('2. Select Lightning and Export > Details Only > Format .xlsx).')
    st.image('images/Guide/Downtime_report.png', width=400)
    st.write('3. Upload file.')
    st.write('4. Review Downtime data (compare to Qliksense).')
    st.markdown("Qliksese Link to Report: (%s) " % QLIK_URL)
    st.write('5. Click generate Power Point.')

# Check if a file is uploaded
if uploaded_file is not None:
    # Customizing the title with color
    st.markdown("<h1 style='color: rgb(43, 101, 124);'>Downtime Report</h1>", unsafe_allow_html=True)
    # Check the file extension
    file_extension = uploaded_file.name.split('.')[-1]

    # Read the uploaded Excel or CSV file
    if file_extension in ['xlsx', 'xls']:
        try:
            df = pd.read_excel(uploaded_file)
            print("Excel file loaded successfully!")
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
    elif file_extension == 'csv':
        try:
            # Specify encoding and error handling for CSV
            df = pd.read_csv(uploaded_file, encoding='utf-8', errors='replace')
            print("CSV file loaded successfully!")
        except UnicodeDecodeError:
            # Try with different encoding if utf-8 fails
            df = pd.read_csv(uploaded_file, encoding='ISO-8859-1', errors='replace')
            print("CSV file loaded successfully with ISO-8859-1 encoding!")
        except Exception as e:
            st.error(f"Error reading CSV file: {e}")

    # Define and remove unnecessary columns
    columns_to_remove = [
        'Case: Installed Product', 'Case: Customer Resolution Statement',
        'Exclude', 'Exclude Reason', 'Case: Opened Date'
    ]
    df.drop(columns=columns_to_remove, inplace=True)

    # Rename columns for easier manipulation
    df.rename(columns={
        'Case: Description': 'description',
        'Date Start of Down Time (Customer Time)': 'start date',
        'Start of Down Time (Customer Time)': 'start time',
        'Date End of Down Time (Customer Time)': 'end date',
        'End of Down Time (Customer Time)': 'end time',
        'Downtime Out Agreed Available Time': 'OAAT',
        'Downtime In Agreed Available Time': 'IAAT',
        'Case: Case Number': 'case',
        'Case: Location': 'location',
        'Case: Opened Date': 'opened_date'
    }, inplace=True)

    

    # Drop rows with any NaN values and ensure case column is of string type
    df_cleaned = df # Channge to df if some values are missing
    st.write(df)
    df_cleaned['case'] = df_cleaned['case'].astype(object)
    # df_cleaned[:,'case'] = df_cleaned['case'].astype(str)
    df_cleaned.loc[:,'case'] = df_cleaned['case'].astype(str)

    # Rearrange columns for better readability and analysis
    df_cleaned = df_cleaned.iloc[:, [1, 0, 7, 2, 3, 4, 5, 6, 8]]

    # Get unique locations for filtering
    locations = df['location'].unique()

    # Sidebar for selecting a specific location to filter
    locations = ['All'] + list(locations)  # Adding 'All' option at the beginning of locations list
    selected_location = st.sidebar.selectbox('Select Location:', locations)


    ######### Function to calculate uptime ##########################################################################

    def calculate_uptime_percentage(hours, total_hours=total_hours, agreement_type=selected_service_agreement_uptime):
        """
        Calculate the uptime percentage based on total hours (100% - calculated percentage).

        Parameters:
        hours (float): The number of hours to calculate the percentage for.
        total_hours (float): The total hours to base the percentage on (default is 3276).
        260 weekdays per year - 8 elekta holidays = 252 days * 13 hrs (8 a - 9 p)
        for contracts 8 to 5 the total number is: 2268

        Returns:
        float: The remaining percentage (100% - calculated percentage).
        """

        if total_hours == 0:  # Prevent division by zero
            raise ValueError("Total hours cannot be zero.")

        calculated_percentage = round(((hours / total_hours) * 100),1)
        remaining_percentage = round((100 - calculated_percentage),1)
        return round(remaining_percentage,1)
    
    def create_pie_chart(iaat_hours, total_hours=total_hours):
        """
        Create a pie chart showing uptime percentage and IAAT.

        Parameters:
        iaat_hours (float): The number of IAAT hours (Inside Agreed Available Time).
        total_hours (float): The total available hours to calculate percentage (default is 3276).

        Returns:
        Plotly pie chart showing the Uptime and IAAT percentages.
        """
        if total_hours == 0:
            raise ValueError("Total hours cannot be zero.")

        # Calculate IAAT percentage and Uptime percentage
        uptime_percentage = (total_hours - iaat_hours) / total_hours * 100
        iaat_percentage = 100 - uptime_percentage
        # Pie chart labels and values
        labels = ['Uptime', 'IAAT']
        values = [uptime_percentage, iaat_percentage]

        # Create pie chart
        fig = go.Figure(data=[go.Pie(labels=labels, values=values)])

        # Add title and customize chart appearance
        fig.update_traces(marker=dict(colors=['rgb(43, 101, 124)', 'rgb(54, 164, 179)']), 
                          textinfo='label+percent', 
                          hoverinfo='label+percent')

        fig.update_layout(
            title_text="Uptime vs Downtime IAAT (Inside Agreed Available Time)",
        )

        st.write(fig)


    ######### Histogram graphs #########################################################################################
    def graph_data(locations):

        # Create a Streamlit container to display data
        container = st.container()

        # Loop through locations and generate histograms for each
        for i in range(1, len(locations), 1):
            current_locations = locations[i:i + 1]
            cols = container.columns(1)  # Create two columns for display
        
            # Loop through each location and create histograms
            for j, location in enumerate(current_locations):
                # Filter data based on the current location
                filtered_df = df_cleaned[df_cleaned['location'] == location].copy()
                df_cleaned['start time'] = df_cleaned['start time'].astype(str)
                filtered_df['start date'] = pd.to_datetime(filtered_df['start date'], errors='coerce')
                filtered_df['end date'] = pd.to_datetime(filtered_df['end date'], errors='coerce')

                # Group data by month
                filtered_df['month'] = filtered_df['start date'].dt.to_period('M')
                iaat_downtime = round(filtered_df['IAAT'].sum(), 2)
                oaat_downtime = round(filtered_df['OAAT'].sum(), 2)

                # Aggregate IAAT and OAAT downtime by month
                monthly_data = filtered_df.groupby('month')[['IAAT', 'OAAT']].sum().reset_index()
                monthly_data['month'] = monthly_data['month'].dt.to_timestamp()

                # Create histogram for IAAT and OAAT
                # fig = px.histogram(monthly_data, x='month', y=['IAAT', 'OAAT'],
                #                    barmode='group',
                #                    nbins=12,
                #                    color_discrete_sequence=['rgb(43, 101, 125)', 'rgb(54, 164, 179)'],
                #                    labels={'OAAT': 'Outside Agreed Available Time', 'IAAT': 'Inside Agreed Available Time'})
                # fig.update_layout(bargap=0.02, title='Months with Downtime', yaxis_title='Downtime Hours')
                fig = px.bar(monthly_data, 
                    x='month', 
                    y=['IAAT', 'OAAT'], 
                    barmode='group',
                    color_discrete_sequence=['rgb(43, 101, 125)', 'rgb(54, 164, 179)'],
                    labels={'OAAT': 'Outside Agreed Available Time', 'IAAT': 'Inside Agreed Available Time'})

                fig.update_layout(bargap=0.4)  # Decrease for wider bars

                # Display the chart in the corresponding column
                with cols[j]:
                    with st.container(border=True):
                        st.markdown(f"<h4 style='color: rgb(43, 101, 124);'>{location}</h4>", unsafe_allow_html=True) # Change text color in Downtime metric to match company
                        metric1, metric2 = st.columns(2)
                        graph1, graph2 = st.columns(2)
                        tab1, tab2 = st.tabs(["📈 Chart", "🗃 Data"])
            

                        # Display histogram in tab 1
                        with tab1:
                            with metric1:
                                st.metric('Total Downtime', f'{iaat_downtime} hrs', delta=f'*OAAT {oaat_downtime} hrs', delta_color='off')
                            with metric2:
                                st.metric('Calculated Uptime %', f'{calculate_uptime_percentage(iaat_downtime, total_hours, agreement_type=selected_service_agreement_uptime)}%', delta=f'{selected_service_agreement_uptime} target')
                            
                            with graph1:
                                st.plotly_chart(fig)

                            with graph2:    
                                st.write(create_pie_chart(iaat_downtime, total_hours=total_hours))
                            st.caption('*IAAT - In Agreed Available Time')
                            st.caption('*OAAT - Out Agreed Available Time')

                        # Display the filtered DataFrame in tab2
                        with tab2:
                            # Display raw data in tab2
                            st.write(filtered_df)

    ######################

    ######################
    
    # Streamlit Sidebar
    st.sidebar.title('Create Downtime PowerPoint Slides')

    # Button to generate PowerPoint
    if st.sidebar.button('Generate PowerPoint'):
        add_slide_with_chart_and_text('Downtime_Report', df_cleaned, locations, total_hours )
        st.success("PowerPoint saved successfully!")
        
    # Calculte Downtime based on (8a to 9pm) or (8a - 5p)
    calculate_8_to_5 = st.sidebar.checkbox("Downtime 8am to 5 pm.")
    if calculate_8_to_5:
        total_hours = 2268
    else:
        total_hours = 3276

    # Select type of service agreement
    option = st.sidebar.selectbox('Service agreement type:', ("Silver", "Gold", "Platinum"), index=None, placeholder="agreement type ...")
    match option:
        case 'Silver':
            selected_service_agreement_uptime = '95%'
        case 'Gold':
            selected_service_agreement_uptime = '97%'
        case 'Platinum':
            selected_service_agreement_uptime = '98%'

    # Filter data based on selected location or show all locations if 'All' is selected
    if selected_location == 'All':
        filtered_df = df_cleaned  # Show all locations
        graph_data(locations)
    else:
        arr = np.array(['All'])
        # filtered_df = df_cleaned[df_cleaned['location'] == selected_location]  # Filter by selected location
        if type(selected_location) == str:
            locations = np.append(arr, selected_location)
            graph_data(locations)
