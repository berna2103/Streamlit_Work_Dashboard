# Import necessary libraries
import streamlit as st
import pandas as pd
import plotly.express as px
from Parts_Slides import generate_parts_slides


st.set_page_config(layout="wide")

st.sidebar.title('Settings')
# parts_file = st.sidebar.file_uploader(label='Load parts file:', key='parts', type=['xlsx', 'xls', 'csv'])

PARTS_URL = 'https://elekta.lightning.force.com/lightning/r/Report/00O0d0000055hArEAI/view?queryScope=userFolders'

st.sidebar.divider()

# Check if a file is uploaded
# Sidebar file uploader
uploaded_file = st.sidebar.file_uploader(label='Load parts file:', key='parts', type=['xlsx', 'xls', 'csv'])

if uploaded_file is None:
# Customizing the title with color
    st.markdown("<h1 style='color: rgb(43, 101, 124);'>Load Parts Report</h1>", unsafe_allow_html=True)
    st.write('**1**. Go to Reports on CLM and select **Parts (or {location} Parts Report Last 120 Days)**.')
    st.markdown("CLM Parts Report: (%s) " % PARTS_URL)
    st.write('**2**. Select **Lightning** and **Export** > **Details Only** > Format .csv).')
    st.image('images/Guide/Parts_CLM.png', width=400)
    st.write('**3**. Upload file.')
    st.write('**4**. Review Parts data (compare to CLM), CLM report will be a bit lower it does not add up if the same part has been consumed multiple times.')
    st.write('**5**. Click generate Power Point.')


# Check if a file is uploaded
if uploaded_file is not None:
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
            df = pd.read_csv(uploaded_file, encoding='utf-8')
            print("CSV file loaded successfully!")
        except UnicodeDecodeError:
            # Try with different encoding if utf-8 fails
            df = pd.read_csv(uploaded_file, encoding='ISO-8859-1')
            print("CSV file loaded successfully with ISO-8859-1 encoding!")
        except Exception as e:
            st.error(f"Error reading CSV file: {e}")

  
        # Define and remove unnecessary columns
    columns_to_remove = [
        'Work Detail: Line Number', 'Line Price Per Unit Currency'
    ]

    df.drop(columns=columns_to_remove, inplace=True)

        # Rename columns for easier manipulation
    df.rename(columns={
        'Work Order: Work Order Number': 'work_order',
        'Work Detail: Created Date': 'created_date',
        'Item Number': 'part_number',
        'Item Qty': 'qty',
        'Line Price Per Unit': 'price_per_unit',
        'Consumed From Location': 'location',
        'Installed Product': 'ip'
    }, inplace=True)

    #convert 'created_date' column to datetime format
    df['created_date'] = pd.to_datetime(df['created_date'])

    # Group by 'ip' and 'created_date' to aggregate the consumed parts
    df_grouped_ip = df.groupby(['ip', 'created_date']).agg(
        total_qty=pd.NamedAgg(column='qty',aggfunc='sum'),
        total_cost=pd.NamedAgg(column='price_per_unit',aggfunc=lambda x: (x*df['qty']).sum())
    ).reset_index()


    # Create a histogram to show part consumption by 'ip'
    st.title("Part Consumption All Locations")

    with st.container(border=True):

        # Display metrics for total cost and total number of parts
        total_cost_ip = df_grouped_ip['total_cost'].sum()
        total_parts_ip = df_grouped_ip['total_qty'].sum()

        st.metric(label="Total Cost", value=f"${total_cost_ip:,.2f}")
        st.metric(label="Total Number of Parts", value=total_parts_ip)
        
        # Plotly histogram
        fig = px.histogram(df_grouped_ip, x='created_date', y='total_cost', color='ip',
                       title='Part Consumption by Installed Product',
                       labels={'created_date': 'Date', 'total_qty': 'Total Quantity'},
                       nbins=10, barmode='group')

        #Display histogram
        st.plotly_chart(fig)

    # Create a Streamlit container for the entire report
    st.title("Parts Consumption Report by Installed Product (IP)")

    # Loop through each unique locations and generate separate histograms and metrics
    for ip in df['ip'].unique():
        # Create a Streamlit container to display data
        container = st.container()
        

        #Filter data for the current IP
        df_ip = df[df['ip'] == ip].copy()

        #Group by loaction and date to calculate total quantity and cost
        df_grouped = df_ip.groupby(['location', 'created_date']).agg(
        total_qty=pd.NamedAgg(column='qty', aggfunc='sum'),
        total_cost=pd.NamedAgg(column='price_per_unit', aggfunc=lambda x: (x * df_ip.loc[x.index, 'qty']).sum()) ).reset_index()

    
         # Display metrics for total cost and total number of parts
        total_cost_ip = df_grouped['total_cost'].sum()
        total_parts_ip = df_grouped['total_qty'].sum()

        

        with st.container(border=True):
            st.subheader(f'Installed Product: {ip}')
            metric1, metric2 = st.columns(2) 
            
            with metric1:
                st.metric(label="Total Cost", value=f"${total_cost_ip:,.2f}")
                st.metric(label="Total Number of Parts", value=total_parts_ip)
                
            # Calculate and display the three most expensive items
            df_ip.loc[:,'total_item_cost'] = df_ip['price_per_unit'] * df_ip['qty']
            top_3_items = df_ip[['Item', 'total_item_cost']].sort_values(by='total_item_cost', ascending=False).head(3)
            
            with metric2:
                st.markdown("### High Cost Items:")
                for index, row in top_3_items.iterrows():
                    st.caption(f"{row['Item']}: ${row[f'total_item_cost']:,.2f}")

            # Create histogram for the current IP
            fig = px.histogram(df_grouped, x='created_date', y='total_cost', color='location',
                               title=f'Part Consumption by month.',
                               color_discrete_sequence=['rgb(43, 101, 125)', 'rgb(54, 164, 179)'],
                               labels={'created_date': 'Date', 'total_qty': 'Total Quantity'},
                               nbins=10, barmode='group')

            # Display the histogram in Streamlit
            st.plotly_chart(fig)
            

            show_data = st.checkbox('Show data:', key=f'{ip}')
            if show_data:
                st.dataframe(df_ip)
        
    # Streamlit Sidebar
    st.sidebar.title('Create Downtime PowerPoint Slides')

    # Button to generate PowerPoint
    if st.sidebar.button('Generate PowerPoint'):
        generate_parts_slides('Parts Report', df, df['ip'].unique(), df_grouped_ip )
        st.success("Part Slides saved successfully!")