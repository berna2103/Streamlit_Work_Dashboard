import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import random
from io import BytesIO

# --- Page Configuration ---
st.set_page_config(
    page_title="E-commerce Sales Performance Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom CSS for a cleaner, modern look ---
st.markdown("""
    <style>
    .main {
        background-color: #f0f2f6;
        padding: 20px;
    }
    .st-emotion-cache-z5fcl4 { /* Adjust padding for main content area */
        padding-top: 2rem;
    }
    .st-emotion-cache-1cyp85f { /* Adjust padding for columns */
        padding-top: 0rem;
        padding-bottom: 0rem;
    }
    .st-emotion-cache-183p0q { /* Adjust font size for headers */
        font-size: 1.2rem;
        font-weight: bold;
        color: #262730;
    }
    .block-container { /* General block container padding */
        padding-top: 1rem;
        padding-bottom: 0rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    .stButton>button { /* Styling for buttons */
        background-color: #4CAF50;
        color: white;
        border-radius: 8px;
        padding: 10px 20px;
        font-size: 16px;
        border: none;
        cursor: pointer;
        transition: all 0.3s ease;
    }
    .stButton>button:hover { /* Hover effect for buttons */
        background-color: #45a049;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
    }
    /* Styling for containers with border */
    .st-emotion-cache-1r6dm1s {
        background-color: #ffffff; /* White background for cards/charts */
        border-radius: 12px; /* More rounded corners */
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08); /* Stronger, softer shadow */
        padding: 20px;
        margin-bottom: 20px;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #2c3e50; /* Darker titles */
    }
    .st-emotion-cache-1r6dm1s h2 { /* Specific styling for section headers within containers */
        color: #34495e;
        border-bottom: 2px solid #e0e0e0;
        padding-bottom: 10px;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# --- Dummy Data Generation (for demonstration purposes) ---
@st.cache_data # Cache the data so it doesn't regenerate on every rerun
def generate_dummy_data(rows=1000):
    data = []
    end_date = datetime.now()
    start_date = end_date - timedelta(days=365 * 2) # 2 years of data

    product_categories = ['Electronics', 'Apparel', 'Home Goods', 'Books', 'Groceries']
    regions = ['North', 'South', 'East', 'West', 'Central']
    customer_segments = ['New Customer', 'Returning Customer']

    for i in range(rows):
        date = start_date + timedelta(days=random.randint(0, (end_date - start_date).days))
        category = random.choice(product_categories)
        region = random.choice(regions)
        segment = random.choice(customer_segments)
        
        units_sold = random.randint(1, 15)
        
        # Simulate varying sales/profit based on category
        if category == 'Electronics':
            sales = round(random.uniform(50, 1500), 2) * units_sold
            profit = round(sales * random.uniform(0.1, 0.3), 2)
        elif category == 'Apparel':
            sales = round(random.uniform(20, 300), 2) * units_sold
            profit = round(sales * random.uniform(0.2, 0.4), 2)
        elif category == 'Home Goods':
            sales = round(random.uniform(30, 800), 2) * units_sold
            profit = round(sales * random.uniform(0.15, 0.35), 2)
        elif category == 'Books':
            sales = round(random.uniform(10, 80), 2) * units_sold
            profit = round(sales * random.uniform(0.3, 0.5), 2)
        else: # Groceries
            sales = round(random.uniform(5, 100), 2) * units_sold
            profit = round(sales * random.uniform(0.05, 0.15), 2)

        data.append([date, category, region, units_sold, sales, profit, segment])

    df = pd.DataFrame(data, columns=['Date', 'Product Category', 'Region', 'Units Sold', 'Sales ($)', 'Profit ($)', 'Customer Segment'])
    df['Average Order Value ($)'] = df['Sales ($)'] / df['Units Sold']
    return df

df = generate_dummy_data()

# --- Title and Description ---
st.title("📈 E-commerce Sales Performance Dashboard")
st.markdown(
    """
    This interactive dashboard showcases key sales performance metrics for an e-commerce business.
    Explore trends, analyze product and regional performance, and understand customer behavior.
    """
)

# --- Sidebar Filters ---
st.sidebar.header("Filters")

# Date Range Filter
min_date = df['Date'].min().to_pydatetime()
max_date = df['Date'].max().to_pydatetime()
date_range = st.sidebar.date_input(
    "Select Date Range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

if len(date_range) == 2:
    filtered_df = df[(df['Date'] >= pd.to_datetime(date_range[0])) & (df['Date'] <= pd.to_datetime(date_range[1]))].copy()
else:
    filtered_df = df.copy() # Show all if date range is not fully selected

# Product Category Filter
all_categories = ['All'] + sorted(filtered_df['Product Category'].unique().tolist())
selected_categories = st.sidebar.multiselect("Filter by Product Category", all_categories, default='All')
if 'All' not in selected_categories:
    filtered_df = filtered_df[filtered_df['Product Category'].isin(selected_categories)]

# Region Filter
all_regions = ['All'] + sorted(filtered_df['Region'].unique().tolist())
selected_regions = st.sidebar.multiselect("Filter by Region", all_regions, default='All')
if 'All' not in selected_regions:
    filtered_df = filtered_df[filtered_df['Region'].isin(selected_regions)]

# Customer Segment Filter
all_segments = ['All'] + sorted(filtered_df['Customer Segment'].unique().tolist())
selected_segments = st.sidebar.multiselect("Filter by Customer Segment", all_segments, default='All')
if 'All' not in selected_segments:
    filtered_df = filtered_df[filtered_df['Customer Segment'].isin(selected_segments)]

st.sidebar.markdown("---")

# Check if filtered data is empty
if filtered_df.empty:
    st.warning("No data matches the selected filters. Please adjust your selections.")
else:
    # --- KPIs ---
    st.header("Overall Sales Metrics")
    with st.container():
        col1, col2, col3, col4 = st.columns(4)

        total_sales = filtered_df['Sales ($)'].sum()
        total_units_sold = filtered_df['Units Sold'].sum()
        avg_order_value = filtered_df['Average Order Value ($)'].mean()
        total_profit = filtered_df['Profit ($)'].sum()

        with col1:
            st.metric("Total Sales", f"${total_sales:,.2f}")
        with col2:
            st.metric("Total Units Sold", f"{total_units_sold:,.0f}")
        with col3:
            st.metric("Average Order Value", f"${avg_order_value:,.2f}")
        with col4:
            st.metric("Total Profit", f"${total_profit:,.2f}")

    st.markdown("---")

    # --- Sales Trend Over Time ---
    st.header("Sales Trend Over Time")
    with st.container():
        # Aggregate sales by day for trend
        daily_sales = filtered_df.groupby(pd.to_datetime(filtered_df['Date']).dt.to_period('D'))['Sales ($)'].sum().reset_index()
        daily_sales['Date'] = daily_sales['Date'].dt.to_timestamp() # Convert Period to Timestamp for Plotly
        
        fig_sales_trend = px.line(daily_sales, x='Date', y='Sales ($)',
                                  title='Daily Sales Trend',
                                  labels={'Sales ($)': 'Sales ($)', 'Date': 'Date'},
                                  color_discrete_sequence=px.colors.qualitative.Plotly,
                                  template="plotly_white")
        fig_sales_trend.update_layout(hovermode="x unified") # Nice hover effect
        st.plotly_chart(fig_sales_trend, use_container_width=True)

    st.markdown("---")

    # --- Sales by Product Category & Region ---
    st.header("Sales Distribution")
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        with st.container():
            sales_by_category = filtered_df.groupby('Product Category')['Sales ($)'].sum().reset_index()
            sales_by_category = sales_by_category.sort_values(by='Sales ($)', ascending=False)
            fig_category_sales = px.bar(sales_by_category, x='Sales ($)', y='Product Category', orientation='h',
                                        title='Sales by Product Category',
                                        labels={'Sales ($)': 'Total Sales ($)', 'Product Category': 'Product Category'},
                                        color='Product Category',
                                        color_discrete_sequence=px.colors.qualitative.Pastel,
                                        template="plotly_white")
            fig_category_sales.update_layout(showlegend=False, yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_category_sales, use_container_width=True)

    with col_chart2:
        with st.container():
            sales_by_region = filtered_df.groupby('Region')['Sales ($)'].sum().reset_index()
            sales_by_region = sales_by_region.sort_values(by='Sales ($)', ascending=False)
            fig_region_sales = px.pie(sales_by_region, values='Sales ($)', names='Region',
                                      title='Sales by Region',
                                      color_discrete_sequence=px.colors.qualitative.Vivid,
                                      template="plotly_white")
            fig_region_sales.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_region_sales, use_container_width=True)

    st.markdown("---")

    # --- Profitability by Product Category & Sales by Customer Segment ---
    st.header("Profitability & Customer Insights")
    col_chart3, col_chart4 = st.columns(2)

    with col_chart3:
        with st.container():
            profit_by_category = filtered_df.groupby('Product Category')['Profit ($)'].sum().reset_index()
            profit_by_category = profit_by_category.sort_values(by='Profit ($)', ascending=False)
            fig_profit_category = px.bar(profit_by_category, x='Profit ($)', y='Product Category', orientation='h',
                                         title='Profit by Product Category',
                                         labels={'Profit ($)': 'Total Profit ($)', 'Product Category': 'Product Category'},
                                         color='Product Category',
                                         color_discrete_sequence=px.colors.qualitative.Safe,
                                         template="plotly_white")
            fig_profit_category.update_layout(showlegend=False, yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_profit_category, use_container_width=True)

    with col_chart4:
        with st.container():
            sales_by_segment = filtered_df.groupby('Customer Segment')['Sales ($)'].sum().reset_index()
            fig_segment_sales = px.pie(sales_by_segment, values='Sales ($)', names='Customer Segment',
                                       title='Sales by Customer Segment',
                                       color_discrete_sequence=px.colors.qualitative.G10,
                                       template="plotly_white")
            fig_segment_sales.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_segment_sales, use_container_width=True)

    st.markdown("---")

    # --- Detailed Data Table ---
    st.header("Detailed Sales Data")
    with st.container():
        st.dataframe(filtered_df.sort_values(by='Date', ascending=False), use_container_width=True)

        # Download button for filtered data
        csv_buffer = BytesIO()
        filtered_df.to_csv(csv_buffer, index=False, encoding='utf-8')
        csv_buffer.seek(0)
        st.download_button(
            label="Download Filtered Data as CSV",
            data=csv_buffer,
            file_name="e_commerce_sales_data.csv",
            mime="text/csv",
        )

