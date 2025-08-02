import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# File Upload
st.set_page_config(page_title="Downtime Dashboard", layout="wide")
st.title("📊 Downtime Dashboard")

uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["xlsx", "csv"])
total_hours = 3276 

if uploaded_file:
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    df.columns = df.columns.str.strip()
    df_cleaned = df.copy()

    locations = sorted(df_cleaned['location'].dropna().unique())
    selected_locations = st.multiselect("Select Locations", locations, default=locations)
    selected_service_agreement_uptime = st.selectbox("Target Uptime %", [95, 97, 98, 99, 99.5, 99.9], index=3)

    def create_pie_chart(downtime, total_hours):
        uptime = total_hours - downtime
        labels = ['Uptime', 'Downtime']
        values = [uptime, downtime]
        fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.5)])
        fig.update_traces(marker=dict(colors=['#2b657d', '#97d2e8']))
        fig.update_layout(title='Uptime vs Downtime')
        return fig

    def calculate_uptime_percentage(downtime, total_hours, agreement_type):
        uptime = total_hours - downtime
        uptime_percentage = (uptime / total_hours) * 100 if total_hours > 0 else 0
        return round(uptime_percentage, 2)

    def graph_data(df_cleaned, locations, selected_service_agreement_uptime):
        container = st.container()

        df_cleaned['Device Downtime'] = pd.to_numeric(df_cleaned['Device Downtime'], errors='coerce').fillna(0)
        # total_hours = df_cleaned['Device Downtime'].sum() if not df_cleaned.empty else 1

        for i in range(0, len(locations), 1):
            current_locations = locations[i:i + 1]
            cols = container.columns(1)

            for j, location in enumerate(current_locations):
                filtered_df = df_cleaned[df_cleaned['location'] == location].copy()
                filtered_df['start date'] = pd.to_datetime(filtered_df['start date'], errors='coerce')
                filtered_df['month'] = filtered_df['start date'].dt.to_period('M')

                iaat_downtime = round(filtered_df['Device Downtime'].sum(), 2)
                oaat_downtime = 0

                monthly_data = filtered_df.groupby('month')[['Device Downtime']].sum().reset_index()
                monthly_data['month'] = monthly_data['month'].dt.to_timestamp()

                fig = px.histogram(monthly_data, x='month', y='Device Downtime',
                                   barmode='group',
                                   color_discrete_sequence=['rgb(43, 101, 125)'],
                                   nbins=12,
                                   labels={'Device Downtime': 'Downtime Hours'})
                fig.update_layout(bargap=0.5, title='Downtime', yaxis_title='Downtime Hours')

                with cols[j]:
                    with st.container(border=True):
                        st.markdown(f"<h4 style='color: rgb(43, 101, 124);'>{location}</h4>", unsafe_allow_html=True)
                        metric1, metric2 = st.columns(2)
                        graph1, graph2 = st.columns(2)
                        tab1, tab2 = st.tabs(["📈 Chart", "🗃 Data"])

                        with tab1:
                            with metric1:
                                st.metric('Total Downtime', f'{iaat_downtime} hrs', delta_color='off')
                            with metric2:
                                st.metric('Calculated Uptime %', f'{calculate_uptime_percentage(iaat_downtime, total_hours, agreement_type=selected_service_agreement_uptime)}%', delta=f'{selected_service_agreement_uptime} target')

                            with graph1:
                                st.plotly_chart(fig)
                            with graph2:
                                st.plotly_chart(create_pie_chart(iaat_downtime, total_hours=total_hours))

                           

                        with tab2:
                            st.write(filtered_df)

    graph_data(df_cleaned, selected_locations, selected_service_agreement_uptime)
else:
    st.info("📁 Upload an Excel or CSV file to begin.")