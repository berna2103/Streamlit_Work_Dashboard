import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import base64
import os
import numpy as np
from datetime import datetime, timedelta
from icalendar import Calendar, Event

# --- Page Configuration ---
st.set_page_config(
    page_title="PM Task Analysis Dashboard",
    page_icon="📊",
    layout="wide"
)

# --- Helper Functions ---

def load_data(uploaded_file):
    """Loads and cleans the data from the uploaded CSV file."""
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            # --- Data Cleaning ---
            if 'Option ID' in df.columns:
                df.rename(columns={'Option ID': 'System'}, inplace=True)

            df['Duration (mins)'] = pd.to_numeric(df['Duration (mins)'], errors='coerce')
            df['Interval (months)'] = pd.to_numeric(df['Interval (months)'], errors='coerce')
            df.fillna({'Duration (mins)': 0, 'Interval (months)': 0}, inplace=True)
            
            if 'Category of PM check' in df.columns:
                df['Category of PM check'] = df['Category of PM check'].str.strip().str.upper()
            if 'System' in df.columns:
                df['System'] = df['System'].str.strip().str.upper().fillna('NOT SPECIFIED')

            return df
        except Exception as e:
            st.error(f"Error loading or processing file: {e}")
            return None
    return None

def create_summary_metrics(df):
    """Creates and displays summary cards based on the filtered data."""
    total_tasks = len(df)
    total_duration_hours = df['Duration (mins)'].sum() / 60
    unique_systems = df['System'].nunique() if 'System' in df.columns else 'N/A'
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total PM Tasks", f"{total_tasks}")
    col2.metric("Total Duration (Hours)", f"{total_duration_hours:.2f}")
    col3.metric("Unique Systems", f"{unique_systems}")

# --- Charting Functions ---

def create_category_duration_chart(df):
    """Bar chart: Total duration by PM category."""
    category_duration = df.groupby('Category of PM check')['Duration (mins)'].sum().reset_index()
    fig = px.bar(
        category_duration, x='Category of PM check', y='Duration (mins)',
        title='Total Maintenance Duration by Category',
        labels={'Category of PM check': 'Category', 'Duration (mins)': 'Total Duration (Minutes)'},
        color='Category of PM check', template='plotly_white'
    )
    fig.update_layout(showlegend=False)
    return fig

def create_task_count_chart(df):
    """Pie chart: Number of tasks per category."""
    task_counts = df['Category of PM check'].value_counts().reset_index()
    task_counts.columns = ['Category', 'Count']
    fig = px.pie(
        task_counts, names='Category', values='Count',
        title='Distribution of Tasks by Category', hole=0.3, color='Category'
    )
    return fig

def create_interval_category_breakdown_chart(df):
    """Stacked bar chart: Duration by interval, broken down by category."""
    interval_category_duration = df.groupby(['Interval (months)', 'Category of PM check'])['Duration (mins)'].sum().reset_index()
    fig = px.bar(
        interval_category_duration, x='Interval (months)', y='Duration (mins)',
        color='Category of PM check', title='Duration Breakdown by Interval and Category',
        labels={'Interval (months)': 'PM Interval (Months)', 'Duration (mins)': 'Total Duration (Minutes)', 'Category of PM check': 'Category'},
        template='plotly_white', barmode='stack'
    )
    fig.update_xaxes(type='category')
    return fig

def create_system_duration_chart(df):
    """Bar chart: Total duration by System."""
    if 'System' not in df.columns: return go.Figure().update_layout(title_text="System data not available.")
    system_duration = df.groupby('System')['Duration (mins)'].sum().sort_values(ascending=False).reset_index()
    fig = px.bar(
        system_duration, x='Duration (mins)', y='System',
        orientation='h', title='Total Maintenance Duration by System',
        labels={'System': 'System', 'Duration (mins)': 'Total Duration (Minutes)'},
        template='plotly_white'
    )
    fig.update_layout(yaxis={'categoryorder':'total ascending'})
    return fig

def create_hierarchical_chart(df):
    """Treemap: Hierarchical view of duration by System and Category."""
    if 'System' not in df.columns or 'Category of PM check' not in df.columns: return go.Figure().update_layout(title_text="System or Category data not available.")
    fig = px.treemap(
        df, path=[px.Constant("All Systems"), 'System', 'Category of PM check'],
        values='Duration (mins)', title='Hierarchical View of Maintenance Duration',
        color='System', template='plotly_white'
    )
    fig.update_traces(root_color="lightgrey")
    fig.update_layout(margin = dict(t=50, l=25, r=25, b=25))
    return fig

def create_longest_tasks_chart(df):
    """Bar chart of the top 10 longest individual tasks."""
    longest_tasks = df.nlargest(10, 'Duration (mins)')
    fig = px.bar(
        longest_tasks, x='Duration (mins)', y='Task Description',
        orientation='h', title='Top 10 Longest Individual Tasks',
        labels={'Task Description': 'Task', 'Duration (mins)': 'Duration (Minutes)'},
        template='plotly_white'
    )
    fig.update_layout(yaxis={'categoryorder':'total ascending'})
    return fig

def create_maintenance_burden_chart(df):
    """Bar chart showing a calculated 'Maintenance Burden Score'."""
    if 'System' not in df.columns: return go.Figure().update_layout(title_text="System data not available.")
    burden_df = df.groupby('System').agg(
        total_duration=('Duration (mins)', 'sum'),
        avg_interval=('Interval (months)', 'mean')
    ).reset_index()
    burden_df['avg_interval'] = burden_df['avg_interval'].replace(0, np.nan)
    burden_df['burden_score'] = (burden_df['total_duration'] / burden_df['avg_interval']).fillna(0)
    burden_df = burden_df.sort_values('burden_score', ascending=False)
    
    fig = px.bar(
        burden_df, x='burden_score', y='System',
        orientation='h', title='Maintenance Burden Score by System',
        labels={'System': 'System', 'burden_score': 'Burden Score (Higher is more effort)'},
        template='plotly_white'
    )
    fig.update_layout(yaxis={'categoryorder':'total ascending'})
    return fig, burden_df

def create_system_category_breakdown_chart(df):
    """Grouped bar chart showing category breakdown for each system."""
    if 'System' not in df.columns: return go.Figure().update_layout(title_text="System data not available.")
    system_category_df = df.groupby(['System', 'Category of PM check'])['Duration (mins)'].sum().reset_index()
    fig = px.bar(
        system_category_df, x='System', y='Duration (mins)',
        color='Category of PM check', barmode='group',
        title='System Maintenance Profile by Category',
        labels={'System': 'System', 'Duration (mins)': 'Total Duration (Minutes)', 'Category of PM check': 'Category'},
        template='plotly_white'
    )
    return fig

def create_yearly_workload_chart(df):
    """Calculates and plots the total maintenance hours for each month of the year."""
    df_workload = df[df['Interval (months)'] > 0].copy()
    monthly_hours = {i: 0 for i in range(1, 13)}
    
    for _, row in df_workload.iterrows():
        interval = int(row['Interval (months)'])
        duration_hours = row['Duration (mins)'] / 60
        for month in range(1, 13, interval):
            monthly_hours[month] += duration_hours
            
    workload_df = pd.DataFrame(list(monthly_hours.items()), columns=['Month', 'Total Hours'])
    month_names = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'}
    workload_df['Month'] = workload_df['Month'].map(month_names)

    fig = px.bar(
        workload_df, x='Month', y='Total Hours',
        title='Projected Annual Maintenance Workload',
        labels={'Month': 'Month', 'Total Hours': 'Total Scheduled Hours'},
        template='plotly_white'
    )
    return fig

def generate_task_schedule(df, start_date):
    """Generates a realistic, non-consecutive task schedule spread across a 60-day window."""
    if df.empty or df['Duration (mins)'].sum() == 0:
        return pd.DataFrame()

    schedule_df = df.sort_values(
        by=['System', 'Interval (months)', 'Duration (mins)'],
        ascending=[True, True, False]
    ).copy()
    schedule_df.columns = [col.replace(' ', '_').replace('(', '').replace(')', '') for col in schedule_df.columns]

    schedule = []
    daily_work_minutes = 300
    total_work_minutes = schedule_df['Duration_mins'].sum()
    
    num_workdays_needed = int(np.ceil(total_work_minutes / daily_work_minutes))
    
    all_available_weekdays = []
    for i in range(90):
        if len(all_available_weekdays) >= 60:
            break
        current_date = start_date + timedelta(days=i)
        if current_date.weekday() < 5:
            all_available_weekdays.append(current_date)
    
    if num_workdays_needed > 0 and len(all_available_weekdays) > 0:
        days_to_pick = min(num_workdays_needed, len(all_available_weekdays))
        indices = np.linspace(0, len(all_available_weekdays) - 1, days_to_pick, dtype=int)
        scheduled_dates = [all_available_weekdays[i] for i in indices]
    else:
        scheduled_dates = []

    if not scheduled_dates:
        return pd.DataFrame()

    date_iterator = iter(scheduled_dates)
    current_scheduled_date = next(date_iterator)
    time_available_today = daily_work_minutes
    start_of_work_time = timedelta(hours=16)

    tasks_to_schedule = list(schedule_df.itertuples(index=False))

    for task in tasks_to_schedule:
        duration = task.Duration_mins
        if duration == 0:
            continue

        while duration > time_available_today:
            try:
                current_scheduled_date = next(date_iterator)
                time_available_today = daily_work_minutes
            except StopIteration:
                st.warning("Total task duration exceeds the capacity of the scheduling window. Some tasks were not scheduled.")
                return pd.DataFrame(schedule)

        start_offset = timedelta(minutes=(daily_work_minutes - time_available_today))
        task_start_time = datetime.combine(current_scheduled_date, datetime.min.time()) + start_of_work_time + start_offset
        task_end_time = task_start_time + timedelta(minutes=duration)

        schedule.append({
            'Date': current_scheduled_date,
            'Task': task.Task_Description,
            'Start': task_start_time,
            'Finish': task_end_time,
            'System': task.System,
            'Duration (mins)': duration,
            'Page Number': task.Page_Number if hasattr(task, 'Page_Number') else 'N/A'
        })
        
        time_available_today -= duration

    return pd.DataFrame(schedule)


def create_gantt_chart(schedule_df):
    """Creates a Gantt chart from the schedule DataFrame."""
    if schedule_df.empty:
        return go.Figure().update_layout(title_text="No tasks to schedule. Check filters or data.")
    
    plot_df = schedule_df.copy()
    
    ref_date = datetime(2000, 1, 1)
    plot_df['Plot_Start'] = plot_df['Start'].apply(lambda dt: ref_date + (dt - dt.replace(hour=0, minute=0, second=0, microsecond=0)))
    plot_df['Plot_Finish'] = plot_df['Finish'].apply(lambda dt: ref_date + (dt - dt.replace(hour=0, minute=0, second=0, microsecond=0)))
    
    num_days = plot_df['Date'].nunique()
    height = max(400, num_days * 40 + 150)
        
    fig = px.timeline(
        plot_df,
        x_start="Plot_Start",
        x_end="Plot_Finish",
        y="Date",
        color="System",
        hover_name="Task",
        title="Proposed PM Task Schedule",
        labels={"Date": "Scheduled Date"}
    )
    fig.update_yaxes(autorange="reversed", tickformat='%Y-%m-%d')
    fig.update_xaxes(tickformat='%H:%M', title_text='Time of Day (4 PM - 9 PM)')
    
    fig.update_layout(height=height)
    return fig

def display_daily_agenda(schedule_df):
    """Displays the schedule as a day-by-day agenda in cards."""
    if schedule_df.empty:
        st.write("No tasks to display in the agenda.")
        return
    
    st.markdown("""
    <style>
    .task-card {
        border: 1px solid #e6e6e6;
        border-radius: 10px;
        padding: 15px;
        margin-bottom: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .task-time {
        font-weight: bold;
        color: #007bff;
    }
    .task-system {
        font-style: italic;
        color: #555;
    }
    .task-page {
        font-size: 0.9em;
        color: #888;
    }
    </style>
    """, unsafe_allow_html=True)
        
    st.write("### Daily Agenda View")
    for date in sorted(schedule_df['Date'].unique()):
        day_str = pd.to_datetime(date).strftime('%A, %B %d, %Y')
        st.subheader(day_str)
        
        day_tasks = schedule_df[schedule_df['Date'] == date]
        
        for _, task in day_tasks.iterrows():
            start_time = task['Start'].strftime('%I:%M %p')
            end_time = task['Finish'].strftime('%I:%M %p')
            page_number = task.get('Page Number', 'N/A')
            
            card_html = f"""
            <div class="task-card">
                <div class="task-time">{start_time} - {end_time} ({task['Duration (mins)']} mins)</div>
                <div class="task-system">System: {task['System']}</div>
                <div>Task: {task['Task']}</div>
                <div class="task-page">Ref. Page: {page_number}</div>
            </div>
            """
            st.markdown(card_html, unsafe_allow_html=True)

def to_csv(df):
    """Converts a dataframe to a CSV string for downloading."""
    return df.to_csv(index=False).encode('utf-8')

def generate_ics_file(schedule_df):
    """Generates an iCalendar (.ics) file from the schedule dataframe."""
    cal = Calendar()
    for _, task in schedule_df.iterrows():
        event = Event()
        event.add('summary', f"PM Task: {task['Task']}")
        event.add('dtstart', task['Start'])
        event.add('dtend', task['Finish'])
        event.add('description', f"System: {task['System']}\nDuration: {task['Duration (mins)']} minutes")
        cal.add_component(event)
    return cal.to_ical()

# --- PDF Report Generation ---
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'PM Task Analysis Report', 0, 1, 'C')
        self.ln(10)
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')
    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, 0, 1, 'L')
        self.ln(5)
    def chapter_body(self, text):
        safe_text = text.encode('latin-1', 'replace').decode('latin-1')
        self.set_font('Arial', '', 10)
        self.multi_cell(0, 5, safe_text)
        self.ln()
    def add_plotly_chart(self, fig, chart_name):
        chart_path = f"{chart_name}.png"
        fig.write_image(chart_path, width=800, height=500)
        img_width = 180
        x_pos = (210 - img_width) / 2
        self.image(chart_path, x=x_pos, w=img_width)
        self.ln(5)
        os.remove(chart_path)
    def add_agenda_to_pdf(self, schedule_df):
        if schedule_df.empty:
            return
        self.add_page()
        self.chapter_title("Daily Agenda")
        for date in sorted(schedule_df['Date'].unique()):
            day_str = pd.to_datetime(date).strftime('%A, %B %d, %Y')
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, day_str, 0, 1, 'L')
            self.ln(2)
            
            day_tasks = schedule_df[schedule_df['Date'] == date]
            for _, task in day_tasks.iterrows():
                # Clean text before processing
                system_text = f"System: {task['System']}".encode('latin-1', 'replace').decode('latin-1')
                task_text = f"Task: {task['Task']}".encode('latin-1', 'replace').decode('latin-1')
                page_text = f"Ref. Page: {task.get('Page Number', 'N/A')}".encode('latin-1', 'replace').decode('latin-1')
                
                # Calculate card height dynamically
                start_y = self.get_y()
                self.set_font('Arial', '', 10)
                self.multi_cell(self.w - self.l_margin - self.r_margin - 10, 5, task_text)
                text_height = self.get_y() - start_y
                card_height = text_height + 25 # Add padding for other lines
                self.set_y(start_y) # Reset Y position to draw the card

                # Draw card
                self.set_fill_color(245, 245, 245)
                self.set_draw_color(220, 220, 220)
                self.set_line_width(0.2)
                self.cell(0, card_height, '', 1, 1, 'L', fill=True)
                self.set_y(start_y + 5) # Padding
                self.set_x(self.l_margin + 5)

                # Write content inside the card
                start_time = task['Start'].strftime('%I:%M %p')
                end_time = task['Finish'].strftime('%I:%M %p')
                
                self.set_font('Arial', 'B', 10)
                self.set_text_color(0, 123, 255)
                self.cell(0, 6, f"{start_time} - {end_time} ({task['Duration (mins)']} mins)", 0, 1)
                self.set_x(self.l_margin + 5)
                
                self.set_font('Arial', 'I', 9)
                self.set_text_color(85, 85, 85)
                self.cell(0, 6, system_text, 0, 1)
                self.set_x(self.l_margin + 5)
                
                self.set_font('Arial', '', 10)
                self.set_text_color(0, 0, 0)
                self.multi_cell(self.w - self.l_margin - self.r_margin - 10, 5, task_text, 0, 'L')
                self.set_x(self.l_margin + 5)
                
                self.set_font('Arial', '', 8)
                self.set_text_color(136, 136, 136)
                self.cell(0, 6, page_text, 0, 1)

                self.set_y(start_y + card_height + 5)

def generate_pdf_report(figs, suggestions_text, schedule_df):
    """Generates a PDF report with all the charts, suggestions, and agenda."""
    pdf = PDF()
    pdf.add_page()
    for title, fig in figs.items():
        pdf.chapter_title(title)
        pdf.add_plotly_chart(fig, title.replace(" ", "_").replace("/", "_"))
        pdf.ln(10)
    
    pdf.add_page()
    pdf.chapter_title("Suggestions for Shortening PM Duration")
    pdf.chapter_body(suggestions_text)
    
    pdf.add_agenda_to_pdf(schedule_df)
    
    pdf_output = pdf.output(dest='S').encode('latin1')
    b64 = base64.b64encode(pdf_output)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="pm_analysis_report.pdf">Download PDF Report</a>'

def generate_agenda_pdf(schedule_df):
    """Generates a standalone PDF of just the daily agenda."""
    pdf = PDF()
    pdf.add_agenda_to_pdf(schedule_df)
    return pdf.output(dest='S').encode('latin1')

# --- Streamlit App UI ---
st.title("📊 PM Task Analysis Dashboard")
st.write("Upload your PM tasks CSV file to generate insights and visualizations.")

uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

if uploaded_file is None:
    st.info("Please upload a CSV file to begin analysis.")
else:
    df = load_data(uploaded_file)
    if df is not None:
        
        st.sidebar.header("Filters")
        intervals = sorted(df['Interval (months)'].unique())
        selected_intervals = st.sidebar.multiselect("Filter by PM Interval", options=intervals, default=intervals)
        
        selected_systems = []
        if 'System' in df.columns:
            systems = sorted(df['System'].unique())
            selected_systems = st.sidebar.multiselect("Filter by System", options=systems, default=systems)

        selected_categories = []
        if 'Category of PM check' in df.columns:
            categories = sorted(df['Category of PM check'].unique())
            selected_categories = st.sidebar.multiselect("Filter by Category", options=categories, default=categories)

        st.sidebar.header("Scheduling")
        start_date = st.sidebar.date_input("Select PM Start Date", datetime.now())

        filtered_df = df
        if selected_intervals: filtered_df = filtered_df[filtered_df['Interval (months)'].isin(selected_intervals)]
        if selected_systems and 'System' in filtered_df.columns: filtered_df = filtered_df[filtered_df['System'].isin(selected_systems)]
        if selected_categories and 'Category of PM check' in filtered_df.columns: filtered_df = filtered_df[filtered_df['Category of PM check'].isin(selected_categories)]

        st.header("🔍 Data Overview")
        st.write("Metrics based on your current filter selection.")
        create_summary_metrics(filtered_df)
        
        st.header("📊 Visualizations")
        
        fig_cat_dur = create_category_duration_chart(filtered_df)
        fig_task_count = create_task_count_chart(filtered_df)
        fig_int_cat = create_interval_category_breakdown_chart(filtered_df)
        fig_sys_dur = create_system_duration_chart(filtered_df)
        fig_hierarchical = create_hierarchical_chart(filtered_df)
        fig_longest_tasks = create_longest_tasks_chart(filtered_df)
        fig_burden, burden_df = create_maintenance_burden_chart(filtered_df)
        fig_sys_cat_breakdown = create_system_category_breakdown_chart(filtered_df)
        fig_workload = create_yearly_workload_chart(filtered_df)
        
        schedule_df = generate_task_schedule(filtered_df, start_date)
        fig_gantt = create_gantt_chart(schedule_df)
        
        tab1, tab2, tab3, tab4 = st.tabs(["Category & Interval Analysis", "System Analysis", "Advanced Insights & Planning", "Task Scheduling"])

        with tab1:
            col1, col2 = st.columns(2)
            with col1: st.plotly_chart(fig_cat_dur, use_container_width=True)
            with col2: st.plotly_chart(fig_task_count, use_container_width=True)
            st.plotly_chart(fig_int_cat, use_container_width=True)

        with tab2:
            col1, col2 = st.columns(2)
            with col1: st.plotly_chart(fig_sys_dur, use_container_width=True)
            with col2: st.plotly_chart(fig_hierarchical, use_container_width=True)
            st.plotly_chart(fig_sys_cat_breakdown, use_container_width=True)
            
        with tab3:
            col1, col2 = st.columns(2)
            with col1: 
                st.plotly_chart(fig_longest_tasks, use_container_width=True)
            with col2: 
                st.plotly_chart(fig_burden, use_container_width=True)
                with st.expander("How is Burden Score calculated?"):
                    st.markdown("""
                    The **Maintenance Burden Score** gives a quick overview of which systems require the most overall effort, factoring in both how long the tasks take and how often they need to be done.

                    **Calculation Breakdown:**
                    1.  **Group by System:** First, the script groups all PM tasks by their 'System' name.
                    2.  **Calculate Totals for Each System:** For each system, it calculates two key numbers:
                        * **Total Duration:** The sum of the 'Duration (mins)' for all tasks related to that system.
                        * **Average Interval:** The *average* of the 'Interval (months)' for all tasks for that system.
                    3.  **Calculate the Score:** `Burden Score = Total Duration / Average Interval`
                    A system with a high total duration and a low average interval will have a much higher burden score.
                    """)
            st.plotly_chart(fig_workload, use_container_width=True)

        with tab4:
            st.plotly_chart(fig_gantt, use_container_width=True)
            
            display_daily_agenda(schedule_df)
            
            if not schedule_df.empty:
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.download_button(
                       label="Download Schedule as CSV",
                       data=to_csv(schedule_df),
                       file_name='pm_schedule.csv',
                       mime='text/csv',
                    )
                with col2:
                    st.download_button(
                        label="Download Agenda as PDF",
                        data=generate_agenda_pdf(schedule_df),
                        file_name="pm_agenda.pdf",
                        mime="application/pdf"
                    )
                with col3:
                    st.download_button(
                        label="Download for Calendar (.ics)",
                        data=generate_ics_file(schedule_df),
                        file_name="pm_schedule.ics",
                        mime="text/calendar"
                    )

            st.write("### Suggestions for Shortening PM Duration")
            
            suggestions_text = ""
            if not burden_df.empty and not filtered_df.empty:
                top_burden_systems = burden_df.nlargest(3, 'burden_score')['System'].tolist()
                top_longest_tasks = filtered_df.nlargest(3, 'Duration (mins)')['Task Description'].tolist()
                
                suggestions_text = f"""
                **1. Focus on High-Burden Systems:**
                The systems with the highest 'Maintenance Burden Score' are **{', '.join(top_burden_systems)}**. Optimizing procedures for these systems will yield the biggest time savings. Consider reviewing their specific tasks for potential efficiencies.

                **2. Review the Longest Tasks:**
                The most time-consuming individual tasks are often the best candidates for process improvement. The top longest tasks in your current selection are:
                - {top_longest_tasks[0] if len(top_longest_tasks) > 0 else 'N/A'}
                - {top_longest_tasks[1] if len(top_longest_tasks) > 1 else 'N/A'}
                - {top_longest_tasks[2] if len(top_longest_tasks) > 2 else 'N/A'}
                Could special tooling, pre-kitting parts, or assigning a second engineer shorten these specific procedures?

                **3. Parallelize Work Where Possible:**
                The generated schedule groups tasks by system to minimize context switching. If multiple engineers are available, consider assigning them to different systems on the same day to perform work in parallel. For example, one engineer could work on 'LINAC' tasks while another works on 'XVI' tasks.

                **4. Pre-Task Preparation:**
                Before each scheduled maintenance day, ensure all necessary tools, parts, and documentation are prepared and staged. This minimizes downtime searching for resources during the limited 4 PM to 9 PM work window.

                **5. Data-Driven Interval Review:**
                For systems that consistently show high reliability and have no history of failures, it may be worthwhile to discuss with the manufacturer whether certain low-impact task intervals can be safely extended. This is a long-term strategy that should be approached with caution and expert consultation.
                """
            st.markdown(suggestions_text)

        st.header("⬇️ Download Report")
        st.write("Click the button below to download all charts (based on current filters) in a single PDF report.")
        
        figs_to_download = {
            "Total Maintenance Duration by Category": fig_cat_dur,
            "Distribution of Tasks by Category": fig_task_count,
            "Duration Breakdown by Interval/Category": fig_int_cat,
            "Total Maintenance Duration by System": fig_sys_dur,
            "Hierarchical View of Maintenance Duration": fig_hierarchical,
            "Top 10 Longest Tasks": fig_longest_tasks,
            "Maintenance Burden Score": fig_burden,
            "System Maintenance Profile": fig_sys_cat_breakdown,
            "Annual Maintenance Workload": fig_workload,
            "Proposed PM Task Schedule": fig_gantt
        }
        
        if st.button("Generate PDF Report"):
            with st.spinner("Generating PDF..."):
                download_link = generate_pdf_report(figs_to_download, suggestions_text, schedule_df)
                st.success("Report generated successfully!")
                st.markdown(download_link, unsafe_allow_html=True)

        st.header("📋 Detailed Task Data (Filtered)")
        st.dataframe(filtered_df)
