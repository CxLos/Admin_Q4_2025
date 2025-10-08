# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
from datetime import datetime
import folium
import os
import sys
from collections import Counter
# -------------------------------
import requests
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# --------------------------------
import dash
from dash import dcc, html, Input, Output, State, dash_table
from dash.development.base_component import Component

# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)

# ============================= System Information ============================= #

current_dir = os.getcwd()
script_dir = os.path.dirname(os.path.abspath(__file__))

# Report Name
file_name = os.path.basename(__file__)  # Gets "admin_q4_25.py"
file_parts = file_name.replace('.py', '').split('_')  # ['admin', 'q4', '25']
report = file_parts[0].capitalize() if file_parts else 'Admin'  # Gets 'Admin'
# print(f"Report: {report}")

# ============================== Date Info ============================== #

report_date = datetime(2025, 7, 1) 
month = report_date.month
report_year = report_date.year

# Get the reporting quarter:
def get_custom_quarter(date_obj):
    month = date_obj.month
    if month in [10, 11, 12]:
        return "Q1"  # October–December
    elif month in [1, 2, 3]:
        return "Q2"  # January–March
    elif month in [4, 5, 6]:
        return "Q3"  # April–June
    elif month in [7, 8, 9]:
        return "Q4"  # July–September

# Adjust the quarter calculation for custom quarters
if month in [10, 11, 12]:
    quarter = 1  # Q1: October–December
elif month in [1, 2, 3]:
    quarter = 2  # Q2: January–March
elif month in [4, 5, 6]:
    quarter = 3  # Q3: April–June
elif month in [7, 8, 9]:
    quarter = 4  # Q4: July–September

# Define a mapping for months to their corresponding quarter
quarter_months = {
    1: ['October', 'November', 'December'],  # Q1
    2: ['January', 'February', 'March'],    # Q2
    3: ['April', 'May', 'June'],            # Q3
    4: ['July', 'August', 'September']      # Q4
}

# Get the months for the current quarter
months_in_quarter = quarter_months[quarter]

# Calculate start and end month indices for the quarter
all_months = [
    'January', 'February', 'March', 
    'April', 'May', 'June',
    'July', 'August', 'September', 
    'October', 'November', 'December'
]
start_month_idx = (quarter - 1) * 3
month_order = all_months[start_month_idx:start_month_idx + 3]

current_quarter = get_custom_quarter(report_date)
# print(f"Reporting Quarter: {current_quarter}")

# ================================== Load Data ================================= #

# Define the Google Sheets URL
sheet_url = "https://docs.google.com/spreadsheets/d/1xZ-OulU-SOfd6jraH2fEvvVdbSXIUOg-RA3PKZHP_GQ/edit?gid=0#gid=0"

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials
encoded_key = os.getenv("GOOGLE_CREDENTIALS")

if encoded_key:
    json_key = json.loads(base64.b64decode(encoded_key).decode("utf-8"))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
else:
    creds_path = r"C:\Users\CxLos\OneDrive\Documents\BMHC\Data\bmhc-timesheet-4808d1347240.json"
    if os.path.exists(creds_path):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    else:
        raise FileNotFoundError("Service account JSON file not found and GOOGLE_CREDENTIALS is not set.")

# Authorize and load the sheet
client = gspread.authorize(creds)
sheet = client.open_by_url(sheet_url)
worksheet = sheet.worksheet(f"{current_quarter}")
data = pd.DataFrame(worksheet.get_all_records())
df = data.copy()

# Strip whitespace from columns and cell values
df.columns = df.columns.str.strip()
df = df.apply(
        lambda col: col.str.strip() if col.dtype == "object" or pd.api.types.is_string_dtype(col) else col
    )

# Define a discrete color sequence
color_sequence = px.colors.qualitative.Plotly

# Get the reporting month:
df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')
df = df.sort_values(by='Start Date', ascending=True)
df['Month'] = df['Start Date'].dt.month_name()

# -------------------------------------------------
# print(df.head())
# print(df[["Date of Activity", "Total travel time (minutes):"]])
# print('Total Marketing Events: ', len(df))
# print('Column Names: \n', df.columns.tolist())
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())
# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

columns =  [
    'Client',
    'Project', 
    'Task', 
    'Kiosk', 
    'User', 
    'Group',
    'Tags', 
    'Description',
    'Collaborated Entity', 
    '# of People Engaged', 
    'Duration (h)',
    'Total Travel Time',
    # ----------------------
    'Email', 
    'Billable', 
    'Start Date',
    'Start Time', 
    'End Date', 
    'End Time', 
    'Duration (h)',
    'Duration (decimal)', 
    'Billable Rate (USD)', 
    'Billable Amount (USD)',
]

# df = df[columns]
df = df[df['Project'] == 'BMHC Administrative Activity']
# print(df.head())

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

#  Please provide public information:    137
# Please explain event-oriented:        13

# ============================== Data Preprocessing ========================== #

# Rename columns
df.rename(
    columns={
        "Client": "Client",
        "Project": "Project",
        "Kiosk": "Kiosk",
        "Description": "Description",
        # --------------------
        "Duration (h)": "Duration",
        "# of People Engaged": "Engaged",
        "Group": "Group",
        "Task": "Task",
        "Tags": "Tags",
        "User": "User",
        "Collaborated Entity": "Collab",
        # "": "",
    }, 
inplace=True)

# print(df.dtypes)

# =========================== Total Events ============================ #

total_events = len(df)
# print("Total events:", total_events)

events_data = []
for month in months_in_quarter:
    events_in_month = df[df['Month'] == month].shape[0]  # Count the number of rows for each month
    events_data.append({
        'Month': month,
        'Total Events': events_in_month
    })

# Create DataFrame for quarterly events data
df_events_quarterly = pd.DataFrame(events_data)

# Get overall events distribution for pie chart
df_events = df_events_quarterly.copy()

# Total Events Bar Chart - Quarterly Format
events_bar = px.bar(
    df_events_quarterly,
    x='Month',
    y='Total Events',
    color='Month',
    text='Total Events',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    title=dict(
        text=f'{current_quarter} Total Events by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>Total Events</b>: %{y}<extra></extra>'
)

# Total Events Pie Chart - Overall Distribution
events_pie = px.pie(
    df_events,
    names='Month',
    values='Total Events',
    color='Month'
).update_layout(
    title=dict(
        text=f'{current_quarter} Total Events Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=180,
    textfont=dict(size=16),
    texttemplate='%{value}<br>(%{percent:.1%})',
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# =========================== Total Events ============================ #

total_events = len(df)
# print("Total events:", total_events)

events_data = []
for month in months_in_quarter:
    events_in_month = df[df['Month'] == month].shape[0]  # Count the number of rows for each month
    events_data.append({
        'Month': month,
        'Total Events': events_in_month
    })

# Create DataFrame for quarterly events data
df_events_quarterly = pd.DataFrame(events_data)

# Get overall events distribution for pie chart
df_events = df_events_quarterly.copy()

# Total Events Bar Chart - Quarterly Format
events_bar = px.bar(
    df_events_quarterly,
    x='Month',
    y='Total Events',
    color='Month',
    text='Total Events',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    title=dict(
        text=f'{current_quarter} Total {report} Events by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>Total Events</b>: %{y}<extra></extra>'
)

# Total Events Pie Chart - Overall Distribution
events_pie = px.pie(
    df_events,
    names='Month',
    values='Total Events',
    color='Month'
).update_layout(
    title=dict(
        text=f'{current_quarter} Total {report} Events Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=180,
    textfont=dict(size=16),
    texttemplate='%{value}<br>(%{percent:.1%})',
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# =========================== Total Hours ============================ #

# print("Duration Unique Before:", df['Duration'].unique().tolist())

# Convert duration strings (HH:MM:SS) to timedelta
df['Duration'] = pd.to_timedelta(df['Duration'], errors='coerce')
df['Duration'] = df['Duration'].dt.total_seconds() / 3600
df['Duration'] = pd.to_numeric(df['Duration'], errors='coerce')

# print("Duration Unique After:", df['Duration'].unique().tolist())

total_hours = df['Duration'].sum()
total_hours = round(total_hours)
# print('Total Activity Duration:', total_hours, 'hours')

# Calculate total hours per month
hours_data = []
for month in months_in_quarter:
    hours_in_month = df[df['Month'] == month]['Duration'].sum()
    hours_data.append({
        'Month': month,
        'Total Hours': round(hours_in_month, 1)
    })

# Create DataFrame for quarterly hours data
df_hours_quarterly = pd.DataFrame(hours_data)

# Get overall hours distribution for pie chart
df_hours = df_hours_quarterly.copy()

# Total Hours Bar Chart - Quarterly Format
hours_bar = px.bar(
    df_hours_quarterly,
    x='Month',
    y='Total Hours',
    color='Month',
    text='Total Hours',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Hours',
    title=dict(
        text=f'{current_quarter} Total {report} Hours by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>Total Hours</b>: %{y}<extra></extra>'
)

# Total Hours Pie Chart - Overall Distribution
hours_pie = px.pie(
    df_hours,
    names='Month',
    values='Total Hours',
    color='Month'
).update_layout(
    title=dict(
        text=f'{current_quarter} Total {report} Hours Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=180,
    textfont=dict(size=16),
    texttemplate='%{value}<br>(%{percent:.1%})',
    hovertemplate='<b>%{label}</b>: %{value} hours<extra></extra>'
)

# =========================== Total Engaged ========================== #

# print("Engaged Unique before:", df['Engaged'].unique().tolist())

df['Engaged'] = df['Engaged'].fillna('0')

engaged_unique = [
    'Between 1 and 10', 
    'None', 
    '', 
    'Between 11 and 20', 
    'Between 20 and 30'
]

df['Engaged'] = (
    df['Engaged']
        .astype(str)
        .str.strip()
            .replace({
                '': '0',
                'None': '0',
                'Between 1 and 10': '10',
                'Between 11 and 20': '20',
                'Between 20 and 30': '30',
                # '': '',
            })
)
# print("Engaged Unique after:", df['Engaged'].unique().tolist())

df['Engaged'] = pd.to_numeric(df['Engaged'], errors='coerce')

df_engaged = df['Engaged'].sum()
# print('Total Engaged:', df_engaged)

# Calculate total people engaged per month
engaged_data = []
for month in months_in_quarter:
    engaged_in_month = df[df['Month'] == month]['Engaged'].sum()
    engaged_data.append({
        'Month': month,
        'Total Engaged': int(engaged_in_month)
    })

# Create DataFrame for quarterly engaged data
df_engaged_quarterly = pd.DataFrame(engaged_data)

# Get overall engaged distribution for pie chart
df_engaged_chart = df_engaged_quarterly.copy()

# Total Engaged Bar Chart - Quarterly Format
engaged_bar = px.bar(
    df_engaged_quarterly,
    x='Month',
    y='Total Engaged',
    color='Month',
    text='Total Engaged',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='People',
    title=dict(
        text=f'{current_quarter} People Engaged by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>People Engaged</b>: %{y}<extra></extra>'
)

# Total Engaged Pie Chart - Overall Distribution
engaged_pie = px.pie(
    df_engaged_chart,
    names='Month',
    values='Total Engaged',
    color='Month'
).update_layout(
    title=dict(
        text=f'{current_quarter} People Engaged Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=180,
    textfont=dict(size=16),
    texttemplate='%{value}<br>(%{percent:.1%})',
    hovertemplate='<b>%{label}</b>: %{value} people<extra></extra>'
)

# ====================== Admin Group ====================== #

# print("Group Unique before:", df['Group'].unique().tolist())

df['Group'] = (
    df['Group']
        .astype(str)
            .str.strip()
            .replace({
                "" : "N/A",
            })
    )

group_categories = [
    'Coordination & Navigation', 
    'Information Technology', 
    'Outreach & Engagement',
    'Permanent Supportive Housing',
    'Administration',
    'Communications',
    'Marketing',
]

group_normalized = {cat.lower().strip(): cat for cat in group_categories}

# Create monthly group data for quarterly view
group_monthly_data = []

for month in months_in_quarter:
    month_df = df[df['Month'] == month]
    month_counter = Counter()
    
    for entry in month_df['Group']:
        # Split and clean each category
        items = [i.strip().lower() for i in entry.split(",")]
        for item in items:
            if item in group_normalized:
                month_counter[group_normalized[item]] += 1
    
    # Convert to DataFrame and add month column
    for group_type, count in month_counter.items():
        group_monthly_data.append({
            'Month': month,
            'Group': group_type,
            'Count': count
        })

# Create DataFrame for quarterly group data
df_group_quarterly = pd.DataFrame(group_monthly_data)

# Overall group data (for pie chart)
counter = Counter()
for entry in df['Group']:
    items = [i.strip().lower() for i in entry.split(",")]
    for item in items:
        if item in group_normalized:
            counter[group_normalized[item]] += 1

df_group = pd.DataFrame(counter.items(), columns=['Group', 'Count']).sort_values(by='Count', ascending=False)

# Group Bar Chart - Quarterly Format
group_bar = px.bar(
    df_group_quarterly,
    x='Month',
    y='Count',
    color='Group',
    barmode='group',
    text='Count',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    title=dict(
        text=f'{current_quarter} {report} Groups by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    bargap=0.08,
    bargroupgap=0,
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>Group</b>: %{fullData.name}<br><b>Count</b>: %{y}<extra></extra>'
)

# Group Pie Chart - Overall Distribution
group_pie = px.pie(
    df_group,
    names="Group",
    values='Count' 
).update_layout(
    title=dict(
        text=f'{current_quarter} Ratio of {report} Groups',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=80,
    textfont=dict(size=16),
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# ======================== Admin Task ========================= #

# print("Task Unique before:", df['Task'].unique().tolist())

df['Task'] = (
    df['Task']
        .astype(str)
            .str.strip()
            .replace({
                "Newsletter - writing, editing, proofing" : "Newsletter",
                "" : "N/A",
                # "" : "",
            })
)

# print("Task Unique after:", df['Task'].unique().tolist())

task_categories = [
'Communication & Correspondence',
'HR Support', 
'Research & Planning', 
'Key Event', 
'Data Archiving',
'General Maintenance',
'Record Keeping & Documentation',
'Desk Help Support', 
'Workforce Development', 
'Academic',
'Content, Line Editing, or Proofing', 
'Financial & Budgetary Management',
'Device Management', 
'Training',
'Social Media/YouTube', 
'Compliance & Policy Enforcement', 
'Health Education or Awareness', 
'Clinical Provider', 
'Website or Intranet Updates', 
'Field Outreach', 
'Tabling',
'Advocacy Partner',
'Board Support', 
'', 
'Selfcare Healing', 
'SDoH Provider', 
'Travel',
'Team Meeting',
'Newsletter - writing, editing, proofing', 
'Office Management'
]

# Calculate task distribution per month
task_data = []
for month in months_in_quarter:
    month_data = df[df['Month'] == month]['Task'].value_counts().reset_index()
    month_data.columns = ['Task', 'Count']
    month_data['Month'] = month
    task_data.append(month_data)

# Combine all months
df_task_quarterly = pd.concat(task_data, ignore_index=True)

# Get overall task distribution for pie chart
df_task = df['Task'].value_counts().reset_index(name='Count').sort_values(by='Count', ascending=False)

# Task Bar Chart - Quarterly Format
task_bar = px.bar(
    df_task_quarterly,
    x='Month',
    y='Count',
    color='Task',
    barmode='group',
    text='Count',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    title=dict(
        text=f'{current_quarter} {report} Tasks by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    bargap=0.08,
    bargroupgap=0,
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>Task</b>: %{fullData.name}<br><b>Count</b>: %{y}<extra></extra>'
)

# Task Pie Chart - Overall Distribution
task_pie = px.pie(
    df_task,
    names="Task",
    values='Count' 
).update_layout(
    title=dict(
        text=f'{current_quarter} {report} Tasks Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=150,
    textfont=dict(size=16),
    texttemplate='%{percent:.1%}',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# --------------------------- Admin Tags -------------------------- #

# print("Tags Unique before:", df['Tags'].unique().tolist())

df['Tags'] = (
    df['Tags']
        .astype(str)
            .str.strip()
            .replace({
                "" : "N/A",
                # "" : "",
            })
    )

tag_categories = [
    "Search icon",
    "Add/Search tags",
    "AmeriCorps Duties",
    "Board Support",
    "Brand Messaging Strategy",
    "Care Network",
    "Data Archiving",
    "Documentation",
    "Email",
    "Equipment",
    "Event Planning",
    "Fundraising",
    "Grant",
    "Graphic and/or Creatives Design",
    "Handout",
    "HealthyCuts",
    "HR Support",
    "Impromptu Discussion",
    "IT",
    "Know Your Numbers",
    "Letter",
    "MarCom Playbook",
    "Materials Review",
    "Meeting",
    "Movement Is Medicine",
    "Newsletter / Announcements",
    "OverComing Mental Hellness",
    "Philanthropy Call",
    "Philanthropy Email",
    "Phone Call",
    "Planned Change",
    "Polls/Surveys",
    "Presentation",
    "Proposal",
    "PSH Work",
    "Public Relations / Press Releases",
    "Recent Change",
    "Research, writing, and editing",
    "Social Media and/or Youtube",
    "Sustainability Binder",
    "Tabling Event",
    "Timesheet / Impact Reporting",
    "Training",
    "Videography",
    "Website"
]

tag_normalized = {cat.lower().strip(): cat for cat in tag_categories}

# Create monthly tag data for quarterly view
tag_monthly_data = []

for month in months_in_quarter:
    month_df = df[df['Month'] == month]
    month_counter = Counter()
    
    for entry in month_df['Tags']:
        # Split and clean each category
        items = [i.strip().lower() for i in entry.split(",")]
        for item in items:
            if item in tag_normalized:
                month_counter[tag_normalized[item]] += 1
    
    # Convert to DataFrame and add month column
    for tag_type, count in month_counter.items():
        tag_monthly_data.append({
            'Month': month,
            'Tags': tag_type,
            'Count': count
        })

# Create DataFrame for quarterly tag data
df_tag_quarterly = pd.DataFrame(tag_monthly_data)

# Overall tag data (for pie chart)
counter = Counter()
for entry in df['Tags']:
    items = [i.strip().lower() for i in entry.split(",")]
    for item in items:
        if item in tag_normalized:
            counter[tag_normalized[item]] += 1

df_tag = pd.DataFrame(counter.items(), columns=['Tags', 'Count']).sort_values(by='Count', ascending=False)

# Tag Bar Chart - Quarterly Format
tag_bar = px.bar(
    df_tag_quarterly,
    x='Month',
    y='Count',
    color='Tags',
    barmode='group',
    text='Count',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    title=dict(
        text=f'{current_quarter} {report} Tags by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    bargap=0.08,
    bargroupgap=0,
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>Tag</b>: %{fullData.name}<br><b>Count</b>: %{y}<extra></extra>'
)

# Tag Pie Chart - Overall Distribution
tag_pie = px.pie(
    df_tag,
    names="Tags",
    values='Count' 
).update_layout(
    title=dict(
        text=f'{current_quarter} {report} Tags Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=180,
    textfont=dict(size=16),
    texttemplate='%{percent:.2%}',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# --------------------------- Collaborated Entity -------------------------- #

# print("Collab Unique before:", df['Collab'].unique().tolist())

df['Collab'] = (
    df['Collab']
        .astype(str)
            .str.strip()
            .replace({
                "" : "N/A",
                # "" : "",
            })
    )

# Normalize collaboration categories (lowercase and stripped for consistency)
collab_normalized = {}

# Create monthly collaboration data for quarterly view
collab_monthly_data = []

for month in months_in_quarter:
    month_df = df[df['Month'] == month]
    month_counter = Counter()
    
    for entry in month_df['Collab']:
        # Split and clean each category
        items = [i.strip() for i in str(entry).split(",") if i.strip() and i.strip().lower() != "n/a"]
        for item in items:
            if item:  # If item is not empty
                month_counter[item] += 1
    
    # Convert to DataFrame and add month column
    for collab_type, count in month_counter.items():
        collab_monthly_data.append({
            'Month': month,
            'Collab': collab_type,
            'Count': count
        })

# Create DataFrame for quarterly collaboration data
df_collab_quarterly = pd.DataFrame(collab_monthly_data)

# Overall collaboration data (for pie chart)
counter = Counter()
for entry in df['Collab']:
    items = [i.strip() for i in str(entry).split(",") if i.strip() and i.strip().lower() != "n/a"]
    counter.update(items)

df_collab = pd.DataFrame(counter.items(), columns=['Collab', 'Count']).sort_values(by='Count', ascending=False)

# Collaborated Entity Bar Chart - Quarterly Format
collab_bar = px.bar(
    df_collab_quarterly,
    x='Month',
    y='Count',
    color='Collab',
    barmode='group',
    text='Count',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    title=dict(
        text=f'{current_quarter} Collaborated Entities by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    bargap=0.08,
    bargroupgap=0,
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>Entity</b>: %{fullData.name}<br><b>Count</b>: %{y}<extra></extra>'
)

# Collaborated Entity Pie Chart - Overall Distribution
collab_pie = px.pie(
    df_collab,
    names="Collab",
    values='Count' 
).update_layout(
    title=dict(
        text=f'{current_quarter} Collaborated Entities Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ) 
).update_traces(
    rotation=90,
    textfont=dict(size=16),
    texttemplate='%{percent:.1%}',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# --------------------------- Admin Users -------------------------- #

# print("User Unique before:", df['User'].unique().tolist())

user_unique = [
'larrywallace.jr', 'Coby Albrecht', 'kiounis williams', 'Areebah Mubin', 'jaqueline.oviedo', 'Jordan Calbert', 'Sashricaa Manoj Kumar', 'Eric Roberts', 'pamela.parker', 'Angelita Delagarza', 'lavonne.williams', 'kimberly.holiday', 'Azaniah Israel', 'arianna.williams', 'antonio.montgomery', 'Michael Lambert', 'steve kemgang', 'tramisha.pete', 'toyacraney', 'felicia.chandler', 'Dominique Holman'
]

df['User'] = (
    df['User']
        .astype(str)
            .str.strip()
            .replace({
                "steve kemgang" : "Steve Kemgang",
                "toyacraney" : "Toya Craney",
                "felicia.chandler" : "Felicia Chandler",
                "tramisha.pete" : "Tramisha Pete",
                "jaqueline.oviedo" : "Jaqueline Oviedo",
                "larrywallace.jr" : "Larry Wallace Jr.",
                "kiounis williams" : "Kiounis Williams",
                "pamela.parker" : "Pamela Parker",
                "lavonne.williams" : "Lavonne Williams",
                "kimberly.holiday" : "Kimberly Holiday",
                "antonio.montgomery" : "Antonio Montgomery",
                "arianna.williams" : "Arianna Williams",
                "carlos.bautista" : "Carlos Bautista",
                "christi.freeman" : "Christi Freeman",
                "" : "N/A",
            })
    )

# print("User Unique after:", df['User'].unique().tolist())
# print("User Values:", df['User'].value_counts())

# Calculate user distribution per month
user_data = []
for month in months_in_quarter:
    month_data = df[df['Month'] == month]['User'].value_counts().reset_index()
    month_data.columns = ['User', 'Count']
    month_data['Month'] = month
    user_data.append(month_data)

# Combine all months
df_user_quarterly = pd.concat(user_data, ignore_index=True)

# Get overall user distribution for pie chart
df_user = df['User'].value_counts().reset_index(name='Count').sort_values(by='Count', ascending=False)

# User Bar Chart - Quarterly Format
user_bar = px.bar(
    df_user_quarterly,
    x='Month',
    y='Count',
    color='User',
    barmode='group',
    text='Count',
    category_orders={'Month': months_in_quarter}
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    title=dict(
        text=f'{current_quarter} User Submissions by Month',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    ),
    xaxis=dict(
        title=dict(
            text=None,
            font=dict(size=20),
        ),
        tickmode='array',
        tickvals=months_in_quarter,
        tickangle=0
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    bargap=0.08,
    bargroupgap=0,
).update_traces(
    texttemplate='%{text}',
    textfont=dict(size=16),
    textposition='auto',
    textangle=0,
    hovertemplate='<b>Month</b>: %{x}<br><b>User</b>: %{fullData.name}<br><b>Count</b>: %{y}<extra></extra>'
)

# User Pie Chart - Overall Distribution
user_pie = px.pie(
    df_user,
    names="User",
    values='Count' 
).update_layout(
    title=dict(
        text=f'{current_quarter} User Submissions Ratio',
        x=0.5, 
        font=dict(
            size=21,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=16,
        color='black'
    )
).update_traces(
    rotation=210,
    textfont=dict(size=16),
    texttemplate='%{value}<br>(%{percent:.2%})',
    hovertemplate='%{label}: <b>%{value}</b><br>Percent: <b>%{percent:.2%}</b><extra></extra>',
)

# ========================== DataFrame Table ========================== #

df = df.sort_values('Start Date', ascending=True)

# create a display index column and prepare table data/columns
# reset index to ensure contiguous numbering after any filtering/sorting upstream
df_indexed = df.reset_index(drop=True).copy()
# Insert '#' as the first column (1-based row numbers)
df_indexed.insert(0, '#', df_indexed.index + 1)

# Convert to records for DataTable
data = df_indexed.to_dict('records')
columns = [{"name": col, "id": col} for col in df_indexed.columns]

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server

app.layout = html.Div(
    children=[ 
        html.Div(
            className='divv', 
            children=[ 
                html.H1(
                    f'BMHC Administrative Activity Report', 
                    className='title'),
                html.H1(
                    f'{current_quarter} {report_year}', 
                    className='title2'),
                html.Div(
                    className='btn-box', 
                    children=[
                        html.A(
                            'Repo',
                            href=f'https://github.com/CxLos/{report}_{current_quarter}_{report_year}',
                            className='btn'
                        ),
                    ]
                ),
            ]
        ),  

# ============================ Rollups ========================== #

html.Div(
    className='rollup-row',
    children=[
        
        html.Div(
            className='rollup-box-tl',
            children=[
                html.Div(
                    className='title-box',
                    children=[
                        html.H3(
                            className='rollup-title',
                            children=[f'{current_quarter} {report} Events']
                        ),
                    ]
                ),

                html.Div(
                    className='circle-box',
                    children=[
                        html.Div(
                            className='circle-1',
                            children=[
                                html.H1(
                                className='rollup-number',
                                children=[total_events]
                            ),
                            ]
                        )
                    ],
                ),
            ]
        ),
        html.Div(
            className='rollup-box-tr',
            children=[
                html.Div(
                    className='title-box',
                    children=[
                        html.H3(
                            className='rollup-title',
                            children=[f'{current_quarter} People Engaged']
                        ),
                    ]
                ),
                html.Div(
                    className='circle-box',
                    children=[
                        html.Div(
                            className='circle-2',
                            children=[
                                html.H1(
                                className='rollup-number',
                                children=[df_engaged]
                            ),
                            ]
                        )
                    ],
                ),
            ]
        ),
    ]
),

html.Div(
    className='rollup-row',
    children=[
        html.Div(
            className='rollup-box-bl',
            children=[
                html.Div(
                    className='title-box',
                    children=[
                        html.H3(
                            className='rollup-title',
                            children=[f'{current_quarter} {report} Hours']
                        ),
                    ]
                ),

                html.Div(
                    className='circle-box',
                    children=[
                        html.Div(
                            className='circle-3',
                            children=[
                                html.H1(
                                className='rollup-number',
                                children=[total_hours]
                            ),
                            ]
                        )
                    ],
                ),
            ]
        ),
        html.Div(
            className='rollup-box-br',
            children=[
                html.Div(
                    className='title-box',
                    children=[
                        html.H3(
                            className='rollup-title',
                            children=[f'Placeholder']
                        ),
                    ]
                ),
                html.Div(
                    className='circle-box',
                    children=[
                        html.Div(
                            className='circle-4',
                            children=[
                                html.H1(
                                className='rollup-number',
                                children=['-']
                                ),
                            ]
                        )
                    ],
                ),
            ]
        ),
    ]
),

# ============================ Visuals ========================== #

html.Div(
    className='graph-container',
    children=[
        
        html.H1(
            className='visuals-text',
            children='Visuals'
        ),

        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=events_bar
                        )
                    ]
                ),
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=events_pie
                        )
                    ]
                ),
            ]
        ),

        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=hours_bar
                        )
                    ]
                ),
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=hours_pie
                        )
                    ]
                ),
            ]
        ),

        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=engaged_bar
                        )
                    ]
                ),
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=engaged_pie
                        )
                    ]
                ),
            ]
        ),

        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=group_bar
                        )
                    ]
                ),
                html.Div(
                    className='graph-box',
                    children=[
                        dcc.Graph(
                            className='graph',
                            figure=group_pie
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=task_bar
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=task_pie
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=collab_bar
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=collab_pie
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=tag_bar
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=tag_pie
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=user_bar
                        )
                    ]
                ),
            ]
        ),
        
        html.Div(
            className='graph-row',
            children=[
                html.Div(
                    className='wide-box',
                    children=[
                        dcc.Graph(
                            className='wide-graph',
                            figure=user_pie
                        )
                    ]
                ),
            ]
        ),
    ]
),

# ============================ Data Table ========================== #

    html.Div(
        className='data-box',
        children=[
            html.H1(
                className='data-title',
                children=f'{report} Activity Table'
            ),
            # html.Div(  
            #     className='table-scroll',
            #     children=[
            #         dcc.Graph(
            #             className='data',
            #             figure=df_table,
            #                 # style={'height': '800px'}, 
            #                 config={'responsive': True}
            #         )
            #     ]
            # )
            
            dash_table.DataTable(
                id='applications-table',
                data=data,
                columns=columns,
                page_size=10,
                sort_action='native',
                filter_action='native',
                row_selectable='multi',
                style_table={
                    'overflowX': 'auto',
                    # 'border': '3px solid #000',
                    # 'borderRadius': '0px'
                },
                style_cell={
                    'textAlign': 'left',
                    'minWidth': '100px', 
                    'whiteSpace': 'normal'
                },
                style_header={
                    'textAlign': 'center', 
                    'fontWeight': 'bold',
                    'backgroundColor': '#34A853', 
                    'color': 'white'
                },
                style_data={
                    'whiteSpace': 'normal',
                    'height': 'auto',
                },
                style_cell_conditional=[
                    # make the index column narrow and centered
                    {'if': {'column_id': '#'},
                    'width': '20px', 'minWidth': '60px', 'maxWidth': '60px', 'textAlign': 'center'},

                    {'if': {'column_id': 'Description'},
                    'width': '350px', 'minWidth': '200px', 'maxWidth': '400px'},

                    {'if': {'column_id': 'Tags'},
                    'width': '250px', 'minWidth': '200px', 'maxWidth': '400px'},

                    {'if': {'column_id': 'Collab'},
                    'width': '250px', 'minWidth': '200px', 'maxWidth': '400px'},
                ]
            ),
        ]
    ),
])

print(f"Serving Flask app '{file_name}'! 🚀")

if __name__ == '__main__':
    app.run(debug=
                   True)
                #    False)
                
# =================================== Updated Database ================================= #

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# updated_path = f'data/Admin_{current_quarter}_{report_year}.xlsx'
# data_path = os.path.join(script_dir, updated_path)

# with pd.ExcelWriter(data_path, engine='xlsxwriter') as writer:
#     df.to_excel(
#             writer, 
#             sheet_name=f'MarCom {current_quarter} {report_year}', 
#             startrow=1, 
#             index=False
#         )

#     # Create the workbook to access the sheet and make formatting changes:
#     workbook = writer.book
#     sheet1 = writer.sheets['MarCom April 2025']
    
#     # Define the header format
#     header_format = workbook.add_format({
#         'bold': True, 
#         'font_size': 13, 
#         'align': 'center', 
#         'valign': 'vcenter',
#         'border': 1, 
#         'font_color': 'black', 
#         'bg_color': '#B7B7B7',
#     })
    
#     # Set column A (Name) to be left-aligned, and B-E to be right-aligned
#     left_align_format = workbook.add_format({
#         'align': 'left',  # Left-align for column A
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })

#     right_align_format = workbook.add_format({
#         'align': 'right',  # Right-align for columns B-E
#         'valign': 'vcenter',  # Vertically center
#         'border': 0  # No border for individual cells
#     })
    
#     # Create border around the entire table
#     border_format = workbook.add_format({
#         'border': 1,  # Add border to all sides
#         'border_color': 'black',  # Set border color to black
#         'align': 'center',  # Center-align text
#         'valign': 'vcenter',  # Vertically center text
#         'font_size': 12,  # Set font size
#         'font_color': 'black',  # Set font color to black
#         'bg_color': '#FFFFFF'  # Set background color to white
#     })

#     # Merge and format the first row (A1:E1) for each sheet
#     sheet1.merge_range('A1:Q1', f'MarCom Report {current_quarter} {report_year}', header_format)

#     # Set column alignment and width
#     # sheet1.set_column('A:A', 20, left_align_format)  

#     print(f"MarCom Excel file saved to {data_path}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create admin-jun-25
# heroku git:remote -a admin-jun-25
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx