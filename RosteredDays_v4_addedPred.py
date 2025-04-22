import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
import io
import plotly.graph_objects as go  # Import plotly
import xlsxwriter


# Set Streamlit Theme Colors
st.set_page_config(
    page_title="Employee Workday Projection App",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'About': "Rajat.Sachan@sodexo.com"
    }
)

st.markdown(
     """
    <style>
    body {
        color: #B5D6BC;
        background-color: #ca6a6a;
    }

    .stApp {
        background-color: #8c3232;
    }

    .stTextInput>div>div>input {
        background-color: #c74956;
        color: #4f0b76;
    }
    .stSelectbox>div>div>div>div>input {
        background-color: #c74956;
        color: #4f0b76;
    }
    .stDateInput>div>div>input{
        background-color: #4C58BD;
        color: #4f0b76;
    }
    .stButton>button {
        background-color: #642D2D;
        color: #B5D6BC;
    }

    .css-10trblm {
        background-color: #4C58BD;
        color: #B5D6BC;
    }

    .css-16idsys p {
        color: #B5D6BC;
    }

    .css-1y4p82d {
        background-color: #4C58BD;
        color: #B5D6BC;
    }
    .stMetricValue {
        color: #B5D6BC;
    }
    .stMetricLabel {
        color: #B5D6BC;
    }
    .stExpander>div>div[data-baseweb="expandable"]{
        background-color: #c74956;
    }
    .stExpander>div>div>div>div{
        background-color: #4C58BD;
    }

     /* Custom style for sidebar */
    .css-1d391kg {  /* This class targets the sidebar container 
        background-color: #c74956;  /* Purple color 
        color: #FFFFFF;  /* White text 
    }

    .css-1d391kg .stSidebar .stSidebarHeader {
        background-color: #4f0b76;  /* Purple background for the sidebar header */
    }

    .css-1d391kg .stSidebar .stSelectbox {
        background-color: #c74956;  /* Purple background for the selection box */
        color: #FFFFFF;  /* White text */
    }
    
    /* Change button color in the sidebar */
    .css-1d391kg .stButton>button {
        background-color: #4C58BD;
        color: #B5D6BC;
    }

    </style>
    """,
    unsafe_allow_html=True,
)



# Define public holidays in WA for the years 2020, 2021, 2022, 2023, 2024, 2025, 2026, and 2027
public_holidays = {
    2020: [
        {"name": "New Year's Day", "date": "2020-01-01"},
        {"name": "Australia Day", "date": "2020-01-27"},
        {"name": "Labour Day", "date": "2020-03-02"},
        {"name": "Good Friday", "date": "2020-04-10"},
        {"name": "Easter Monday", "date": "2020-04-13"},
        {"name": "Anzac Day", "date": "2020-04-25"},
        {"name": "Anzac Day (Additional Day)", "date": "2020-04-27"},
        {"name": "Western Australia Day", "date": "2020-06-01"},
        {"name": "Queen's Birthday", "date": "2020-09-28"},
        {"name": "Christmas Day", "date": "2020-12-25"},
        {"name": "Boxing Day", "date": "2020-12-26"},
        {"name": "Boxing Day (Additional Day)", "date": "2020-12-28"}
    ],
    2021: [
        {"name": "New Year's Day", "date": "2021-01-01"},
        {"name": "Australia Day", "date": "2021-01-26"},
        {"name": "Labour Day", "date": "2021-03-01"},
        {"name": "Good Friday", "date": "2021-04-02"},
        {"name": "Easter Monday", "date": "2021-04-05"},
        {"name": "Anzac Day", "date": "2021-04-25"},
        {"name": "Anzac Day (Additional Day)", "date": "2021-04-26"},
        {"name": "Western Australia Day", "date": "2021-06-07"},
        {"name": "Queen's Birthday", "date": "2021-09-27"},
        {"name": "Christmas Day", "date": "2021-12-25"},
        {"name": "Christmas Day (Additional Day)", "date": "2021-12-27"},
        {"name": "Boxing Day", "date": "2021-12-26"},
        {"name": "Boxing Day (Additional Day)", "date": "2021-12-28"}
    ],
    2022: [
        {"name": "New Year's Day", "date": "2022-01-01"},
        {"name": "New Year's Day (Additional Day)", "date": "2022-01-03"},
        {"name": "Australia Day", "date": "2022-01-26"},
        {"name": "Labour Day", "date": "2022-03-07"},
        {"name": "Good Friday", "date": "2022-04-15"},
        {"name": "Easter Sunday", "date": "2022-04-17"},
        {"name": "Easter Monday", "date": "2022-04-18"},
        {"name": "Anzac Day", "date": "2022-04-25"},
        {"name": "Western Australia Day", "date": "2022-06-06"},
        {"name": "National Day of Mourning", "date": "2022-09-22"},
        {"name": "King's Birthday", "date": "2022-09-26"},
        {"name": "Christmas Day", "date": "2022-12-25"},
        {"name": "Christmas Day (Additional Day)", "date": "2022-12-26"},
        {"name": "Boxing Day", "date": "2022-12-26"},
        {"name": "Boxing Day (Additional Day)", "date": "2022-12-27"}
    ],
    2023: [
    {"name": "New Year's Day","date": "2023-01-01"},
    {"name": "New Year's Day (Additional Day)", "date": "2023-01-02"},
    {"name": "Australia Day", "date": "2023-01-26"},
    {"name": "Labour Day", "date": "2023-03-06"},
    {"name": "Good Friday", "date": "2023-04-07"},
    {"name": "Easter Monday", "date": "2023-04-10"},
    {"name": "Anzac Day", "date": "2023-04-25"},
    {"name": "Western Australia Day", "date": "2023-06-05"},
    {"name": "King's Birthday", "date": "2023-09-25"},
    {"name": "Christmas Day","date": "2023-12-25"},
    {"name": "Boxing Day", "date": "2023-12-26"}
  ],
  2024: [
    {"name": "New Year's Day", "date": "2024-01-01"},
    {"name": "Australia Day", "date": "2024-01-26"},
    {"name": "Labour Day", "date": "2024-03-04"},
    {"name": "Good Friday", "date": "2024-03-29"},
    {"name": "Easter Monday", "date": "2024-04-01"},
    {"name": "Anzac Day", "date": "2024-04-25"},
    {"name": "Western Australia Day", "date": "2024-06-03"},
    {"name": "King's Birthday", "date": "2024-09-23"},
    {"name": "Christmas Day", "date": "2024-12-25"},
    {"name": "Boxing Day", "date": "2024-12-26"}
  ],
    2025: [
        {"name": "New Year's Day", "date": "2025-01-01"},
        {"name": "Australia Day", "date": "2025-01-27"},
        {"name": "Labour Day", "date": "2025-03-03"},
        {"name": "Good Friday", "date": "2025-04-18"},
        {"name": "Easter Sunday", "date": "2025-04-20"},
        {"name": "Easter Monday", "date": "2025-04-21"},
        {"name": "Anzac Day", "date": "2025-04-25"},
        {"name": "Western Australia Day", "date": "2025-06-02"},
        {"name": "King's Birthday", "date": "2025-09-29"},
        {"name": "Christmas Day", "date": "2025-12-25"},
        {"name": "Boxing Day", "date": "2025-12-26"}
    ],
    2026: [
        {"name": "New Year's Day", "date": "2026-01-01"},
        {"name": "Australia Day", "date": "2026-01-26"},
        {"name": "Labour Day", "date": "2026-03-02"},
        {"name": "Good Friday", "date": "2026-04-03"},
        {"name": "Easter Sunday", "date": "2026-04-05"},
        {"name": "Easter Monday", "date": "2026-04-06"},
        {"name": "Anzac Day", "date": "2026-04-25"},
        {"name": "Anzac Day", "date": "2026-04-27"},
        {"name": "Western Australia Day", "date": "2026-06-01"},
        {"name": "King's Birthday", "date": "2026-09-28"},
        {"name": "Christmas Day", "date": "2026-12-25"},
        {"name": "Boxing Day", "date": "2026-12-26"},
        {"name": "Boxing Day", "date": "2026-12-28"}
    ],
    2027: [
        {"name": "New Year's Day", "date": "2027-01-01"},
        {"name": "Australia Day", "date": "2027-01-26"},
        {"name": "Labour Day", "date": "2027-03-01"},
        {"name": "Good Friday", "date": "2027-03-26"},
        {"name": "Easter Sunday", "date": "2027-03-28"},
        {"name": "Easter Monday", "date": "2027-03-29"},
        {"name": "Anzac Day", "date": "2027-04-25"},
        {"name": "Anzac Day", "date": "2027-04-26"},
        {"name": "Western Australia Day", "date": "2027-06-07"},
        {"name": "King's Birthday", "date": "2027-09-27"},
        {"name": "Christmas Day", "date": "2027-12-25"},
        {"name": "Christmas Day", "date": "2027-12-27"},
        {"name": "Boxing Day", "date": "2027-12-26"},
        {"name": "Boxing Day", "date": "2027-12-28"}
    ]
}


# Define public holidays in QLD for the years 2025, 2026, and 2027
public_holidays_Queensland = {
     2020: [
        {"name": "New Year's Day", "date": "2020-01-01"},
        {"name": "Australia Day", "date": "2020-01-27"},
        {"name": "Good Friday", "date": "2020-04-10"},
        {"name": "The day after Good Friday", "date": "2020-04-11"},
        {"name": "Easter Sunday", "date": "2020-04-12"},
        {"name": "Easter Monday", "date": "2020-04-13"},
        {"name": "Anzac Day", "date": "2020-04-25"},
        {"name": "Labour Day", "date": "2020-05-04"},
        {"name": "Queen's Birthday", "date": "2020-10-05"},
        {"name": "Christmas Eve (from 6pm)", "date": "2020-12-24"},
        {"name": "Christmas Day", "date": "2020-12-25"},
        {"name": "Boxing Day", "date": "2020-12-26"},
        {"name": "Boxing Day (Additional Day)", "date": "2020-12-28"}
    ],
    2021: [
        {"name": "New Year's Day", "date": "2021-01-01"},
        {"name": "Australia Day", "date": "2021-01-26"},
        {"name": "Good Friday", "date": "2021-04-02"},
        {"name": "The day after Good Friday", "date": "2021-04-03"},
        {"name": "Easter Sunday", "date": "2021-04-04"},
        {"name": "Easter Monday", "date": "2021-04-05"},
        {"name": "Anzac Day", "date": "2021-04-26"},
        {"name": "Labour Day", "date": "2021-05-03"},
        {"name": "Queen's Birthday", "date": "2021-10-04"},
        {"name": "Christmas Eve (from 6pm)", "date": "2021-12-24"},
        {"name": "Christmas Day", "date": "2021-12-25"},
        {"name": "Christmas Day (Additional Day)", "date": "2021-12-27"},
        {"name": "Boxing Day", "date": "2021-12-26"},
        {"name": "Boxing Day (Additional Day)", "date": "2021-12-28"}
    ],
    2022: [
        {"name": "New Year's Day", "date": "2022-01-01"},
        {"name": "New Year's Day (Additional Day)", "date": "2022-01-03"},
        {"name": "Australia Day", "date": "2022-01-26"},
        {"name": "Good Friday", "date": "2022-04-15"},
        {"name": "The day after Good Friday", "date": "2022-04-16"},
        {"name": "Easter Sunday", "date": "2022-04-17"},
        {"name": "Easter Monday", "date": "2022-04-18"},
        {"name": "Anzac Day", "date": "2022-04-25"},
        {"name": "Labour Day", "date": "2022-05-02"},
        {"name": "National Day of Mourning for Her Majesty The Queen", "date": "2022-09-22"},
        {"name": "Queen's Birthday", "date": "2022-10-03"},
        {"name": "Christmas Eve (from 6pm)", "date": "2022-12-24"},
        {"name": "Christmas Day", "date": "2022-12-25"},
        {"name": "Christmas Day (Additional Day)", "date": "2022-12-27"},
        {"name": "Boxing Day", "date": "2022-12-26"}
    ],
    
    2023: [
      {
        "name": "New Year's Day",
        "date": "2023-01-01"
      },
      {
        "name": "New Year's Day (Additional Day)",
        "date": "2023-01-02"
      },
      {
        "name": "Australia Day",
        "date": "2023-01-26"
      },
      {
        "name": "Good Friday",
        "date": "2023-04-07"
      },
      {
        "name": "Easter Saturday",
        "date": "2023-04-08"
      },
      {
        "name": "Easter Sunday",
        "date": "2023-04-09"
      },
      {
        "name": "Easter Monday",
        "date": "2023-04-10"
      },
      {
        "name": "Anzac Day",
        "date": "2023-04-25"
      },
      {
        "name": "Labour Day",
        "date": "2023-05-01"
      },
      {
        "name": "Queen's Birthday",
        "date": "2023-10-02"
      },
      {
        "name": "Christmas Day",
        "date": "2023-12-25"
      },
      {
        "name": "Boxing Day",
        "date": "2023-12-26"
      }
    ],
    2024: [
      {
        "name": "New Year's Day",
        "date": "2024-01-01"
      },
      {
        "name": "Australia Day",
        "date": "2024-01-26"
      },
      {
        "name": "Good Friday",
        "date": "2024-03-29"
      },
      {
        "name": "Easter Saturday",
        "date": "2024-03-30"
      },
      {
        "name": "Easter Sunday",
        "date": "2024-03-31"
      },
      {
        "name": "Easter Monday",
        "date": "2024-04-01"
      },
      {
        "name": "Anzac Day",
        "date": "2024-04-25"
      },
      {
        "name": "Labour Day",
        "date": "2024-05-06"
      },
      {
        "name": "King's Birthday",
        "date": "2024-10-07"
      },
      {
        "name": "Christmas Day",
        "date": "2024-12-25"
      },
      {
        "name": "Boxing Day",
        "date": "2024-12-26"
      }
    ],
    2025: [
        {"name": "New Year's Day", "date": "2025-01-01"},
        {"name": "Australia Day", "date": "2025-01-26"},
        {"name": "Labour Day", "date": "2025-05-04"},
        {"name": "Anzac Day", "date": "2025-04-25"},
        {"name": "Good Friday", "date": "2025-04-18"},
        {"name": "Easter Saturday", "date": "2025-04-19"},
        {"name": "Easter Sunday", "date": "2025-04-20"},
        {"name": "Easter Monday", "date": "2025-04-21"},
        {"name": "Queensland Day", "date": "2025-10-05"},
        {"name": "Labour Day (South East Queensland)", "date": "2025-10-05"},
        {"name": "Christmas Day", "date": "2025-12-25"},
        {"name": "Boxing Day", "date": "2025-12-26"}
    ],
    2026: [
        {"name": "New Year's Day", "date": "2026-01-01"},
        {"name": "Australia Day", "date": "2026-01-26"},
        {"name": "Labour Day", "date": "2026-05-03"},
        {"name": "Anzac Day", "date": "2026-04-25"},
        {"name": "Good Friday", "date": "2026-04-03"},
        {"name": "Easter Saturday", "date": "2026-04-04"},
        {"name": "Easter Sunday", "date": "2026-04-05"},
        {"name": "Easter Monday", "date": "2026-04-06"},
        {"name": "Queensland Day", "date": "2026-10-04"},
        {"name": "Labour Day (South East Queensland)", "date": "2026-10-04"},
        {"name": "Christmas Day", "date": "2026-12-25"},
        {"name": "Boxing Day", "date": "2026-12-26"},
        {"name": "Boxing Day (Additional Day)", "date": "2026-12-28"}
    ],
    2027: [
        {"name": "New Year's Day", "date": "2027-01-01"},
        {"name": "Australia Day", "date": "2027-01-26"},
        {"name": "Labour Day", "date": "2027-05-02"},
        {"name": "Anzac Day", "date": "2027-04-25"},
        {"name": "Good Friday", "date": "2027-04-02"},
        {"name": "Easter Saturday", "date": "2027-04-03"},
        {"name": "Easter Sunday", "date": "2027-04-04"},
        {"name": "Easter Monday", "date": "2027-04-05"},
        {"name": "Queensland Day", "date": "2027-10-03"},
        {"name": "Labour Day (South East Queensland)", "date": "2027-10-03"},
        {"name": "Christmas Day", "date": "2027-12-25"},
        {"name": "Christmas Day (Additional Day)", "date": "2027-12-27"},
        {"name": "Boxing Day", "date": "2027-12-26"},
        {"name": "Boxing Day (Additional Day)", "date": "2027-12-28"}
    ]
}

# Function to find the best start date
def find_best_start_date(start_date_str, roster_pattern, selected_public_holidays, location, weeks):
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = start_date + timedelta(weeks=weeks)

    best_start_date = None
    min_holidays_worked = float('inf')
    best_holiday_names = []  # Store holiday names for the best date

    holiday_counts = {}

    current_date = start_date
    while current_date <= end_date:
        current_date_str = current_date.strftime("%Y-%m-%d")
        workdays_in_period, holidays_on_workdays, _, _, _, _ = calculate_workdays(
            current_date_str, roster_pattern, selected_public_holidays, location, year_span=1, manual_check=True
        )

        holiday_counts[current_date] = {
            "count": workdays_in_period,
            "names": ", ".join(holidays_on_workdays)
        }

        if workdays_in_period < min_holidays_worked:
            min_holidays_worked = workdays_in_period
            best_start_date = current_date
            best_holiday_names = holidays_on_workdays  # Capture holiday names

        current_date += timedelta(days=1)

    if min_holidays_worked > 7:
        return None, None, None, None  # Return None if the best option is still over the risk threshold.

    return best_start_date, min_holidays_worked, holiday_counts, best_holiday_names


# Function to calculate the end date based on a 12-month span
def calculate_end_date(start_date):
    return start_date + timedelta(days=365)

def add_custom_holiday(location, holiday_name, holiday_date):
    if location == "Western Australia":
        st.session_state.custom_holidays_wa.append({"name": holiday_name, "date": holiday_date.strftime('%Y-%m-%d')})
    elif location == "Queensland":
        st.session_state.custom_holidays_qld.append({"name": holiday_name, "date": holiday_date.strftime('%Y-%m-%d')})

# Initialize session state for custom holidays
if "custom_holidays_wa" not in st.session_state:
    st.session_state.custom_holidays_wa = []
if "custom_holidays_qld" not in st.session_state:
    st.session_state.custom_holidays_qld = []



# Sidebar for Location Selection
location = st.sidebar.radio("Select Location for Public Holidays", ["Queensland", "Western Australia"])

# Custom Holiday Input
st.sidebar.subheader("Enter Custom Public Holiday")
holiday_name = st.sidebar.text_input("Holiday Name")
holiday_date = st.sidebar.date_input("Holiday Date", min_value=datetime(2020, 1, 1))

# Button to add custom holiday
if st.sidebar.button("Add Custom Holiday"):
    if holiday_name and holiday_date:
        add_custom_holiday(location, holiday_name, holiday_date)
        st.success(f"Holiday '{holiday_name}' added for {location} on {holiday_date.strftime('%d %B %Y')}")


# Function to remove a custom holiday
def remove_custom_holiday(location, holiday_name, holiday_date):
    if location == "Western Australia":
        year = holiday_date.year
        # Remove the holiday from the list in session state
        st.session_state.custom_holidays_wa = [
            holiday for holiday in st.session_state.custom_holidays_wa
            if holiday["name"] != holiday_name or holiday["date"] != holiday_date.strftime('%Y-%m-%d')
        ]
    elif location == "Queensland":
        year = holiday_date.year
        # Remove the holiday from the list in session state
        st.session_state.custom_holidays_qld = [
            holiday for holiday in st.session_state.custom_holidays_qld
            if holiday["name"] != holiday_name or holiday["date"] != holiday_date.strftime('%Y-%m-%d')
        ]


# Display custom holidays in the sidebar with remove functionality
st.sidebar.subheader("Custom Holidays Added")
if location == "Western Australia" and st.session_state.custom_holidays_wa:
    for holiday in st.session_state.custom_holidays_wa:
        st.sidebar.write(f"{holiday['name']} on {holiday['date']}")
        # Add a remove button for each custom holiday
        if st.sidebar.button(f"Remove {holiday['name']} on {holiday['date']}"):
            remove_custom_holiday(location, holiday['name'], datetime.strptime(holiday['date'], "%Y-%m-%d"))
            st.success(f"Holiday '{holiday['name']}' removed for {location}.")
elif location == "Queensland" and st.session_state.custom_holidays_qld:
    for holiday in st.session_state.custom_holidays_qld:
        st.sidebar.write(f"{holiday['name']} on {holiday['date']}")
        # Add a remove button for each custom holiday
        if st.sidebar.button(f"Remove {holiday['name']} on {holiday['date']}"):
            remove_custom_holiday(location, holiday['name'], datetime.strptime(holiday['date'], "%Y-%m-%d"))
            st.success(f"Holiday '{holiday['name']}' removed for {location}.")
else:
    st.sidebar.write("No custom holidays added yet.")
# Display custom holidays
st.sidebar.subheader("Custom Holidays Added")
if location == "Western Australia" and st.session_state.custom_holidays_wa:
    for holiday in st.session_state.custom_holidays_wa:
        st.sidebar.write(f"{holiday['name']} on {holiday['date']}")
elif location == "Queensland" and st.session_state.custom_holidays_qld:
    for holiday in st.session_state.custom_holidays_qld:
        st.sidebar.write(f"{holiday['name']} on {holiday['date']}")
else:
    st.sidebar.write("No custom holidays added yet.")


# Print custom holidays for the selected location
if location == "Western Australia":
    st.write("Custom Holidays for Western Australia:")
    if st.session_state.custom_holidays_wa:
        for holiday in st.session_state.custom_holidays_wa:
            st.write(f"- {holiday['name']} on {holiday['date']}")
    else:
        st.write("No custom holidays added for Western Australia.")
        
elif location == "Queensland":
    st.write("Custom Holidays for Queensland:")
    if st.session_state.custom_holidays_qld:
        for holiday in st.session_state.custom_holidays_qld:
            st.write(f"- {holiday['name']} on {holiday['date']}")
    else:
        st.write("No custom holidays added for Queensland.")


# Initialize employee_name
employee_name = None

# Sidebar for Navigation
page = st.sidebar.selectbox("Select Page", ["Projection", "All Rosters", "Start Date Optimizer"])

def calculate_workdays(start_date, roster_pattern, public_holidays, location, year_span=1, manual_check=False):
    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = start_date + timedelta(days=365 * year_span)
    work_days = []
    off_days = []
    current_day = start_date
    pattern_length = sum(roster_pattern)
    cycles = []  # Store cycle details

    all_public_holidays = []
    for year in range(2020, 2028):
        if location == "Queensland":
            if year in public_holidays:  # Add this check
                all_public_holidays.extend(public_holidays[year])
        elif location == "Western Australia":
            if year in public_holidays:  # Add this check
                all_public_holidays.extend(public_holidays[year])
    # Add custom holidays from session state
    if location == "Queensland":
        all_public_holidays.extend(st.session_state.custom_holidays_qld)
    elif location == "Western Australia":
        all_public_holidays.extend(st.session_state.custom_holidays_wa)

    while current_day <= end_date:
        cycle_start = current_day
        work_end = min(current_day + timedelta(days=roster_pattern[0]), end_date + timedelta(days=1))
        work_period = []
        while current_day < work_end and current_day <= end_date:
            work_days.append(current_day)
            work_period.append(current_day)
            current_day += timedelta(days=1)

        off_end = min(current_day + timedelta(days=roster_pattern[1]), end_date + timedelta(days=1))
        off_period = []
        while current_day < off_end and current_day <= end_date:
            off_days.append(current_day)
            off_period.append(current_day)
            current_day += timedelta(days=1)

        if manual_check:
            cycles.append({
                "cycle_start": cycle_start,
                "work_period": work_period,
                "off_period": off_period,
                "work_days_count": len(work_period),  # Count work days
                "off_days_count": len(off_period)  # Count off days
            })

        if current_day > end_date:
            break

    holidays_on_workdays = []
    holidays_on_rr = []
    workdays_in_period = 0
    for holiday in all_public_holidays:
        holiday_date = datetime.strptime(holiday["date"], "%Y-%m-%d")
        if holiday_date in off_days:
            holidays_on_rr.append(f"{holiday['name']} on {holiday_date.strftime('%d %B %Y')}")
        if holiday_date in work_days:
            holidays_on_workdays.append(f"{holiday['name']} on {holiday_date.strftime('%d %B %Y')}")
            workdays_in_period += 1
            # Add print statement with holiday details
            print(f"Public Holiday on Workday: {holiday['name']} on {holiday_date.strftime('%A, %d %B %Y')}")

    total_workdays = len(work_days)
    total_off_days = len(off_days)

    if manual_check:
        return workdays_in_period, holidays_on_workdays, holidays_on_rr, total_workdays, total_off_days, cycles
    else:
        return workdays_in_period, holidays_on_workdays, holidays_on_rr, total_workdays, total_off_days


if page == "Projection":
    st.title("Employee Workday Projection App")
    
    # Input fields
    employee_name = st.text_input("Enter Employee Name:")
    
    # User selects Start Date
    start_date = st.date_input("Select Start Date", min_value=datetime(2020, 1, 1), key="start_date_input")
    
    # Calculate end date (12 months after the start date)
    end_date = calculate_end_date(start_date)
    
    # Display End Date (read-only field)
    st.text_input("End Date (Automatically Calculated)", value=end_date.strftime("%Y-%m-%d"), disabled=True)
    
    # Roster pattern selection
    roster_pattern = st.selectbox("Select Roster Pattern", options=["14/7", "8/6", "5/2","14/14"])

    # Convert the roster pattern into the number of workdays and R&R days
    if roster_pattern == "14/7":
        roster_pattern = [14, 7]
    elif roster_pattern == "8/6":
        roster_pattern = [8, 6]
    elif roster_pattern == "5/2":
        roster_pattern = [5, 2]
    elif roster_pattern == "14/14":
        roster_pattern = [14,14]

    # Handle user selection for holiday location
    if location == "Queensland":
        selected_public_holidays = public_holidays_Queensland
    else:
        selected_public_holidays = public_holidays
    
    

    # Run the calculation when valid inputs are provided
    if employee_name and start_date and roster_pattern:
        start_date_str = start_date.strftime("%Y-%m-%d")

        workdays_in_period, holidays_on_workdays, holidays_on_rr, total_work_days, total_off_days, cycles = calculate_workdays(
            start_date_str, roster_pattern, selected_public_holidays, location, year_span=1, manual_check=True)
        
        st.session_state.employee_name = employee_name
        st.session_state.start_date = start_date
        st.session_state.roster_pattern = roster_pattern
        st.session_state.workdays_in_period = workdays_in_period
        st.session_state.holidays_on_workdays = holidays_on_workdays
        st.session_state.holidays_on_rr = holidays_on_rr
        st.session_state.total_work_days = total_work_days
        st.session_state.total_off_days = total_off_days
        st.session_state.cycles = cycles
        
         # Access the variables from session state
        all_public_holidays = []
        for year in range(2020, 2028):
            if location == "Queensland":
                all_public_holidays.extend(public_holidays_Queensland[year])
            elif location == "Western Australia":
                all_public_holidays.extend(public_holidays[year])

        if location == "Queensland":
            all_public_holidays.extend(st.session_state.custom_holidays_qld)
        elif location == "Western Australia":
            all_public_holidays.extend(st.session_state.custom_holidays_wa)

        work_days = []
        off_days = []
        for cycle in st.session_state.cycles:
            work_days.extend(cycle['work_period'])
            off_days.extend(cycle['off_period'])

        st.subheader(f"Workday Projection for {employee_name}")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Workdays", total_work_days)
        col2.metric("Total Off Days", total_off_days)
        col3.metric("Public Holidays (Work)", workdays_in_period)
        col4.metric("Public Holidays on R&R", len(holidays_on_rr))

        st.markdown("---")

        st.subheader("Public Holidays on Workdays:")
        if holidays_on_workdays:
            for holiday in holidays_on_workdays:
                st.write(f"- {holiday}")
        else:
            st.write("No public holidays fell on workdays.")

        st.markdown("---")

        st.subheader("Public Holidays on R&R Days:")
        if holidays_on_rr:
            for holiday in holidays_on_rr:
                st.write(f"- {holiday}")
        else:
            st.write("No public holidays fell on R&R days.")

        # Manual Check Button
        if st.button("Detailed Cycle Breakdown"):
            st.subheader("Cycle Breakdown:")
            total_work_cycle_days = 0
            total_off_cycle_days = 0

            for i, cycle in enumerate(cycles):
                with st.expander(f"Cycle {i+1}: Start: {cycle['cycle_start'].strftime('%d %B %Y')}"):
                    st.write(f"**Start:** {cycle['cycle_start'].strftime('%d %B %Y')}")

                    st.write("**Work Days:**")
                    for work_day in cycle['work_period']:
                        st.write(f"- {work_day.strftime('%d %B %Y')}")
                    st.write(f"(Total Days: {cycle['work_days_count']})")

                    st.write("**Off Days:**")
                    for off_day in cycle['off_period']:
                        st.write(f"- {off_day.strftime('%d %B %Y')}")
                    st.write(f"(Total Days: {cycle['off_days_count']})")

                    cycle_holidays = []
                    for holiday_year in selected_public_holidays:
                        for holiday in selected_public_holidays[holiday_year]:
                            holiday_date = datetime.strptime(holiday['date'], '%Y-%m-%d').date()
                            if holiday_date in [d.date() for d in cycle['work_period']]:
                                cycle_holidays.append((holiday['name'], holiday_date, "work"))
                            elif holiday_date in [d.date() for d in cycle['off_period']]:
                                cycle_holidays.append((holiday['name'], holiday_date, "off"))

                    if cycle_holidays:
                        st.write("**Public Holidays:**")
                        for holiday_name, holiday_date, status in cycle_holidays:
                            st.write(f"- {holiday_name} ({holiday_date.strftime('%d %B %Y')}) - {status}")

                total_work_cycle_days += cycle['work_days_count']
                total_off_cycle_days += cycle['off_days_count']

            st.write(f"**Total Work Cycle Days:** {total_work_cycle_days}")
            st.write(f"**Total Off Cycle Days:** {total_off_cycle_days}")
        

         # Export Functionality
        if st.button("Export to Excel"):
            # Worksheet 1: Key Details
            data_worksheet1 = {
                "Key Stats": ["Total Workdays", "Total Off Days", "Public Holidays (Work)", "Public Holidays on R&R"],
                "Value": [st.session_state.total_work_days, st.session_state.total_off_days, st.session_state.workdays_in_period, len(st.session_state.holidays_on_rr)]
            }
            df_worksheet1_stats = pd.DataFrame(data_worksheet1)

            data_worksheet1_table = []
            for holiday in all_public_holidays:
                holiday_date = datetime.strptime(holiday["date"], "%Y-%m-%d").date()
                if holiday_date in [d.date() for d in work_days]:
                    status = "Worked"
                elif holiday_date in [d.date() for d in off_days]:
                    status = "Not Worked"
                else:
                    status = "Outside 12 Month Period" #Handles holidays outside of the projected dates.
                data_worksheet1_table.append({
                    "Start Date": st.session_state.start_date.strftime("%Y-%m-%d"),
                    "Public Holidays Date": holiday_date.strftime("%Y-%m-%d"),
                    "Only Public Holidays": holiday["name"],
                    "Roster": f"{st.session_state.roster_pattern[0]}/{st.session_state.roster_pattern[1]}",
                    "Status Worked/Not Worked": status
                })
            df_worksheet1_table = pd.DataFrame(data_worksheet1_table)

            # Worksheet 2: Detailed Data
            data_worksheet2 = []
            current_date = st.session_state.start_date
            cycle_num = 1
            cycle_start = st.session_state.start_date
            work_end = min(cycle_start + timedelta(days=st.session_state.roster_pattern[0]), end_date + timedelta(days=1))
            off_end = min(work_end + timedelta(days=st.session_state.roster_pattern[1]), end_date + timedelta(days=1))

            while current_date <= end_date:
                is_holiday = "N"
                for holiday in all_public_holidays:
                    if current_date == datetime.strptime(holiday["date"], "%Y-%m-%d").date():
                        is_holiday = "Y"
                        break

                # Corrected status calculation
                current_datetime = datetime(current_date.year, current_date.month, current_date.day) #Convert date to datetime
                if current_datetime in work_days:
                    status = "Worked"
                elif current_datetime in off_days:
                    status = "Not Worked"
                else:
                    status = "Outside 12 Month Period" #Handle dates outside of project dates.

                data_worksheet2.append({
                    "Start Date": st.session_state.start_date.strftime("%Y-%m-%d"),
                    "All Dates": current_date.strftime("%Y-%m-%d"),
                    "Cycle Number": cycle_num,
                    "Public Holidays (Y/N)": is_holiday,
                    "Roster": f"{st.session_state.roster_pattern[0]}/{st.session_state.roster_pattern[1]}",
                    "Status": status
                })

                current_date += timedelta(days=1)
                if current_date == work_end:
                    cycle_start = work_end
                    work_end = min(cycle_start + timedelta(days=st.session_state.roster_pattern[1]), end_date + timedelta(days=1))
                if current_date == off_end:
                    cycle_start = off_end
                    work_end = min(cycle_start + timedelta(days=st.session_state.roster_pattern[0]), end_date + timedelta(days=1))
                    cycle_num += 1

            df_worksheet2 = pd.DataFrame(data_worksheet2)

            # Create Excel file in memory
            excel_file = io.BytesIO()
            with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:
                df_worksheet1_stats.to_excel(writer, sheet_name="Key Stats", index=False)
                df_worksheet1_table.to_excel(writer, sheet_name="Public Holidays", index=False)
                df_worksheet2.to_excel(writer, sheet_name="Detailed Data", index=False)

            excel_file.seek(0)

            # Download Excel file
            st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name=f"{st.session_state.employee_name}_workday_projection.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

elif page == "All Rosters":
    st.title("All Rosters Projection")

    start_date = st.date_input("Select Start Date", min_value=datetime(2020, 1, 1), key="all_rosters_start_date")

    if st.button("Export All Rosters to Excel"):
        excel_file = io.BytesIO()
        with pd.ExcelWriter(excel_file, engine="xlsxwriter") as writer:
            roster_patterns = [[14, 7], [8, 6], [5, 2]]
            roster_names = ["14/7", "8/6", "5/2"]

            for i, roster_pattern in enumerate(roster_patterns):
                roster_name = roster_names[i]
                start_date_str = start_date.strftime("%Y-%m-%d")

                selected_public_holidays = []
                for year in range(2020, 2028):
                    if location == "Queensland":
                        selected_public_holidays.extend(public_holidays_Queensland[year])
                    elif location == "Western Australia":
                        selected_public_holidays.extend(public_holidays[year])

                if location == "Queensland":
                    selected_public_holidays.extend(st.session_state.custom_holidays_qld)
                elif location == "Western Australia":
                    selected_public_holidays.extend(st.session_state.custom_holidays_wa)

                
                

                # Calculate end_date (e.g., one year from start_date)
                end_date = start_date + timedelta(days=365)

                workdays_in_period, holidays_on_workdays, holidays_on_rr, total_work_days, total_off_days, cycles = calculate_workdays(
                    start_date_str, roster_pattern, selected_public_holidays, location, year_span=1, manual_check=True
                )

                all_public_holidays = selected_public_holidays

                work_days = []
                off_days = []
                for cycle in cycles:
                    work_days.extend(cycle['work_period'])
                    off_days.extend(cycle['off_period'])

                # Worksheet 1: Key Details
                data_worksheet1 = {
                    "Key Stats": ["Total Workdays", "Total Off Days", "Public Holidays (Work)", "Public Holidays on R&R"],
                    "Value": [total_work_days, total_off_days, workdays_in_period, len(holidays_on_rr)]
                }
                df_worksheet1_stats = pd.DataFrame(data_worksheet1)

                data_worksheet1_table = []
                for holiday in all_public_holidays:
                    holiday_date = datetime.strptime(holiday["date"], "%Y-%m-%d").date()
                    if holiday_date in [d.date() for d in work_days]:
                        status = "Worked"
                    elif holiday_date in [d.date() for d in off_days]:
                        status = "Not Worked"
                    else:
                        status = "Outside Time Period"
                    data_worksheet1_table.append({
                        "Start Date": start_date.strftime("%Y-%m-%d"),
                        "Public Holidays Date": holiday_date.strftime("%Y-%m-%d"),
                        "Only Public Holidays": holiday["name"],
                        "Roster": roster_name,
                        "Status Worked/Not Worked": status
                    })
                df_worksheet1_table = pd.DataFrame(data_worksheet1_table)

                # Worksheet 2: Detailed Data
                data_worksheet2 = []
                current_date = start_date
                cycle_num = 1
                cycle_start = start_date
                work_end = min(cycle_start + timedelta(days=roster_pattern[0]), end_date + timedelta(days=1))
                off_end = min(work_end + timedelta(days=roster_pattern[1]), end_date + timedelta(days=1))

                while current_date <= end_date:
                    is_holiday = "N"
                    for holiday in all_public_holidays:
                        if current_date == datetime.strptime(holiday["date"], "%Y-%m-%d").date():
                            is_holiday = "Y"
                            break

                    current_datetime = datetime(current_date.year, current_date.month, current_date.day)
                    if current_datetime in work_days:
                        status = "Worked"
                    elif current_datetime in off_days:
                        status = "Not Worked"
                    else:
                        status = "Outside Time Period"

                    data_worksheet2.append({
                        "Start Date": start_date.strftime("%Y-%m-%d"),
                        "All Dates": current_date.strftime("%Y-%m-%d"),
                        "Cycle Number": cycle_num,
                        "Public Holidays (Y/N)": is_holiday,
                        "Roster": roster_name,
                        "Status": status
                    })

                    current_date += timedelta(days=1)
                    if current_date == work_end:
                        cycle_start = work_end
                        work_end = min(cycle_start + timedelta(days=roster_pattern[1]), end_date + timedelta(days=1))
                    if current_date == off_end:
                        cycle_start = off_end
                        work_end = min(cycle_start + timedelta(days=roster_pattern[0]), end_date + timedelta(days=1))
                        cycle_num += 1

                df_worksheet2 = pd.DataFrame(data_worksheet2)

                
                df_worksheet1_table.to_excel(writer, sheet_name=f"{roster_name.replace('/', '-') } Public Holidays", index=False)
                df_worksheet2.to_excel(writer, sheet_name=f"{roster_name.replace('/', '-') } Detailed Data", index=False)

        excel_file.seek(0)

        st.download_button(
            label="Download All Rosters Excel",
            data=excel_file,
            file_name=f"all_rosters_projection.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
elif page == "Start Date Optimizer":
    st.title("Start Date Optimizer")

    start_date = st.date_input("Select Start Date for Optimization", min_value=datetime(2020, 1, 1), key="optimizer_start_date")
    window_weeks = st.selectbox("Select Optimization Window (weeks)", options=[2, 3, 4, 6, 8])

    if location == "Queensland":
        selected_public_holidays = public_holidays_Queensland
    else:
        selected_public_holidays = public_holidays

    roster_patterns = {
        "14-7": [14, 7],
        "8-6": [8, 6],
        "5-2": [5, 2],
        "14-14": [14, 14]
    }

    if st.button("Find Best Start Date for All Rosters"):
        st.session_state.export_data = []

        for roster_name, roster_pattern in roster_patterns.items():
            start_date_str = start_date.strftime("%Y-%m-%d")
            best_start_date, min_holidays_worked, holiday_counts, best_holiday_names = find_best_start_date(
                start_date_str, roster_pattern, selected_public_holidays, location, window_weeks
            )

            if best_start_date:
                holiday_names_str = ", ".join(best_holiday_names)
                st.markdown(
                    f"Best Start Date (within {window_weeks} weeks): {best_start_date.strftime('%Y-%m-%d')}<br>"
                    f"Public Holidays (Work): {min_holidays_worked}<br>"
                    f"Roster: {roster_name}<br>"
                    f"Holidays: {holiday_names_str}",
                    unsafe_allow_html=True
                )

                # Plotly Chart Logic
                if holiday_counts:
                    data = []
                    for date, info in holiday_counts.items():
                        data.append({
                            "Date": date,
                            "Public Holidays (Work)": info["count"],
                            "Holiday Names": info["names"]
                        })

                    df_holiday_counts = pd.DataFrame(data)
                    df_holiday_counts['Date'] = df_holiday_counts['Date'].dt.strftime('%Y-%m-%d')
                    df_holiday_counts = df_holiday_counts.set_index('Date')

                    fig = go.Figure(data=go.Scatter(
                        x=df_holiday_counts.index,
                        y=df_holiday_counts['Public Holidays (Work)'],
                        mode='lines+markers+text',
                        text=df_holiday_counts['Public Holidays (Work)'], #use holiday names as text instead of holiday count.
                        textposition='top center'
                    ))

                    fig.update_layout(title=f'Public Holidays (Work) Count - {roster_name}', xaxis_title='Start Date', yaxis_title='Public Holidays (Work)')

                    st.plotly_chart(fig)

                    # Store data for export
                    export_data_to_store = df_holiday_counts.copy() #create a copy of the dataframe.
                    export_data_to_store.index = export_data_to_store.index.astype(str) #convert index to string.

                    st.session_state.export_data.append({
                        "Roster": roster_name,
                        "Best Start Date": best_start_date.strftime('%Y-%m-%d'),
                        "Public Holidays (Work)": min_holidays_worked,
                        "Holidays": holiday_names_str,
                        "Chart Data": export_data_to_store #store the converted dataframe.
                    })

            else:
                st.warning(f"No suitable start date found within the {window_weeks}-week window for {roster_name}.")
        print(st.session_state.export_data)

    if "export_data" in st.session_state and st.session_state.export_data:
        if st.button("Export to Excel"):
            try:
                output = io.BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')

                # Summary sheet
                summary_data = []
                for item in st.session_state.export_data:
                    summary_data.append({
                        "Roster": item["Roster"],
                        "Best Start Date": item["Best Start Date"],
                        "Public Holidays (Work)": item["Public Holidays (Work)"],
                        "Holidays": item["Holidays"]
                    })
                pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

                # Chart data sheets
                for item in st.session_state.export_data:
                    item["Chart Data"].to_excel(writer, sheet_name=item["Roster"])

                writer.close()
                processed_data = output.getvalue()

                st.download_button(
                    label="Download Excel file",
                    data=processed_data,
                    file_name="start_date_optimizer_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"An error occurred during Excel export: {e}")
