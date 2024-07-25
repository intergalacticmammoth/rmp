import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta
import pytz
import io

# Set page configuration
st.set_page_config(page_title="Hospital Shift Calendar", page_icon="👨‍⚕️", layout="wide")


@st.cache_data
def read_excel_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    monthly_schedule = pd.read_excel(xls, "Monatsplan")
    names = pd.read_excel(xls, "Namen")
    definitions = pd.read_excel(xls, "Dienstdefinition")
    return monthly_schedule, names, definitions


def detect_shifts(monthly_schedule, employee_name):
    employee_row = monthly_schedule[monthly_schedule["Name"] == employee_name].iloc[0]
    shifts = [
        (date, shift_code)
        for date, shift_code in employee_row.iloc[2:].items()
        if pd.notna(shift_code) and shift_code != "0"
    ]
    # Exclude N and -DF shift codes
    shifts = [
        (date, shift_code)
        for date, shift_code in shifts
        if not shift_code.startswith("N") and not shift_code.endswith("DF")
    ]
    return shifts


def get_shift_times(shift_code, definitions):
    shift_def = definitions[
        definitions["PersOff"].str.contains(shift_code, na=False, regex=False)
    ]
    if not shift_def.empty and pd.notna(shift_def["Werktag"].iloc[0]):
        start_time, end_time = shift_def["Werktag"].iloc[0].split("-")
        return start_time.strip(), end_time.strip()
    return "00:00", "23:59"  # Default to full day if not found


def create_ics_file(shifts, employee_name, definitions):
    cal = Calendar()
    for date, shift_code in shifts:
        event = Event()
        event.add("summary", f"Shift: {shift_code}")
        start_time, end_time = get_shift_times(shift_code, definitions)
        start_datetime = datetime.combine(
            date.date(), datetime.strptime(start_time, "%H:%M").time()
        )
        end_datetime = datetime.combine(
            date.date(), datetime.strptime(end_time, "%H:%M").time()
        )
        if end_datetime <= start_datetime:
            end_datetime += timedelta(days=1)
        event.add("dtstart", start_datetime)
        event.add("dtend", end_datetime)
        event.add("dtstamp", datetime.now(tz=pytz.UTC))
        event.add("location", "Hospital")
        cal.add_component(event)
    return cal.to_ical()


# App header
st.title("👨‍⚕️ Hospital Shift Calendar")
st.markdown("---")

# Sidebar for file upload
with st.sidebar:
    st.header("Upload File")
    uploaded_file = st.file_uploader("Choose an XLSM file", type="xlsm")

if uploaded_file is not None:
    monthly_schedule, names, definitions = read_excel_file(uploaded_file)

    st.subheader("Who are you?")
    employee_names = names["Name"].tolist()
    selected_employee = st.selectbox(
        "Select your name", employee_names, key="employee_select"
    )

    if selected_employee:
        shifts = detect_shifts(monthly_schedule, selected_employee)

        st.markdown("---")
        st.subheader("Here's Your Schedule")
        if shifts:
            df = pd.DataFrame(shifts, columns=["Date", "Shift Code"])
            df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m-%d")
            st.dataframe(df, use_container_width=True, hide_index=True)

            st.subheader("Download Calendar Invites")
            ics_data = create_ics_file(shifts, selected_employee, definitions)
            st.download_button(
                label="Download Calendar Invites",
                data=ics_data,
                file_name=f"{selected_employee}_shifts.ics",
                mime="text/calendar",
            )
        else:
            st.info("No shifts found for this employee.")
    else:
        st.info("Please select an employee from the dropdown menu.")
else:
    st.info("Please upload an XLSM file using the sidebar to get started.")

# Footer
st.markdown("---")
st.markdown("Created with ❤️ by Your Friendly Big Tech Engineer")
