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
    return shifts


def create_ics_file(shifts, employee_name, definitions):
    cal = Calendar()
    for date, shift_code in shifts:
        event = Event()
        event.add("summary", f"Shift: {shift_code}")
        event.add("dtstart", date.date())
        event.add("dtend", date.date() + timedelta(days=1))
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
        all_shifts = detect_shifts(monthly_schedule, selected_employee)

        # Get unique shift codes
        unique_shift_codes = list(set([shift[1] for shift in all_shifts]))

        # Multiselect for shift codes
        st.subheader("Select Shift Codes")
        selected_shift_codes = st.multiselect(
            "Choose which shift codes to include",
            options=unique_shift_codes,
            default=unique_shift_codes,
        )

        # Filter shifts based on selected shift codes
        shifts = [shift for shift in all_shifts if shift[1] in selected_shift_codes]

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
            st.info("No shifts found for the selected shift codes.")
    else:
        st.info("Please select an employee from the dropdown menu.")
else:
    st.info("Please upload an XLSM file using the sidebar to get started.")

# Footer
st.markdown("---")
st.markdown("Created with ❤️ by Your Friendly Engineer")
st.caption("v.0.1.2")
