import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta
import pytz
import base64

# Set page configuration
st.set_page_config(
    page_title="Hospital Night Shift Calendar", page_icon="👨‍⚕️", layout="wide"
)


@st.cache_data
def read_excel_file(uploaded_file):
    # Read the Excel file
    xls = pd.ExcelFile(uploaded_file)

    # Read relevant sheets
    monthly_schedule = pd.read_excel(xls, "Monatsplan")
    names = pd.read_excel(xls, "Namen")
    definitions = pd.read_excel(xls, "Dienstdefinition")

    return monthly_schedule, names, definitions


def detect_night_shifts(monthly_schedule, employee_name, employee_code, definitions):
    # Get the row for the employee
    employee_row = monthly_schedule[monthly_schedule["Name"] == employee_name].iloc[0]

    # Get the night shift codes
    night_shift_rows = definitions[
        definitions["Funktion"].str.contains("Nacht", case=False, na=False)
    ]
    night_shift_codes = []
    for persoff in night_shift_rows["PersOff"]:
        if isinstance(persoff, str):
            night_shift_codes.extend(
                [code.strip() for code in persoff.split(",") if code.strip()]
            )

    # Detect night shifts
    night_shifts = []
    for col in employee_row.index[2:]:  # Skip 'Name' and 'Tag' columns
        if employee_row[col] in night_shift_codes:
            night_shifts.append((col, employee_row[col]))

    return night_shifts


def create_ics_file(night_shifts, employee_name):
    cal = Calendar()

    for date, shift_code in night_shifts:
        event = Event()
        event.add("summary", f"Night Shift: {shift_code}")
        event.add(
            "dtstart", date.replace(hour=20, minute=0, second=0)
        )  # Assume night shift starts at 8 PM
        event.add(
            "dtend", (date + timedelta(days=1)).replace(hour=8, minute=0, second=0)
        )  # Assume night shift ends at 8 AM next day
        event.add("dtstamp", datetime.now(tz=pytz.UTC))
        event.add("location", "Hospital")

        cal.add_component(event)

    return cal.to_ical()


# App header
st.title("👨‍⚕️ Hospital Night Shift Calendar")
st.markdown("---")

# Sidebar for file upload
with st.sidebar:
    st.header("Upload File")
    uploaded_file = st.file_uploader("Choose an XLSM file", type="xlsm")

if uploaded_file is not None:
    monthly_schedule, names, definitions = read_excel_file(uploaded_file)

    st.subheader("Employee Selection")
    # Create dropdown for employee selection
    employee_names = names["Name"].tolist()
    selected_employee = st.selectbox(
        "Select your name", employee_names, key="employee_select"
    )

    if selected_employee:
        employee_code = names[names["Name"] == selected_employee]["Kuerzel"].iloc[0]

        night_shifts = detect_night_shifts(
            monthly_schedule, selected_employee, employee_code, definitions
        )

        st.markdown("---")
        st.subheader("Night Shifts Schedule")
        if night_shifts:
            df = pd.DataFrame(night_shifts, columns=["Date", "Shift Code"])
            df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m-%d")
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.write("No night shifts found for the selected employee.")

        st.subheader("Actions")
        if night_shifts:
            if st.button("Generate Calendar Invites", key="generate_button"):
                ics_data = create_ics_file(night_shifts, selected_employee)
                b64 = base64.b64encode(ics_data).decode()
                href = f'<a href="data:text/calendar;base64,{b64}" download="{selected_employee}_night_shifts.ics" class="btn">Download ICS file</a>'
                st.markdown(href, unsafe_allow_html=True)
        else:
            st.info("No night shifts found for this employee.")
    else:
        st.info("Please select an employee from the dropdown menu.")

else:
    st.info("Please upload an XLSM file using the sidebar to get started.")

# Footer
st.markdown("---")
st.markdown("Created with ❤️ by Your Friendly Big Tech Engineer")
