import streamlit as st
import pandas as pd
from io import BytesIO

# === CORE FUNCTIONALITY ===

def group_days(attendance_row):
    ranges = []
    start = None
    last_status = None

    for day in range(1, 32):
        status = str(attendance_row.get(str(day), "")).strip().upper()
        if status == "P":
            if start is None:
                start = day
                last_status = status
            elif status != last_status:
                ranges.append((start, day - 1, last_status))
                start = day
                last_status = status
        else:
            if start is not None:
                ranges.append((start, day - 1, last_status))
                start = None
                last_status = None

    if start is not None:
        ranges.append((start, 31, last_status))
    return ranges

def convert_attendance_excel(file, shift_value, reason_value, explanation_value):
    df = pd.read_excel(file, dtype=str)
    df.columns = df.columns.map(lambda x: str(x).strip())

    if "ID" not in df.columns:
        st.error("❌ Column 'ID' not found in the uploaded file.")
        return None, None

    attendance_cols = [str(i) for i in range(1, 32)]
    output_rows = []
    present_summary = []

    for _, row in df.iterrows():
        emp_id = str(row["ID"]).strip()
        if not emp_id:
            continue

        attendance_data = {day: row.get(day, "") for day in attendance_cols}
        date_ranges = group_days(attendance_data)

        # Count Present days
        total_present = sum(1 for v in attendance_data.values() if str(v).strip().upper() == "P")
        present_summary.append({
            "Employee": emp_id,
            "Total Present Days": total_present
        })

        for start_day, end_day, status in date_ranges:
            output_rows.append({
                "Employee": emp_id,
                "From Date": f"{start_day:02d}-05-2025",
                "To Date": f"{end_day:02d}-05-2025",
                "Holiday": 0,
                "Shift": shift_value,
                "Reason": reason_value,
                "Explanation": explanation_value
            })

    return pd.DataFrame(output_rows), pd.DataFrame(present_summary)

# === STREAMLIT UI ===

st.set_page_config(page_title="Attendance Converter", page_icon="📊")
st.title("📊 Attendance Tracker → Bulk Update Format Converter")

uploaded_file = st.file_uploader("📥 Upload Attendance Tracker Excel File", type=["xlsx"])

st.markdown("### ⚙️ Conversion Settings")

shift_value = st.text_input("Shift Name", value="CFL Day")
reason_value = st.selectbox("Reason", options=["On Duty", "Work From Home"], index=0)
explanation_value = st.text_input("Explanation", value="Bulk Att. From Excel")

if uploaded_file:
    with st.spinner("⏳ Processing..."):
        result_df, summary_df = convert_attendance_excel(uploaded_file, shift_value, reason_value, explanation_value)
        if result_df is not None and not result_df.empty:
            st.success(f"✅ File converted successfully with {len(result_df)} attendance records!")
            st.dataframe(result_df.head())
            st.markdown("### 📈 Total Present Summary")
            st.dataframe(summary_df)

            # Prepare Excel download with 2 sheets
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name="Bulk Format", index=False)
                summary_df.to_excel(writer, sheet_name="Present Summary", index=False)
            output.seek(0)

            st.download_button(
                label="📤 Download Excel File",
                data=output,
                file_name="Converted_Attendance_Data_With_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ No data to export. Please check your input file.")
