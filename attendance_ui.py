import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

# === CORE FUNCTIONALITY ===

def group_days(attendance_row, status_filter):
    ranges = []
    start = None
    last_status = None

    for day in range(1, 32):
        status = str(attendance_row.get(str(day), "")).strip().upper()
        if status == status_filter:
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
        return None, None, None

    attendance_cols = [str(i) for i in range(1, 32)]
    output_rows = []
    leave_rows = []
    present_summary = []

    for _, row in df.iterrows():
        emp_id = str(row["ID"]).strip()
        if not emp_id:
            continue

        attendance_data = {day: row.get(day, "") for day in attendance_cols}

        # Present record processing
        present_ranges = group_days(attendance_data, "P")
        total_present = sum(1 for v in attendance_data.values() if str(v).strip().upper() == "P")
        present_summary.append({
            "Employee": emp_id,
            "Total Present Days": total_present
        })
        for start_day, end_day, status in present_ranges:
            output_rows.append({
                "Employee": emp_id,
                "From Date": f"{start_day:02d}-05-2025",
                "To Date": f"{end_day:02d}-05-2025",
                "Holiday": 0,
                "Shift": shift_value,
                "Reason": reason_value,
                "Explanation": explanation_value
            })

        # Leave record processing
        leave_ranges = group_days(attendance_data, "L")
        for start_day, end_day, status in leave_ranges:
            leave_rows.append({
                "Company": "AWOKE India Foundation",
                "Employee": emp_id,
                "From Date": f"2025-05-{start_day:02d} 00:00:00",
                "To Date": f"2025-05-{end_day:02d} 00:00:00",
                "Leave Type": "Casual Leave",
                "Status": "Approved"
            })

    return pd.DataFrame(output_rows), pd.DataFrame(present_summary), pd.DataFrame(leave_rows)

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
        result_df, summary_df, leave_df = convert_attendance_excel(uploaded_file, shift_value, reason_value, explanation_value)

        if result_df is not None and not result_df.empty:
            st.success(f"✅ File converted successfully with {len(result_df)} attendance records!")
            st.dataframe(result_df.head())
            st.markdown("### 📈 Total Present Summary")
            st.dataframe(summary_df)

            if not leave_df.empty:
                st.markdown("### 📋 Leave Records Found")
                st.dataframe(leave_df)

            # Create ZIP with 3 Excel files
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                # 1. Bulk Format
                bulk_bytes = BytesIO()
                with pd.ExcelWriter(bulk_bytes, engine="openpyxl") as writer:
                    result_df.to_excel(writer, index=False)
                zf.writestr("Bulk Format.xlsx", bulk_bytes.getvalue())

                # 2. Present Summary
                summary_bytes = BytesIO()
                with pd.ExcelWriter(summary_bytes, engine="openpyxl") as writer:
                    summary_df.to_excel(writer, index=False)
                zf.writestr("Present Summary.xlsx", summary_bytes.getvalue())

                # 3. Leave Records
                leave_bytes = BytesIO()
                with pd.ExcelWriter(leave_bytes, engine="openpyxl") as writer:
                    leave_df.to_excel(writer, index=False)
                zf.writestr("Leave Records.xlsx", leave_bytes.getvalue())

            zip_buffer.seek(0)
            st.download_button(
                label="📦 Download All as ZIP",
                data=zip_buffer,
                file_name="Converted_Attendance_Package.zip",
                mime="application/zip"
            )
        else:
            st.warning("⚠️ No data to export. Please check your input file.")

# === 📄 Sample Template ===
# === 📄 Sample Template ===
st.markdown("---")
st.markdown("### 📄 Need a sample file?")

try:
    with open("Supaul Region Attendance Tracker May.xlsx", "rb") as f:
        sample_bytes = f.read()
        sample_file = BytesIO(sample_bytes)

    st.download_button(
        label="📥 Download Real Sample File (Excel)",
        data=sample_file,
        file_name="Attendance_Tracker_Sample.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except FileNotFoundError:
    st.error("❌ Sample file not found. Please make sure it's in the app directory.")