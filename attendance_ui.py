import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import tempfile

# === CORE FUNCTIONALITY ===

def group_days(attendance_row, match_status="P"):
    ranges = []
    start = None

    for day in range(1, 32):
        status = str(attendance_row.get(str(day), "")).strip().upper()
        if status == match_status:
            if start is None:
                start = day
        else:
            if start is not None:
                ranges.append((start, day - 1))
                start = None

    if start is not None:
        ranges.append((start, 31))

    return ranges

def convert_attendance_excel(file, shift_value, reason_value, explanation_value):
    df = pd.read_excel(file, dtype=str)
    df.columns = df.columns.map(lambda x: str(x).strip())

    if "ID" not in df.columns:
        st.error("❌ Column 'ID' not found in the uploaded file.")
        return None, None, None

    attendance_cols = [str(i) for i in range(1, 32)]
    bulk_rows, present_summary, leave_rows = [], [], []

    for _, row in df.iterrows():
        emp_id = str(row["ID"]).strip()
        if not emp_id:
            continue

        attendance_data = {day: row.get(day, "") for day in attendance_cols}

        # Present data
        present_ranges = group_days(attendance_data, match_status="P")
        total_present = sum(1 for v in attendance_data.values() if str(v).strip().upper() == "P")
        present_summary.append({
            "Employee": emp_id,
            "Total Present Days": total_present
        })

        for start_day, end_day in present_ranges:
            bulk_rows.append({
                "Employee": emp_id,
                "From Date": f"{start_day:02d}-05-2025",
                "To Date": f"{end_day:02d}-05-2025",
                "Holiday": 0,
                "Shift": shift_value,
                "Reason": reason_value,
                "Explanation": explanation_value
            })

        # Leave data
        leave_ranges = group_days(attendance_data, match_status="A")
        for start_day, end_day in leave_ranges:
            leave_rows.append({
                "Company": "AWOKE India Foundation",
                "Employee": emp_id,
                "From Date": f"2025-05-{start_day:02d} 00:00:00",
                "To Date": f"2025-05-{end_day:02d} 00:00:00",
                "Leave Type": "Casual Leave",
                "Status": "Approved"
            })

    return (
        pd.DataFrame(bulk_rows),
        pd.DataFrame(present_summary),
        pd.DataFrame(leave_rows)
    )

# === STREAMLIT UI ===

st.set_page_config(page_title="Attendance Converter", page_icon="📊")
st.title("📊 Attendance Tracker → Bulk Format, Leave Upload & Summary Generator")

uploaded_file = st.file_uploader("📥 Upload Attendance Tracker Excel File", type=["xlsx"])

st.markdown("### ⚙️ Conversion Settings")

shift_value = st.text_input("Shift Name", value="CFL Day")
reason_value = st.selectbox("Reason", options=["On Duty", "Work From Home"], index=0)
explanation_value = st.text_input("Explanation", value="Bulk Att. From Excel")

if uploaded_file:
    with st.spinner("⏳ Processing..."):
        bulk_df, summary_df, leave_df = convert_attendance_excel(
            uploaded_file, shift_value, reason_value, explanation_value
        )

        if bulk_df is not None and not bulk_df.empty:
            st.success(f"✅ Converted {len(bulk_df)} attendance records!")
            st.dataframe(bulk_df.head())
            st.markdown("### 📈 Present Summary")
            st.dataframe(summary_df)
            st.markdown("### 📝 Leave Upload Format")
            st.dataframe(leave_df)

            # Prepare ZIP file
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                with BytesIO() as b1:
                    bulk_df.to_excel(b1, index=False, sheet_name="Bulk Format", engine="openpyxl")
                    b1.seek(0)
                    zipf.writestr("Attendance_Upload.xlsx", b1.read())

                with BytesIO() as b3:
                    leave_df.to_excel(b3, index=False, sheet_name="Leave Upload", engine="openpyxl")
                    b3.seek(0)
                    zipf.writestr("Leave_Upload.xlsx", b3.read())

            zip_buffer.seek(0)

            st.download_button(
                label="📦 Download All Files as ZIP",
                data=zip_buffer,
                file_name="Attendance_Files.zip",
                mime="application/zip"
            )
        else:
            st.warning("⚠️ No data generated. Please check the uploaded file.")
