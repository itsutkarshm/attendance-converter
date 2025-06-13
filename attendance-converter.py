import pandas as pd

INPUT_FILE = "Supaul Region Attendance Tracker May.xlsx"
OUTPUT_FILE = "Converted Attendance Data.xlsx"

def group_days(attendance_row):
    ranges = []
    start = None
    last_status = None

    for day in range(1, 32):
        status = str(attendance_row.get(str(day), "")).strip().upper()
        if status in ["P"]:
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

def convert_tracker_to_format():
    df = pd.read_excel(INPUT_FILE, dtype=str)
    
    # Normalize column names
    df.columns = df.columns.map(lambda x: str(x).strip())

    # Use the exact column name
    if "ID" not in df.columns:
        print("❌ Column 'ID' not found in Excel.")
        return

    attendance_cols = [str(i) for i in range(1, 32)]
    output_rows = []

    for _, row in df.iterrows():
        emp_id = str(row["ID"]).strip()
        if not emp_id:
            continue

        attendance_data = {day: row.get(day, "") for day in attendance_cols}
        date_ranges = group_days(attendance_data)

        for start_day, end_day, status in date_ranges:
            output_rows.append({
                "Employee": emp_id,
                "From Date": f"{start_day:02d}-05-2025",
                "To Date": f"{end_day:02d}-05-2025",
                "Holiday": 0,
                "Shift": "CFL Day",
                "Reason": "On Duty",
                "Explanation": "Bulk Att. From Excel"
            })

    if not output_rows:
        print("⚠️ No attendance data found to write.")
    else:
        output_df = pd.DataFrame(output_rows)
        output_df.to_excel(OUTPUT_FILE, index=False)
        print(f"✅ Converted file saved as: {OUTPUT_FILE} with {len(output_rows)} rows.")

if __name__ == "__main__":
    convert_tracker_to_format()
