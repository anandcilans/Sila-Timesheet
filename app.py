import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import warnings

warnings.filterwarnings("ignore")


# Helper functions (defined first to avoid NameError)
def process_timesheet(df, start_date, end_date, ref_hours=9):
    """Process timesheet data from Excel"""

    if not start_date:
        raise ValueError("Could not extract date range from file")

    # Generate dates
    header_row = 2
    start_data_row = 4
    num_days = df.iloc[header_row, 2:].notna().sum()

    generated_dates = []
    current_date = start_date
    for i in range(num_days):
        generated_dates.append(current_date)
        current_date += timedelta(days=1)

    records = []

    for r in range(start_data_row, len(df)):
        emp_name = df.iloc[r, 1]
        if pd.isna(emp_name) or emp_name == "":
            continue

        for col_idx, (idx, day_abbr) in enumerate(df.iloc[header_row, 2:].items()):
            if col_idx >= len(generated_dates):
                break

            actual_date = generated_dates[col_idx].strftime("%d/%m/%Y")
            day_name = generated_dates[col_idx].strftime("%a")
            day_abbr_2 = day_name[:2]

            cell = df.iloc[r, idx]

            punch_in = ""
            punch_out = ""
            total_str = ""
            more_than_9 = ""
            less_than_9 = ""

            if not pd.isna(cell):
                times = [t for t in str(cell).replace("\n", " ").split() if ":" in t]
                if len(times) >= 2:
                    punch_in = times[0]
                    punch_out = times[-1]

                    try:
                        total = datetime.strptime(
                            punch_out, "%H:%M"
                        ) - datetime.strptime(punch_in, "%H:%M")
                        total_str = str(total)

                        total_seconds = total.total_seconds()
                        ref_hours_seconds = ref_hours * 3600
                        difference_seconds = total_seconds - ref_hours_seconds

                        if difference_seconds > 0:
                            sign = "+"
                            abs_seconds = difference_seconds
                        elif difference_seconds < 0:
                            sign = "-"
                            abs_seconds = abs(difference_seconds)
                        else:
                            sign = ""
                            abs_seconds = 0

                        hours = int(abs_seconds // 3600)
                        minutes = int((abs_seconds % 3600) // 60)
                        seconds = int(abs_seconds % 60)

                        if difference_seconds > 0:
                            more_than_9 = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                        elif difference_seconds < 0:
                            less_than_9 = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                    except:
                        pass

            records.append(
                [
                    emp_name,
                    actual_date,
                    day_abbr_2,
                    punch_in,
                    punch_out,
                    total_str,
                    more_than_9,
                    less_than_9,
                ]
            )

    return pd.DataFrame(
        records,
        columns=[
            "Employee Name",
            "Date",
            "Day",
            "Punch In",
            "Punch Out",
            "Total Hours",
            "More than 9Hours",
            "Less than 9 hours",
        ],
    )


def sum_timedelta(series):
    """Sum timedelta strings"""
    total = pd.Timedelta(0)
    for time_str in series:
        if time_str and isinstance(time_str, str):
            try:
                parts = time_str.split(":")
                hours = int(parts[0])
                minutes = int(parts[1])
                seconds = int(parts[2])
                total += pd.Timedelta(hours=hours, minutes=minutes, seconds=seconds)
            except:
                pass
    return str(total)


# Configure page
st.set_page_config(
    page_title="Timesheet Processor",
    page_icon="⏱️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for premium look
st.markdown(
    """
    <style>
        /* Main styling */
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        /* Header styling */
        .main-header {
            font-size: 2.5rem;
            font-weight: 700;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 1rem;
        }
        
        /* Metric cards */
        .metric-card {
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            padding: 1.5rem;
            border-radius: 10px;
            border-left: 4px solid #667eea;
        }
        
        /* Success message */
        .success-box {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
            padding: 1rem;
            border-radius: 5px;
            margin-top: 1rem;
        }
        
        /* Info box */
        .info-box {
            background-color: #d1ecf1;
            border: 1px solid #bee5eb;
            color: #0c5460;
            padding: 1rem;
            border-radius: 5px;
            margin-top: 1rem;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# Title
st.markdown('<p class="main-header">⏱️ Timesheet Processor</p>', unsafe_allow_html=True)

# # Sidebar
# with st.sidebar:
#     st.markdown("### 📋 Configuration")
#     st.markdown("---")

#     col1, col2 = st.columns([1, 1])
#     with col1:
#         st.markdown("**Version:** 2.0.1")
#     with col2:
#         st.markdown("**Status:** ✅ Active")

# Main content
tab1, tab2, tab3 = st.tabs(
    ["📁 Upload & Process", "📊 Preview Data", "📥 Download Results"]
)

with tab1:
    st.markdown("### Upload Your File")

    col1, col2 = st.columns([2, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Upload Excel file (.xls, .xlsx)",
            type=["xls", "xlsx"],
            help="Select the timesheet file with punch data",
        )

    with col2:
        st.info("💡 Tip: File should contain punch times in the standard format")

    if uploaded_file:
        st.success(f"✅ File uploaded: {uploaded_file.name}")

        try:
            # Read file
            df = pd.read_excel(uploaded_file, header=None)
            st.markdown("### Step 2: Processing Configuration")

            col1, col2 = st.columns([1, 1])
            with col1:
                reference_hours = st.number_input(
                    "Reference Working Hours",
                    min_value=1,
                    max_value=12,
                    value=9,
                    help="Standard working hours to compare against",
                )

            with col2:
                st.markdown("**Date Range**")
                try:
                    date_range_cell = df.iloc[1, 2]
                    date_parts = str(date_range_cell).split("~")
                    start_date = datetime.strptime(date_parts[0].strip(), "%d/%m/%Y")
                    end_date = datetime.strptime(date_parts[1].strip(), "%d/%m/%Y")
                    st.write(
                        f"{start_date.strftime('%d/%m/%Y')} → {end_date.strftime('%d/%m/%Y')}"
                    )
                except:
                    st.warning("Could not parse date range")
                    start_date = None

            if st.button(
                "🚀 Process File", key="process_btn", use_container_width=True
            ):
                with st.spinner("Processing timesheet..."):
                    try:
                        # Store in session state
                        st.session_state["processed_df"] = process_timesheet(
                            df, start_date, end_date, reference_hours
                        )
                        st.session_state["file_name"] = uploaded_file.name
                        st.markdown(
                            '<div class="success-box">✅ File processed successfully!</div>',
                            unsafe_allow_html=True,
                        )
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Error processing file: {str(e)}")

        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")

with tab2:
    if "processed_df" in st.session_state:
        processed_df = st.session_state["processed_df"]

        st.markdown("### 📊 Processed Data Preview")

        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(
                "👥 Total Employees",
                processed_df["Employee Name"].nunique(),
                delta=None,
            )

        with col2:
            st.metric("📅 Total Records", len(processed_df), delta=None)

        with col3:
            # Count days with overtime
            overtime = processed_df[processed_df["More than 9Hours"] != ""].shape[0]
            st.metric("⏱️ Overtime Days", overtime, delta=None)

        with col4:
            # Count days with undertime
            undertime = processed_df[processed_df["Less than 9 hours"] != ""].shape[0]
            st.metric("📉 Undertime Days", undertime, delta=None)

        st.markdown("---")

        # Filter options
        col1, col2 = st.columns([1, 1])

        with col1:
            selected_employees = st.multiselect(
                "Filter by Employee",
                processed_df["Employee Name"].unique(),
                default=(
                    processed_df["Employee Name"].unique()[:5]
                    if len(processed_df["Employee Name"].unique()) > 5
                    else processed_df["Employee Name"].unique()
                ),
            )

        with col2:
            sort_by = st.selectbox("Sort by", ["Employee Name", "Date", "Total Hours"])

        # Filter and sort
        filtered_df = processed_df[
            processed_df["Employee Name"].isin(selected_employees)
        ]
        if sort_by == "Date":
            filtered_df = filtered_df.sort_values("Date")
        else:
            filtered_df = filtered_df.sort_values(sort_by)

        # Display data with formatting
        st.dataframe(filtered_df, use_container_width=True, hide_index=True, height=400)

        # Detailed statistics
        st.markdown("### 📈 Employee Summary")

        summary_stats = (
            processed_df.groupby("Employee Name")
            .agg({"Punch In": "count", "Total Hours": lambda x: sum_timedelta(x)})
            .rename(columns={"Punch In": "Total Days"})
        )

        st.dataframe(summary_stats, use_container_width=True)

    else:
        st.info("📋 Process a file first to see the preview")

with tab3:
    if "processed_df" in st.session_state:
        processed_df = st.session_state["processed_df"]

        st.markdown("### 📥 Download Options")

        col1, col2 = st.columns([1, 1])

        with col1:
            # Download as Excel with employee-wise sheets
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                summary_data = []

                # Group by employee and create sheets
                for emp_name, emp_data in processed_df.groupby("Employee Name"):
                    emp_data_to_write = emp_data.reset_index(drop=True).copy()

                    # Calculate totals for summary
                    total_hours_sum = pd.Timedelta(0)
                    more_than_9_sum = pd.Timedelta(0)
                    less_than_9_sum = pd.Timedelta(0)

                    # Sum Total Hours
                    for time_str in emp_data_to_write["Total Hours"]:
                        if time_str and isinstance(time_str, str):
                            try:
                                parts = time_str.split(":")
                                hours = int(parts[0])
                                minutes = int(parts[1])
                                seconds = int(parts[2])
                                total_hours_sum += pd.Timedelta(
                                    hours=hours, minutes=minutes, seconds=seconds
                                )
                            except:
                                pass

                    # Sum More than 9 Hours
                    for time_str in emp_data_to_write["More than 9Hours"]:
                        if time_str and isinstance(time_str, str):
                            try:
                                parts = time_str.split(":")
                                hours = int(parts[0])
                                minutes = int(parts[1])
                                seconds = int(parts[2])
                                more_than_9_sum += pd.Timedelta(
                                    hours=hours, minutes=minutes, seconds=seconds
                                )
                            except:
                                pass

                    # Sum Less than 9 hours
                    for time_str in emp_data_to_write["Less than 9 hours"]:
                        if time_str and isinstance(time_str, str):
                            try:
                                parts = time_str.split(":")
                                hours = int(parts[0])
                                minutes = int(parts[1])
                                seconds = int(parts[2])
                                less_than_9_sum += pd.Timedelta(
                                    hours=hours, minutes=minutes, seconds=seconds
                                )
                            except:
                                pass

                    # Helper function to format timedelta
                    def timedelta_to_hours_format(td):
                        total_seconds = int(td.total_seconds())
                        hours = total_seconds // 3600
                        minutes = (total_seconds % 3600) // 60
                        seconds = total_seconds % 60
                        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

                    # Add to summary
                    summary_data.append(
                        {
                            "Employee Name": emp_name,
                            "Total Hours": timedelta_to_hours_format(total_hours_sum),
                            "More than 9 Hours": timedelta_to_hours_format(
                                more_than_9_sum
                            ),
                            "Less than 9 Hours": timedelta_to_hours_format(
                                less_than_9_sum
                            ),
                        }
                    )

                    # Create TOTAL row
                    sum_row_dict = {
                        "Employee Name": ["TOTAL"],
                        "Date": [""],
                        "Day": [""],
                        "Punch In": [""],
                        "Punch Out": [""],
                        "Total Hours": [timedelta_to_hours_format(total_hours_sum)],
                        "More than 9Hours": [
                            timedelta_to_hours_format(more_than_9_sum)
                        ],
                        "Less than 9 hours": [
                            timedelta_to_hours_format(less_than_9_sum)
                        ],
                    }
                    sum_row = pd.DataFrame(sum_row_dict)

                    # Concatenate data with sum row
                    emp_data_to_write = pd.concat(
                        [emp_data_to_write, sum_row], ignore_index=True
                    )

                    # Write to sheet (max 31 chars for sheet name)
                    sheet_name = str(emp_name)[:31]
                    emp_data_to_write.to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )

                # Create Summary sheet
                if summary_data:
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name="Summary", index=False)

            excel_buffer.seek(0)
            st.download_button(
                label="📊 Download as Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=f"Timesheet_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col2:
            # Download as CSV
            csv_buffer = processed_df.to_csv(index=False)
            st.download_button(
                label="📄 Download as CSV (.csv)",
                data=csv_buffer,
                file_name=f"Timesheet_{datetime.now().strftime('%d_%m_%Y')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

        st.markdown("---")
        st.markdown("### ℹ️ File Information")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.write(f"**Rows:** {len(processed_df)}")
        with col2:
            st.write(f"**Columns:** {len(processed_df.columns)}")
        with col3:
            st.write(f"**Generated:** {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    else:
        st.info("📋 Process a file first to download results")


# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #888; font-size: 0.85rem;'>
        <p>Cilans Timesheet Processor v1.0</p>
    </div>
""",
    unsafe_allow_html=True,
)
