
import streamlit as st
import pandas as pd
from matplotlib import pyplot as plot
import re
import os
import openpyxl
from copy import copy
from datetime import datetime, date, timedelta
from openpyxl import load_workbook
from io import BytesIO
import plotly.express as px
import base64
import json
from langchain_openai import ChatOpenAI
from gmail_automation import (
    process_new_emails, start_background_poller, gmail_setup_status,
    get_pending_job_notifications,
)


st.set_page_config(
    page_title="Automation Workflow System",
    page_icon="🟠",
    layout="wide",
    initial_sidebar_state="expanded",
)


# -----------------------------------------------------
# Performance helpers
# -----------------------------------------------------
@st.cache_data(ttl=10, show_spinner=False)
def cached_excel_dataframe(file_path, modified_time):
    """Cache pandas Excel reads briefly to avoid rereading the workbook on every rerun."""
    return pd.read_excel(file_path)


@st.cache_data(show_spinner=False)
def get_background_image_base64(image_path, modified_time):
    """Read and encode the background image once per file version."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("ascii")


def clear_workflow_cache():
    """Call after a successful workbook update so the next read gets fresh data."""
    cached_excel_dataframe.clear()




# -----------------------------------------------------
# DashboardStatus 
# -----------------------------------------------------

def DashboardStatus():

    # Job Status heading and Refresh button
    col_status, col_refresh = st.columns([6, 1])

    with col_status:
        st.subheader("Job Status")

    with col_refresh:
        if st.button(
            "🔄 Refresh",
            key="dashboard_refresh",
            use_container_width=True
        ):
            # Close workflow forms
            st.session_state.show_new_job_form = False
            st.session_state.show_assign_form = False
            st.session_state.show_qc_form = False
            st.session_state.show_move_form = False

            # Close other sections
            st.session_state.deadline_category = None
            st.session_state.show_reports_export = False
            st.session_state.show_employee_workload = False

            # Clear search and chatbot fields
            st.session_state.help_text = ""
            st.session_state.job_search_text = ""
            # Streamlit already reruns on button click; avoid a second rerun.

    workbook_path = "workflow.xlsx"

    def get_status_counts(ws):
        header_row = None
        headers = []

        # Find the header row within the first 20 rows
        for row_num in range(1, min(ws.max_row, 20) + 1):
            row_values = [
                str(ws.cell(row_num, col).value or "").strip()
                for col in range(1, ws.max_column + 1)
            ]

            normalized = [value.casefold() for value in row_values]

            if "status" in normalized:
                header_row = row_num
                headers = normalized
                break

        counts = {
            "tbi": 0,
            "inprogress": 0,
            "qc": 0,
            "completed": 0
        }

        if header_row is None:
            return counts

        status_index = headers.index("status") + 1

        for row_num in range(header_row + 1, ws.max_row + 1):
            status_value = ws.cell(
                row_num, status_index
            ).value

            if status_value is None:
                continue

            status = str(status_value).strip().casefold()

            if status in ("delivered", "completed"):
                counts["completed"] += 1
            elif status in ("tbi", "inprogress", "qc"):
                counts[status] += 1

        return counts

    try:
        # Read-only access protects the original workbook
        wb = load_workbook(
            workbook_path,
            read_only=False,
            data_only=True
        )

        inprogress_counts = {
            "tbi": 0,
            "inprogress": 0,
            "qc": 0,
            "completed": 0
        }

        completed_counts = {
            "tbi": 0,
            "inprogress": 0,
            "qc": 0,
            "completed": 0
        }

        if "Inprogress" in wb.sheetnames:
            inprogress_counts = get_status_counts(
                wb["Inprogress"]
            )

        if "Completed Jobs" in wb.sheetnames:
            completed_counts = get_status_counts(
                wb["Completed Jobs"]
            )

        wb.close()

        # Combine status counts
        tbi_count = inprogress_counts["tbi"]
        inprogress_count = inprogress_counts["inprogress"]
        qc_count = inprogress_counts["qc"]

        completed_count = (
            inprogress_counts["completed"]
            + completed_counts["completed"]
        )

        # Display summary cards
        st.markdown('<div class="section-heading"><span class="section-icon orange">▦</span>Job Status Summary</div>', unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            with st.container(key="status_tbi_card"):
                st.metric("🕒  TBI", tbi_count)
        with col2:
            with st.container(key="status_inprogress_card"):
                st.metric("🔄  Inprogress", inprogress_count)
        with col3:
            with st.container(key="status_qc_card"):
                st.metric("📋  QC", qc_count)
        with col4:
            with st.container(key="status_completed_card"):
                st.metric("✅  Completed", completed_count)

        st.divider()

        # Prepare chart data
        st.markdown("### Job Status Overview")

        status_chart_data = pd.DataFrame({
            "Job Status": [
                "TBI",
                "Inprogress",
                "QC",
                "Completed"
            ],
            "Job Count": [
                tbi_count,
                inprogress_count,
                qc_count,
                completed_count
            ]
        })

        # Different colors for each status
        status_colors = {
            "TBI": "#F39C12",
            "Inprogress": "#3498DB",
            "QC": "#9B59B6",
            "Completed": "#27AE60"
        }

        fig = px.bar(
            status_chart_data,
            x="Job Status",
            y="Job Count",
            color="Job Status",
            color_discrete_map=status_colors,
            text="Job Count",
            category_orders={
                "Job Status": [
                    "TBI",
                    "Inprogress",
                    "QC",
                    "Completed"
                ]
            },
            title="Current Job Status"
        )

        # Show the job count above each bar
        fig.update_traces(
            textposition="outside",
            cliponaxis=False
        )

        # Fix the Y-axis from 0 to 100
        fig.update_layout(
            xaxis_title="Job Status",
            yaxis_title="Number of Jobs",
            showlegend=False,
            height=400,
            margin=dict(l=20, r=20, t=60, b=20),
            yaxis=dict(
                range=[0, 110],
                tickmode="linear",
                tick0=0,
                dtick=10,
                fixedrange=True
            )
        )

        # Display chart
        st.plotly_chart(
            fig,
            use_container_width=True
        )

    except FileNotFoundError:
        st.error(
            "workflow.xlsx was not found. "
            "Please check the file path."
        )

    except Exception as e:
        st.error(f"Unable to load Job Status: {e}")

#-------------------------------------------------------Dashboard Job Status end-----------------------





# =====================================================
# New Job Button
# =====================================================

FILE_PATH = "workflow.xlsx"
SOURCE_SHEET = "Inprogress"


def NewJob():

    # =====================================================
    # SESSION STATE
    # =====================================================

    if "show_new_job_form" not in st.session_state:
        st.session_state.show_new_job_form = False


    # =====================================================
    # SHOW ONLY NEW JOB BUTTON INITIALLY
    # =====================================================

    if not st.session_state.show_new_job_form:

        button_left, button_center, button_right = st.columns([1, 2, 1])
        with button_center:
            if st.button(
                "➕ New Job",
                key="new_job_button",
                use_container_width=True,
                type="primary"
            ):
                st.session_state.show_new_job_form = (
                    not st.session_state.show_new_job_form
                )
                st.rerun()

        return


    # =====================================================
    # NEW JOB FORM
    # =====================================================

    st.subheader("New Job Entry")


    # =====================================================
    # DATE
    # =====================================================

    job_date = st.date_input(
        "Date",
        value=date.today()
    )


    # =====================================================
    # TWO COLUMNS
    # =====================================================

    col1, col2 = st.columns(2)


    # =====================================================
    # LEFT COLUMN
    # =====================================================

    with col1:

        # -----------------------------------------------
        # Shift
        # -----------------------------------------------

        shift = st.selectbox(
            "Shift",
            [
                "Shift-1",
                "Shift-2",
                "Shift-3"
            ]
        )


        # -----------------------------------------------
        # Shift Lead
        # -----------------------------------------------

        shift_lead = st.text_input(
            "Shift Lead"
        )


        # -----------------------------------------------
        # Requestor / Client Name
        # -----------------------------------------------

        requestor = st.text_input(
            "Requestor/Client Name"
        )


        # -----------------------------------------------
        # Project Number
        # -----------------------------------------------

        project_number = st.text_input(
            "Project Number"
        )


        # -----------------------------------------------
        # WO Number
        # -----------------------------------------------

        wo_number = st.text_input(
            "WO Number"
        )


    # =====================================================
    # RIGHT COLUMN
    # =====================================================

    with col2:

        # -----------------------------------------------
        # Status
        # -----------------------------------------------

        status = st.selectbox(
            "Status",
            [
                "Inprogress",
                "TBI",
                "QC"
            ]
        )


        # -----------------------------------------------
        # Job Name
        # -----------------------------------------------

        job_name = st.text_input(
            "Job Name"
        )


        # -----------------------------------------------
        # Job Type
        # -----------------------------------------------

        job_type = st.selectbox(
            "Job Type",
            [
                "Print",
                "Posting",
                "Web",
                "Headshot"
            ]
        )


        # -----------------------------------------------
        # Region
        # -----------------------------------------------

        region = st.selectbox(
            "Region",
            [
                "Newyork",
                "Mumbai",
                "London"
            ]
        )


    # =====================================================
    # ADD NEW JOB BUTTON
    # =====================================================

    if st.button(
        "Add New Job",
        type="primary"
    ):

        # =================================================
        # VALIDATION
        # =================================================

        if not job_name.strip():

            st.warning(
                "Please enter Job Name."
            )

            return


        if not requestor.strip():

            st.warning(
                "Please enter Requestor/Client Name."
            )

            return


        if not project_number.strip():

            st.warning(
                "Please enter Project Number."
            )

            return


        if not wo_number.strip():

            st.warning(
                "Please enter WO Number."
            )

            return


        if not shift_lead.strip():

            st.warning(
                "Please enter Shift Lead."
            )

            return


        # =================================================
        # CAPTURE RECEIPT TIME
        # =================================================
        # This is captured only when Add New Job
        # is clicked.

        receipt_time = datetime.now()


        # =================================================
        # LOAD EXISTING WORKBOOK
        # =================================================

        try:

            wb = openpyxl.load_workbook(
                FILE_PATH
            )

        except FileNotFoundError:

            st.error(
                f"Excel file '{FILE_PATH}' was not found."
            )

            return


        # =================================================
        # CHECK SHEET
        # =================================================

        if SOURCE_SHEET not in wb.sheetnames:

            st.error(
                f"Sheet '{SOURCE_SHEET}' was not found "
                "in the Excel workbook."
            )

            return


        ws = wb[SOURCE_SHEET]


        # =================================================
        # READ HEADERS
        # =================================================

        headers = {}

        for cell in ws[1]:

            if cell.value is not None:

                headers[
                    str(cell.value).strip()
                ] = cell.column


        # =================================================
        # REQUIRED COLUMNS
        # =================================================

        required_columns = [
            "Date",
            "Receipt Time",
            "Shift",
            "Shift Lead",
            "Requestor/Client Name",
            "Project Number",
            "WO Number",
            "Status",
            "Job Name",
            "Job Type",
            "Region"
        ]


        # =================================================
        # CHECK MISSING COLUMNS
        # =================================================

        missing_columns = [
            column
            for column in required_columns
            if column not in headers
        ]


        if missing_columns:

            st.error(
                "The following columns are missing "
                "from the Inprogress sheet:\n\n"
                + ", ".join(missing_columns)
            )

            return


        # =================================================
        # FIND NEW ROW
        # =================================================

        new_row = ws.max_row + 1


        # =================================================
        # COPY FORMATTING FROM PREVIOUS ROW
        # =================================================

        previous_row = new_row - 1


        if previous_row >= 2:

            for column in range(
                1,
                ws.max_column + 1
            ):

                source_cell = ws.cell(
                    row=previous_row,
                    column=column
                )

                destination_cell = ws.cell(
                    row=new_row,
                    column=column
                )


                # -----------------------------------------
                # Copy font
                # -----------------------------------------

                if source_cell.has_style:

                    destination_cell.font = copy(
                        source_cell.font
                    )


                    # -------------------------------------
                    # Copy fill
                    # -------------------------------------

                    destination_cell.fill = copy(
                        source_cell.fill
                    )


                    # -------------------------------------
                    # Copy border
                    # -------------------------------------

                    destination_cell.border = copy(
                        source_cell.border
                    )


                    # -------------------------------------
                    # Copy alignment
                    # -------------------------------------

                    destination_cell.alignment = copy(
                        source_cell.alignment
                    )


                    # -------------------------------------
                    # Copy number format
                    # -------------------------------------

                    destination_cell.number_format = (
                        source_cell.number_format
                    )


                    # -------------------------------------
                    # Copy protection
                    # -------------------------------------

                    destination_cell.protection = copy(
                        source_cell.protection
                    )


            # ---------------------------------------------
            # Copy row height
            # ---------------------------------------------

            previous_height = ws.row_dimensions[
                previous_row
            ].height


            if previous_height is not None:

                ws.row_dimensions[
                    new_row
                ].height = previous_height


        # =================================================
        # WRITE NEW JOB DATA
        # =================================================

        # -----------------------------------------------
        # Date
        # -----------------------------------------------

        date_cell = ws.cell(
            row=new_row,
            column=headers["Date"]
        )

        date_cell.value = job_date

        date_cell.number_format = "dd-mm-yyyy"


        # -----------------------------------------------
        # Receipt Time
        # -----------------------------------------------

        receipt_cell = ws.cell(
            row=new_row,
            column=headers["Receipt Time"]
        )

        receipt_cell.value = receipt_time

        receipt_cell.number_format = (
            "dd-mm-yyyy hh:mm:ss"
        )


        # -----------------------------------------------
        # Shift
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["Shift"]
        ).value = shift


        # -----------------------------------------------
        # Shift Lead
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["Shift Lead"]
        ).value = shift_lead


        # -----------------------------------------------
        # Requestor / Client Name
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers[
                "Requestor/Client Name"
            ]
        ).value = requestor


        # -----------------------------------------------
        # Project Number
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["Project Number"]
        ).value = project_number


        # -----------------------------------------------
        # WO Number
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["WO Number"]
        ).value = wo_number


        # -----------------------------------------------
        # Status
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["Status"]
        ).value = status


        # -----------------------------------------------
        # Job Name
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["Job Name"]
        ).value = job_name


        # -----------------------------------------------
        # Job Type
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["Job Type"]
        ).value = job_type


        # -----------------------------------------------
        # Region
        # -----------------------------------------------

        ws.cell(
            row=new_row,
            column=headers["Region"]
        ).value = region


        # =================================================
        # SAVE WORKBOOK
        # =================================================

        try:

            wb.save(FILE_PATH)


            clear_workflow_cache()
        except PermissionError:

            st.error(
                "Unable to save workflow.xlsx. "
                "Please close the Excel file and try again."
            )

            return


        # =================================================
        # SUCCESS MESSAGE
        # =================================================

        st.success(
            f"✅ New job '{job_name}' added successfully!"
        )


        # =================================================
        # CLOSE NEW JOB FORM
        # =================================================

        st.session_state.show_new_job_form = False

        st.rerun()

#----------------------------New Job end----------------------------





def AssignJob():
    workbook_path = "workflow.xlsx"

    # Full-width Assign Job button
    if "show_assign_form" not in st.session_state:
        st.session_state.show_assign_form = False

    button_left, button_center, button_right = st.columns([1, 2, 1])
    with button_center:
        if st.button(
            "👤 Assign Job",
            key="assign_job_button",
            use_container_width=True,
            type="primary"
        ):
            st.session_state.show_assign_form = (
                not st.session_state.show_assign_form
            )

    # Show the assignment form only when the button is clicked
    if not st.session_state.show_assign_form:
        return

    try:
        wb = load_workbook(workbook_path)
        if "Inprogress" not in wb.sheetnames:
            wb.close()
            st.error("The Inprogress sheet was not found.")
            return

        ws = wb["Inprogress"]

        def normalize(value):
            return str(value or "").strip().casefold().replace("_", " ")

        # Find the header row
        header_row = None
        headers = []

        for row_num in range(1, min(ws.max_row, 20) + 1):
            row_values = [
                ws.cell(row_num, col).value
                for col in range(1, ws.max_column + 1)
            ]

            normalized = [normalize(value) for value in row_values]

            if "job name" in normalized and "status" in normalized:
                header_row = row_num
                headers = [
                    str(value or "").strip()
                    for value in row_values
                ]
                break

        if header_row is None:
            wb.close()
            st.error("Could not find the Job Name and Status headers.")
            return

        # Map column names to Excel column numbers
        columns = {
            normalize(header): index + 1
            for index, header in enumerate(headers)
            if normalize(header)
        }

        def find_column(*names):
            for name in names:
                if normalize(name) in columns:
                    return columns[normalize(name)]
            return None

        job_col = find_column("Job Name")
        status_col = find_column("Status")
        operator_col = find_column(
            "Operator", "Assigned Operator", "Employee"
        )
        receipt_col = find_column("Receipt Time")
        initiation_col = find_column("Initiation Time")
        deadline_col = find_column("Deadline IST", "Deadline")

        required_columns = {
            "Job Name": job_col,
            "Status": status_col,
            "Operator": operator_col,
            "Receipt Time": receipt_col,
            "Initiation Time": initiation_col,
            "Deadline IST": deadline_col
        }

        missing = [
            name for name, column in required_columns.items()
            if column is None
        ]

        if missing:
            wb.close()
            st.error(
                "Missing required columns in the Inprogress sheet: "
                + ", ".join(missing)
            )
            return

        # Find jobs with TBI status
        tbi_jobs = []

        for row_num in range(header_row + 1, ws.max_row + 1):
            job_name = ws.cell(row_num, job_col).value
            status = ws.cell(row_num, status_col).value

            if (
                job_name is not None
                and str(status or "").strip().casefold() == "tbi"
            ):
                tbi_jobs.append({
                    "row": row_num,
                    "name": str(job_name).strip()
                })

        if not tbi_jobs:
            wb.close()
            st.info("No TBI jobs are available for assignment.")
            return

        st.subheader("Assign Job to Operator")

        # Include row number in the option to distinguish duplicate names
        selected_job = st.selectbox(
            "Select TBI Job",
            options=tbi_jobs,
            format_func=lambda job: job["name"],
            key="assign_selected_job"
        )

        operators = [
            "Bala",
            "Operator 1",
            "Operator 2",
            "Operator 3",
            "Operator 4"
        ]

        selected_operator = st.selectbox(
            "Select Operator",
            operators,
            key="assign_selected_operator"
        )

        if st.button(
            "🚀 Initiate Job",
            key="initiate_selected_job",
            use_container_width=True,
            type="primary"
        ):
            row_num = selected_job["row"]

            # Recheck status before changing the row
            current_status = str(
                ws.cell(row_num, status_col).value or ""
            ).strip().casefold()

            if current_status != "tbi":
                wb.close()
                st.error(
                    "This job is no longer in TBI status. "
                    "Refresh and select another job."
                )
                return

            receipt_time = ws.cell(row_num, receipt_col).value

            if not isinstance(receipt_time, datetime):
                wb.close()
                st.error(
                    "Receipt Time is missing or is not a valid Excel "
                    "date/time. Please correct it before assigning the job."
                )
                return

            # Set initiation time when the job is assigned
            initiation_time = datetime.now()

            # Deadline is 24 hours after Receipt Time
            deadline_time = receipt_time + timedelta(hours=24)

            ws.cell(row_num, operator_col).value = selected_operator
            ws.cell(row_num, initiation_col).value = initiation_time
            ws.cell(row_num, deadline_col).value = deadline_time
            ws.cell(row_num, status_col).value = "Inprogress"

            # Preserve readable date/time formatting
            ws.cell(row_num, initiation_col).number_format = (
                "dd-mm-yyyy hh:mm:ss AM/PM"
            )
            ws.cell(row_num, deadline_col).number_format = (
                "dd-mm-yyyy hh:mm:ss AM/PM"
            )

            wb.save(workbook_path)

            clear_workflow_cache()
            wb.close()

            st.session_state.show_assign_form = False

            st.success(
                f"Job '{selected_job['name']}' assigned to "
                f"{selected_operator} successfully."
            )

            st.rerun()

        wb.close()

    except FileNotFoundError:
        st.error("workflow.xlsx was not found. Check the file path.")

    except PermissionError:
        st.error(
            "Cannot update workflow.xlsx. Close the workbook in Excel "
            "and try again."
        )

    except Exception as e:
        st.error(f"Unable to assign job: {e}")
        

#----------------------------Assign Job end----------------------------




def MoveJobToQC():
    from openpyxl import load_workbook
    import streamlit as st

    workbook_path = "workflow.xlsx"

    # Initialize form visibility
    if "show_qc_form" not in st.session_state:
        st.session_state.show_qc_form = False

    # Full-width Move to QC button
    button_left, button_center, button_right = st.columns([1, 2, 1])
    with button_center:
        if st.button(
            "✅ Move to QC",
            key="move_to_qc_button",
            use_container_width=True,
            type="primary"
        ):
            st.session_state.show_qc_form = (
                not st.session_state.show_qc_form
            )

    # Show the form only after clicking the button
    if not st.session_state.show_qc_form:
        return

    wb = None

    try:
        wb = load_workbook(workbook_path)

        # Check required worksheets
        if "Inprogress" not in wb.sheetnames:
            st.error("The Inprogress sheet was not found.")
            return

        if "Dropdown" not in wb.sheetnames:
            st.error("The Dropdown sheet was not found.")
            return

        ws = wb["Inprogress"]
        dropdown_ws = wb["Dropdown"]

        def normalize(value):
            return str(value or "").strip().casefold().replace("_", " ")

        # Find the Inprogress header row
        header_row = None
        headers = []

        for row_num in range(1, min(ws.max_row, 20) + 1):
            row_values = [
                ws.cell(row_num, col).value
                for col in range(1, ws.max_column + 1)
            ]

            normalized = [normalize(value) for value in row_values]

            if "job name" in normalized and "status" in normalized:
                header_row = row_num
                headers = [
                    str(value or "").strip()
                    for value in row_values
                ]
                break

        if header_row is None:
            st.error(
                "Could not find the Job Name and Status headers "
                "in the Inprogress sheet."
            )
            return

        # Map column names to Excel column numbers
        columns = {
            normalize(header): index + 1
            for index, header in enumerate(headers)
            if normalize(header)
        }

        job_col = columns.get("job name")
        status_col = columns.get("status")
        qc_col = columns.get("qc")

        if job_col is None or status_col is None:
            st.error("The Job Name or Status column is missing.")
            return

        if qc_col is None:
            st.error(
                "The QC column was not found in the Inprogress sheet."
            )
            return

        # Read QC names from the Dropdown sheet
        qc_header_row = None
        qc_dropdown_col = None

        for row_num in range(1, min(dropdown_ws.max_row, 20) + 1):
            for col_num in range(1, dropdown_ws.max_column + 1):
                value = dropdown_ws.cell(row_num, col_num).value

                if normalize(value) == "qc":
                    qc_header_row = row_num
                    qc_dropdown_col = col_num
                    break

            if qc_header_row is not None:
                break

        if qc_header_row is None:
            st.error(
                "Could not find the QC column header in the Dropdown sheet."
            )
            return

        qc_names = []

        for row_num in range(qc_header_row + 1, dropdown_ws.max_row + 1):
            value = dropdown_ws.cell(
                row_num, qc_dropdown_col
            ).value

            if value is None:
                continue

            qc_name = str(value).strip()

            if (
                qc_name
                and qc_name.casefold() not in {
                    "none", "nan", "not assigned", "unassigned"
                }
                and qc_name not in qc_names
            ):
                qc_names.append(qc_name)

        if not qc_names:
            st.warning(
                "No QC names were found under the QC column "
                "in the Dropdown sheet."
            )
            return

        # Find jobs whose current status is Inprogress
        available_jobs = []

        for row_num in range(header_row + 1, ws.max_row + 1):
            job_name = ws.cell(row_num, job_col).value
            status = ws.cell(row_num, status_col).value

            if (
                job_name is not None
                and str(status or "").strip().casefold() == "inprogress"
            ):
                available_jobs.append({
                    "row": row_num,
                    "name": str(job_name).strip()
                })

        if not available_jobs:
            st.info("No Inprogress jobs are available to move to QC.")
            return

        st.subheader("Move Job to QC")

        # Select a job; row number distinguishes duplicate job names
        selected_job = st.selectbox(
            "Select Inprogress Job",
            options=available_jobs,
            format_func=lambda job: job["name"],
            key="selected_job_for_qc"
        )

        # Select QC name from the Dropdown worksheet
        selected_qc = st.selectbox(
            "Select QC Name",
            options=qc_names,
            key="selected_qc_name"
        )

        # Confirm the status update
        if st.button(
            "Confirm Move to QC",
            key="confirm_move_to_qc",
            use_container_width=True,
            type="primary"
        ):
            row_num = selected_job["row"]

            # Recheck status before updating
            current_status = str(
                ws.cell(row_num, status_col).value or ""
            ).strip().casefold()

            if current_status != "inprogress":
                st.error(
                    "This job is no longer Inprogress. "
                    "Please refresh and select another job."
                )
                return

            # Update only the QC and Status cells
            ws.cell(row_num, qc_col).value = selected_qc
            ws.cell(row_num, status_col).value = "QC"

            wb.save(workbook_path)


            clear_workflow_cache()
            st.session_state.show_qc_form = False

            st.success(
                f"Job '{selected_job['name']}' moved to QC "
                f"and assigned to {selected_qc}."
            )

            st.rerun()

    except FileNotFoundError:
        st.error(
            "workflow.xlsx was not found. Check the file path."
        )

    except PermissionError:
        st.error(
            "Cannot update workflow.xlsx. Please close the workbook "
            "in Excel and try again."
        )

    except Exception as e:
        st.error(f"Unable to move job to QC: {e}")

    finally:
        if wb is not None:
            wb.close()


#----------------------------Move to QC end----------------------------



FILE_PATH = "workflow.xlsx"


def MoveJobToCompleted():

    workbook_path = "workflow.xlsx"

    if "show_move_form" not in st.session_state:
        st.session_state.show_move_form = False

    # Full-width Move Job to Completed button
    button_left, button_center, button_right = st.columns([1, 2, 1])
    with button_center:
        if st.button(
            "🏁 Move Job to Completed",
            key="move_job_completed_button",
            use_container_width=True,
            type="primary"
        ):
            st.session_state.show_move_form = (
                not st.session_state.show_move_form
            )

    # Display the form only after clicking the button
    if not st.session_state.show_move_form:
        return

    wb = None

    try:
        wb = load_workbook(workbook_path)

        if "Inprogress" not in wb.sheetnames:
            st.error("The Inprogress sheet was not found.")
            return

        if "Completed Jobs" not in wb.sheetnames:
            st.error("The Completed Jobs sheet was not found.")
            return

        ws = wb["Inprogress"]
        completed_ws = wb["Completed Jobs"]

        def normalize(value):
            return (
                str(value or "")
                .strip()
                .casefold()
                .replace("_", " ")
            )

        # Find the header row dynamically
        header_row = None
        headers = []

        for row_num in range(1, min(ws.max_row, 20) + 1):
            row_values = [
                ws.cell(row_num, col).value
                for col in range(1, ws.max_column + 1)
            ]

            normalized = [normalize(value) for value in row_values]

            if "job name" in normalized and "status" in normalized:
                header_row = row_num
                headers = [
                    str(value or "").strip()
                    for value in row_values
                ]
                break

        if header_row is None:
            st.error(
                "Could not find Job Name and Status headers "
                "in the Inprogress sheet."
            )
            return

        source_columns = {
            normalize(header): index + 1
            for index, header in enumerate(headers)
            if normalize(header)
        }

        job_col = source_columns.get("job name")
        status_col = source_columns.get("status")

        if job_col is None or status_col is None:
            st.error("The Job Name or Status column is missing.")
            return

        # Find active Inprogress jobs
        available_jobs = []

        for row_num in range(header_row + 1, ws.max_row + 1):
            job_name = ws.cell(row_num, job_col).value
            status = ws.cell(row_num, status_col).value

            if (
                job_name is not None
                and str(status or "").strip().casefold() == "inprogress"
            ):
                available_jobs.append({
                    "row": row_num,
                    "name": str(job_name).strip()
                })

        if not available_jobs:
            st.info("No Inprogress jobs are available to complete.")
            return

        st.subheader("Move Job to Completed")

        selected_job = st.selectbox(
            "Select Inprogress Job",
            options=available_jobs,
            format_func=lambda job: job["name"],
            key="selected_job_to_complete"
        )

        # Collect completion details
        complexity = st.selectbox(
            "Complexity",
            ["Low", "Medium", "High"],
            key="completed_complexity"
        )

        total_pages = st.number_input(
            "Total Pages",
            min_value=0,
            step=1,
            value=0,
            key="completed_total_pages"
        )

        # Use existing operator and QC values when available
        operator_col = source_columns.get("operator")
        qc_col = source_columns.get("qc")

        existing_operator = ""
        existing_qc = ""

        if operator_col is not None:
            existing_operator = str(
                ws.cell(
                    selected_job["row"], operator_col
                ).value or ""
            ).strip()

        if qc_col is not None:
            existing_qc = str(
                ws.cell(
                    selected_job["row"], qc_col
                ).value or ""
            ).strip()

        operator = st.text_input(
            "Operator",
            value=existing_operator,
            key="completed_operator"
        )

        qc_name = st.text_input(
            "QC Name",
            value=existing_qc,
            key="completed_qc_name"
        )

        st.caption(
            "Delivered Time will be recorded automatically "
            "when you confirm completion."
        )

        if st.button(
            "✅ Confirm Completion",
            key="confirm_job_completion",
            use_container_width=True,
            type="primary"
        ):
            row_num = selected_job["row"]

            # Recheck status before moving the job
            current_status = str(
                ws.cell(row_num, status_col).value or ""
            ).strip().casefold()

            if current_status != "inprogress":
                st.error(
                    "This job is no longer Inprogress. "
                    "Please refresh and select another job."
                )
                return

            if not operator.strip():
                st.error("Please enter the Operator name.")
                return

            if not qc_name.strip():
                st.error("Please enter the QC name.")
                return

            # Find destination headers
            destination_header_row = None
            destination_headers = []

            for r in range(
                1, min(completed_ws.max_row, 20) + 1
            ):
                row_values = [
                    completed_ws.cell(r, c).value
                    for c in range(1, completed_ws.max_column + 1)
                ]

                normalized = [normalize(value) for value in row_values]

                if "job name" in normalized and "status" in normalized:
                    destination_header_row = r
                    destination_headers = [
                        str(value or "").strip()
                        for value in row_values
                    ]
                    break

            if destination_header_row is None:
                st.error(
                    "Could not find Job Name and Status headers "
                    "in the Completed Jobs sheet."
                )
                return

            destination_columns = {
                normalize(header): index + 1
                for index, header in enumerate(destination_headers)
                if normalize(header)
            }

            # Destination row: next available row after the header/data
            destination_row = max(
                destination_header_row + 1,
                completed_ws.max_row + 1
            )

            # Copy values and formatting using matching column headers
            source_values = {
                normalize(headers[c - 1]): ws.cell(
                    row_num, c
                ).value
                for c in range(1, len(headers) + 1)
                if normalize(headers[c - 1])
            }

            for dest_name, dest_col in destination_columns.items():
                if dest_name in source_values:
                    completed_ws.cell(
                        destination_row, dest_col
                    ).value = source_values[dest_name]

                    source_col = source_columns.get(dest_name)

                    if source_col is not None:
                        source_cell = ws.cell(row_num, source_col)
                        destination_cell = completed_ws.cell(
                            destination_row, dest_col
                        )

                        if source_cell.has_style:
                            destination_cell._style = copy(
                                source_cell._style
                            )

                        if source_cell.number_format:
                            destination_cell.number_format = (
                                source_cell.number_format
                            )

            # Add or update completion fields if the headers exist
            completion_values = {
                "status": "Delivered",
                "complexity": complexity,
                "total pages": int(total_pages),
                "operator": operator.strip(),
                "qc": qc_name.strip(),
                "delivered time": datetime.now()
            }

            for field_name, value in completion_values.items():
                dest_col = destination_columns.get(field_name)

                if dest_col is not None:
                    completed_ws.cell(
                        destination_row, dest_col
                    ).value = value

                    if field_name == "delivered time":
                        completed_ws.cell(
                            destination_row, dest_col
                        ).number_format = "dd-mm-yyyy hh:mm:ss AM/PM"

            # Copy row height when available
            source_height = ws.row_dimensions[row_num].height

            if source_height is not None:
                completed_ws.row_dimensions[
                    destination_row
                ].height = source_height

            # Save destination before removing source
            wb.save(workbook_path)

            clear_workflow_cache()
            # Remove the original job row after successful save
            ws.delete_rows(row_num, 1)

            wb.save(workbook_path)


            clear_workflow_cache()
            st.session_state.show_move_form = False

            st.success(
                f"Job '{selected_job['name']}' moved to "
                "Completed Jobs successfully."
            )

            st.rerun()

    except FileNotFoundError:
        st.error("workflow.xlsx was not found. Check the file path.")

    except PermissionError:
        st.error(
            "Cannot update workflow.xlsx. Close the workbook in Excel "
            "and try again."
        )

    except Exception as e:
        st.error(f"Unable to complete job: {e}")

    finally:
        if wb is not None:
            wb.close()


#----------------------------Move Job to Completed end----------------------------




def DeadlineMonitoring():
    import streamlit as st
    import openpyxl
    from datetime import datetime

    file_path = "workflow.xlsx"

    # ---------------------------------------------------------
    # Read Excel workbook
    # ---------------------------------------------------------
    try:
        wb = openpyxl.load_workbook(
            file_path,
            data_only=True,
            read_only=False
        )
    except Exception as e:
        st.error(f"Unable to open workflow.xlsx: {e}")
        return

    try:
        if "Inprogress" not in wb.sheetnames:
            st.error("Inprogress sheet not found.")
            return

        ws = wb["Inprogress"]

        # -----------------------------------------------------
        # Find header row and columns
        # -----------------------------------------------------
        header_row = None
        headers = {}

        for row_num in range(1, min(ws.max_row, 20) + 1):
            row_headers = {}

            for cell in ws[row_num]:
                if cell.value is not None:
                    row_headers[str(cell.value).strip()] = cell.column

            if (
                "Status" in row_headers
                and "Job Name" in row_headers
                and "Deadline IST" in row_headers
            ):
                header_row = row_num
                headers = row_headers
                break

        if header_row is None:
            st.error(
                "Required columns not found: "
                "Status, Job Name, Deadline IST."
            )
            return

        # -----------------------------------------------------
        # Current time
        # -----------------------------------------------------
        current_time = datetime.now()

        deadline_jobs = {
            "Normal": [],
            "Due Soon": [],
            "Critical": [],
            "Overdue": []
        }

        active_statuses = {"tbi", "inprogress", "qc"}

        # -----------------------------------------------------
        # Read active jobs
        # -----------------------------------------------------
        for row_num in range(header_row + 1, ws.max_row + 1):
            status = ws.cell(
                row=row_num,
                column=headers["Status"]
            ).value

            job_name = ws.cell(
                row=row_num,
                column=headers["Job Name"]
            ).value

            deadline = ws.cell(
                row=row_num,
                column=headers["Deadline IST"]
            ).value

            if not status or not job_name or not deadline:
                continue

            if str(status).strip().lower() not in active_statuses:
                continue

            # -------------------------------------------------
            # Convert deadline to datetime
            # -------------------------------------------------
            if isinstance(deadline, datetime):
                deadline_datetime = deadline
            else:
                deadline_datetime = None
                deadline_text = str(deadline).strip()

                for date_format in (
                    "%d-%m-%Y %I:%M:%S %p",
                    "%d-%m-%Y  %I:%M:%S %p",
                    "%d-%m-%Y %H:%M:%S",
                    "%d-%m-%Y %I:%M %p",
                    "%Y-%m-%d %H:%M:%S"
                ):
                    try:
                        deadline_datetime = datetime.strptime(
                            deadline_text,
                            date_format
                        )
                        break
                    except ValueError:
                        continue

                if deadline_datetime is None:
                    continue

            # -------------------------------------------------
            # Calculate remaining time
            # -------------------------------------------------
            total_seconds = (
                deadline_datetime - current_time
            ).total_seconds()

            if total_seconds < 0:
                category = "Overdue"
            elif total_seconds <= 3600:
                category = "Critical"
            elif total_seconds <= 14400:
                category = "Due Soon"
            else:
                category = "Normal"

            # -------------------------------------------------
            # Format remaining time
            # -------------------------------------------------
            seconds = abs(int(total_seconds))
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60

            if total_seconds < 0:
                time_text = f"{hours}h {minutes}m overdue"
            else:
                time_text = f"{hours}h {minutes}m remaining"

            # -------------------------------------------------
            # Operator
            # -------------------------------------------------
            operator = ""

            if "Operator" in headers:
                operator_value = ws.cell(
                    row=row_num,
                    column=headers["Operator"]
                ).value

                if operator_value:
                    operator = str(operator_value)

            deadline_jobs[category].append({
                "Job Name": str(job_name),
                "Status": str(status),
                "Operator": operator,
                "Deadline": deadline_datetime,
                "Time": time_text
            })

    finally:
        wb.close()

    # ---------------------------------------------------------
    # Heading
    # ---------------------------------------------------------
    st.markdown('<div class="section-heading"><span class="section-icon purple">◷</span>Deadline Monitoring</div>', unsafe_allow_html=True)

    # ---------------------------------------------------------
    # Category counts
    # ---------------------------------------------------------
    normal_count = len(deadline_jobs["Normal"])
    due_soon_count = len(deadline_jobs["Due Soon"])
    critical_count = len(deadline_jobs["Critical"])
    overdue_count = len(deadline_jobs["Overdue"])

    # ---------------------------------------------------------
    # ---------------------------------------------------------
    # ---------------------------------------------------------
    # Deadline Monitoring cards: taller cards and consistent count typography
    # ---------------------------------------------------------
    st.markdown("""
    <style>
    /* Deadline Monitoring cards: match Job Status Summary label/count proportions */
    .st-key-deadline_normal button,
    .st-key-deadline_due_soon button,
    .st-key-deadline_critical button,
    .st-key-deadline_overdue button {
        width: 100% !important;
        min-height: 126px !important;
        height: 126px !important;
        padding: 18px !important;
        box-sizing: border-box !important;
        border-radius: 18px !important;
        font-size: 1rem !important;
        font-weight: 500 !important;
        line-height: 1.35 !important;
        white-space: pre-wrap !important;
        transition: all 0.18s ease-in-out !important;
    }

    /* Explicit label typography */
    .st-key-deadline_normal button div[data-testid="stMarkdownContainer"],
    .st-key-deadline_due_soon button div[data-testid="stMarkdownContainer"],
    .st-key-deadline_critical button div[data-testid="stMarkdownContainer"],
    .st-key-deadline_overdue button div[data-testid="stMarkdownContainer"] {
        font-size: 1rem !important;
        font-weight: 500 !important;
        line-height: 1.35 !important;
    }

    .st-key-deadline_normal button p,
    .st-key-deadline_due_soon button p,
    .st-key-deadline_critical button p,
    .st-key-deadline_overdue button p {
        margin: 0 !important;
        font-size: 1rem !important;
        font-weight: 500 !important;
        line-height: 1.35 !important;
    }

    /* Number styling uses the same large bold scale as Job Status Summary */
    .st-key-deadline_normal button p strong,
    .st-key-deadline_due_soon button p strong,
    .st-key-deadline_critical button p strong,
    .st-key-deadline_overdue button p strong {
        font-size: 2.35rem !important;
        font-weight: 780 !important;
        line-height: 1.12 !important;
        letter-spacing: -0.025em !important;
    }

    /* Color variants */
    .st-key-deadline_normal button {
        background: linear-gradient(135deg, #EAF8EF, #D9F3E2) !important;
        color: #166534 !important;
        border: 1px solid #86D6A0 !important;
        border-top: 4px solid #16A34A !important;
        box-shadow: 0 4px 12px rgba(22,163,74,.10) !important;
    }
    .st-key-deadline_normal button:hover { background: #D1FAE0 !important; }

    .st-key-deadline_due_soon button {
        background: linear-gradient(135deg, #FFF7E6, #FFEDC2) !important;
        color: #92400E !important;
        border: 1px solid #F5CD78 !important;
        border-top: 4px solid #F59E0B !important;
        box-shadow: 0 4px 12px rgba(245,158,11,.10) !important;
    }
    .st-key-deadline_due_soon button:hover { background: #FDE7A6 !important; }

    .st-key-deadline_critical button {
        background: linear-gradient(135deg, #FFF0E5, #FFDDC7) !important;
        color: #9A3412 !important;
        border: 1px solid #FDBA8C !important;
        border-top: 4px solid #F97316 !important;
        box-shadow: 0 4px 12px rgba(249,115,22,.12) !important;
    }
    .st-key-deadline_critical button:hover { background: #FED7AA !important; }

    .st-key-deadline_overdue button {
        background: linear-gradient(135deg, #FFF0F0, #FFDCDC) !important;
        color: #991B1B !important;
        border: 1px solid #F5A0A0 !important;
        border-top: 4px solid #DC2626 !important;
        box-shadow: 0 4px 12px rgba(220,38,38,.10) !important;
    }
    .st-key-deadline_overdue button:hover { background: #FECACA !important; }
    </style>
    """, unsafe_allow_html=True)

    # ---------------------------------------------------------
    # Deadline category buttons
    # ---------------------------------------------------------
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button(
            f"🟢 Normal\n**{normal_count}**",
            key="deadline_normal",
            type="primary",
            use_container_width=True
        ):
            st.session_state.deadline_category = "Normal"

    with col2:
        if st.button(
            f"🟡 Due Soon\n**{due_soon_count}**",
            key="deadline_due_soon",
            type="primary",
            use_container_width=True
        ):
            st.session_state.deadline_category = "Due Soon"

    with col3:
        if st.button(
            f"🟠 Critical\n**{critical_count}**",
            key="deadline_critical",
            type="primary",
            use_container_width=True
        ):
            st.session_state.deadline_category = "Critical"

    with col4:
        if st.button(
            f"🚨 Overdue\n**{overdue_count}**",
            key="deadline_overdue",
            type="primary",
            use_container_width=True
        ):
            st.session_state.deadline_category = "Overdue"

    # ---------------------------------------------------------
    # Show selected category
    # ---------------------------------------------------------
    selected_category = st.session_state.get(
        "deadline_category"
    )

    if selected_category in deadline_jobs:

        st.divider()
        st.subheader(f"{selected_category} Jobs")

        selected_jobs = deadline_jobs[selected_category]

        if not selected_jobs:
            st.info(f"No {selected_category} jobs found.")
        else:
            for job in selected_jobs:
                st.write(f"**Job Name:** {job['Job Name']}")
                st.write(f"**Status:** {job['Status']}")

                if job["Operator"]:
                    st.write(f"**Operator:** {job['Operator']}")

                st.write(
                    "**Deadline IST:** "
                    + job["Deadline"].strftime(
                        "%d-%m-%Y %I:%M:%S %p"
                    )
                )

                st.write(f"**Time:** {job['Time']}")
                st.divider()

#----------------------------Deadline Monitoring end----------------------------



def JobSearch():
    st.subheader("🔎 Job Search")

    search_text = st.text_input(
        "Search by Job Name, Project Number, WO Number, or Requestor/Client Name",
        placeholder="Enter job name or number...",
        key="job_search_text"
    )

    status_filter = st.selectbox(
        "Filter by Status",
        ["All", "TBI", "Inprogress", "QC", "Delivered"],
        key="job_search_status"
    )

    # Do not display any message when the search field is empty
    if not search_text.strip():
        return

    workbook_path = "workflow.xlsx"

    try:
        wb = load_workbook(
            workbook_path,
            read_only=False,
            data_only=True
        )

        results = []

        sheets_to_search = ["Inprogress", "Completed Jobs"]

        search_columns = [
            "job name",
            "project number",
            "wo number",
            "requestor",
            "requestor/client name",
            "client name"
        ]

        search_term = search_text.strip().casefold()

        for sheet_name in sheets_to_search:
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]
            header_row = None
            headers = []

            # Find the header row
            for row_num in range(1, min(ws.max_row, 20) + 1):
                row_values = [
                    str(ws.cell(row_num, col).value or "").strip()
                    for col in range(1, ws.max_column + 1)
                ]

                normalized = [h.casefold() for h in row_values]

                if "job name" in normalized and "status" in normalized:
                    header_row = row_num
                    headers = row_values
                    break

            if header_row is None:
                continue

            normalized_headers = [h.casefold() for h in headers]

            searchable_indexes = [
                i for i, header in enumerate(normalized_headers)
                if (
                    header in search_columns
                    or "requestor" in header
                    or "client name" in header
                    or "project number" in header
                    or "wo number" in header
                    or "job name" in header
                )
            ]

            status_index = normalized_headers.index("status")

            for row_num in range(header_row + 1, ws.max_row + 1):
                values = [
                    ws.cell(row_num, col).value
                    for col in range(1, len(headers) + 1)
                ]

                if not any(value is not None for value in values):
                    continue

                row_status = str(
                    values[status_index] or ""
                ).strip()

                if (
                    status_filter != "All"
                    and row_status.casefold() != status_filter.casefold()
                ):
                    continue

                matched = any(
                    search_term in str(values[i] or "").casefold()
                    for i in searchable_indexes
                )

                if matched:
                    record = {
                        headers[i] or f"Column {i + 1}": values[i]
                        for i in range(len(headers))
                    }

                    record["Source Sheet"] = sheet_name
                    results.append(record)

        wb.close()

        if results:
            st.success(f"Found {len(results)} matching job(s).")

            df = pd.DataFrame(results)

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("No matching jobs found.")

    except FileNotFoundError:
        st.error("workflow.xlsx was not found. Check the file path.")

    except Exception as e:
        st.error(f"Unable to search jobs: {e}")

#----------------------------Job Search end----------------------------




def EmployeeWorkloadContent():
    st.subheader("👥 Employee Workload Dashboard")

    workbook_path = "workflow.xlsx"
    sheets_to_read = ["Inprogress", "Completed Jobs"]

    def normalize_header(value):
        return str(value or "").strip().casefold().replace("_", " ")

    def get_column_index(headers, possible_names):
        normalized_headers = [
            normalize_header(header) for header in headers
        ]

        for name in possible_names:
            normalized_name = normalize_header(name)
            if normalized_name in normalized_headers:
                return normalized_headers.index(normalized_name)

        return None

    try:
        wb = load_workbook(
            workbook_path,
            read_only=False,
            data_only=True
        )

        all_jobs = []
        unassigned_active_jobs = 0
        current_time = datetime.now()

        for sheet_name in sheets_to_read:
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]
            header_row = None
            headers = []

            # Locate the header row
            for row_num in range(1, min(ws.max_row, 20) + 1):
                row_values = [
                    ws.cell(row_num, col).value
                    for col in range(1, ws.max_column + 1)
                ]

                normalized = [
                    normalize_header(value) for value in row_values
                ]

                if "job name" in normalized and "status" in normalized:
                    header_row = row_num
                    headers = [
                        str(value or "").strip()
                        for value in row_values
                    ]
                    break

            if header_row is None:
                continue

            status_idx = get_column_index(headers, ["Status"])
            job_idx = get_column_index(
                headers, ["Job Name", "Job_Name"]
            )
            operator_idx = get_column_index(
                headers,
                ["Operator", "Assigned Operator", "Employee", "Employee Name"]
            )
            deadline_idx = get_column_index(
                headers, ["Deadline IST", "Deadline"]
            )

            if status_idx is None:
                continue

            for row_num in range(header_row + 1, ws.max_row + 1):
                values = [
                    ws.cell(row_num, col).value
                    for col in range(1, len(headers) + 1)
                ]

                if not any(value is not None for value in values):
                    continue

                status = str(
                    values[status_idx] or ""
                ).strip().casefold()

                job_name = (
                    str(values[job_idx] or "").strip()
                    if job_idx is not None else ""
                )

                operator = (
                    str(values[operator_idx] or "").strip()
                    if operator_idx is not None else ""
                )

                deadline_value = (
                    values[deadline_idx]
                    if deadline_idx is not None else None
                )

                # Only active jobs belong to the Inprogress sheet
                is_active_sheet = sheet_name == "Inprogress"

                if is_active_sheet and status in {
                    "tbi", "inprogress", "qc"
                }:
                    if not operator or operator.casefold() in {
                        "not assigned", "unassigned", "none", "nan"
                    }:
                        unassigned_active_jobs += 1

                # Ignore empty job rows
                if not job_name and not operator:
                    continue

                # Determine whether the active job is overdue
                is_overdue = False

                if (
                    is_active_sheet
                    and status in {"tbi", "inprogress", "qc"}
                    and deadline_value is not None
                ):
                    try:
                        deadline = pd.to_datetime(
                            deadline_value, errors="coerce"
                        )

                        if not pd.isna(deadline):
                            deadline_dt = deadline.to_pydatetime()

                            # Compare without timezone information
                            if deadline_dt.tzinfo is not None:
                                deadline_dt = deadline_dt.replace(
                                    tzinfo=None
                                )

                            is_overdue = deadline_dt < current_time

                    except (ValueError, TypeError, OverflowError):
                        is_overdue = False

                all_jobs.append({
                    "Job Name": job_name,
                    "Operator": operator,
                    "Status": status,
                    "Source Sheet": sheet_name,
                    "Overdue": is_overdue
                })

        wb.close()

        if not all_jobs:
            st.info("No job records were found in the selected sheets.")
            return

        jobs_df = pd.DataFrame(all_jobs)

        # Exclude unassigned placeholders from operator summaries
        invalid_operators = {
            "", "not assigned", "unassigned", "none", "nan"
        }

        assigned_df = jobs_df[
            jobs_df["Operator"].fillna("").str.strip().str.casefold()
            .map(lambda name: name not in invalid_operators)
        ].copy()

        if assigned_df.empty:
            st.warning(
                "No jobs have been assigned to an operator yet."
            )
            st.metric(
                "Unassigned Active Jobs",
                unassigned_active_jobs
            )
            return

        # Preserve the original display name but group case-insensitively
        assigned_df["_operator_key"] = (
            assigned_df["Operator"].str.strip().str.casefold()
        )

        operator_names = (
            assigned_df.groupby("_operator_key")["Operator"]
            .first()
            .to_dict()
        )

        # Summary counts for each operator
        summary_rows = []

        for operator_key, operator_name in operator_names.items():
            operator_jobs = assigned_df[
                assigned_df["_operator_key"] == operator_key
            ]

            pending_count = int(
                operator_jobs["Status"].isin(
                    ["tbi", "inprogress"]
                ).sum()
            )

            qc_count = int(
                (operator_jobs["Status"] == "qc").sum()
            )

            completed_count = int(
                (
                    (operator_jobs["Source Sheet"] == "Completed Jobs")
                    & operator_jobs["Status"].isin(
                        ["delivered", "completed"]
                    )
                ).sum()
            )

            overdue_count = int(
                (
                    (operator_jobs["Source Sheet"] == "Inprogress")
                    & operator_jobs["Overdue"]
                ).sum()
            )

            summary_rows.append({
                "Operator": operator_name,
                "Assigned Jobs": len(operator_jobs),
                "Pending Jobs": pending_count,
                "QC Jobs": qc_count,
                "Completed Jobs": completed_count,
                "Overdue Jobs": overdue_count
            })

        summary_df = pd.DataFrame(summary_rows)
        summary_df = summary_df.sort_values(
            by="Assigned Jobs",
            ascending=False
        ).reset_index(drop=True)

        # Dashboard totals
        total_assigned = int(summary_df["Assigned Jobs"].sum())
        total_pending = int(summary_df["Pending Jobs"].sum())
        total_qc = int(summary_df["QC Jobs"].sum())
        total_completed = int(summary_df["Completed Jobs"].sum())
        total_overdue = int(summary_df["Overdue Jobs"].sum())

        st.markdown('<div class="section-heading"><span class="section-icon blue">▥</span>Overall Workload</div>', unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        col1.metric("Assigned Jobs", total_assigned)
        col2.metric("Pending Jobs", total_pending)
        col3.metric("QC Jobs", total_qc)

        col4, col5, col6 = st.columns(3)
        col4.metric("Completed Jobs", total_completed)
        col5.metric("Overdue Jobs", total_overdue)
        col6.metric("Unassigned Active Jobs", unassigned_active_jobs)

        st.divider()
        st.markdown("### Operator-wise Workload")

        # Search and filter operators
        operator_options = ["All Operators"] + sorted(
            summary_df["Operator"].unique().tolist(),
            key=str.casefold
        )

        selected_operator = st.selectbox(
            "Select Operator",
            operator_options,
            key="workload_operator_filter"
        )

        display_df = summary_df.copy()

        if selected_operator != "All Operators":
            display_df = display_df[
                display_df["Operator"] == selected_operator
            ]

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True
        )

        # Workload chart
        st.markdown("### Assigned Jobs by Operator")

        chart_df = display_df.set_index("Operator")[
            ["Pending Jobs", "QC Jobs", "Completed Jobs"]
        ]

        st.bar_chart(chart_df)

        # Export workload report without modifying workflow.xlsx
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            display_df.to_excel(
                writer,
                index=False,
                sheet_name="Employee Workload"
            )

        output.seek(0)

        st.download_button(
            label="📥 Download Employee Workload Report",
            data=output.getvalue(),
            file_name=(
                "Employee_Workload_"
                + datetime.now().strftime("%Y%m%d_%H%M%S")
                + ".xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_employee_workload"
        )

    except FileNotFoundError:
        st.error(
            "workflow.xlsx was not found. Check that it is in "
            "the correct folder."
        )

    except Exception as e:
        st.error(f"Unable to load Employee Workload Dashboard: {e}")


def EmployeeWorkloadDashboard():
    if "show_employee_workload" not in st.session_state:
        st.session_state.show_employee_workload = False

    if st.button(
        "👥 Employee Workload Dashboard",
        key="open_employee_workload",
        use_container_width=True,
        type="primary"
    ):
        st.session_state.show_employee_workload = (
            not st.session_state.show_employee_workload
        )

    if st.session_state.show_employee_workload:
        EmployeeWorkloadContent()

#--------------------------------------------- Employee Workload end-------------------------




def ReportsExportContent():
    st.subheader("📊 Reports & Excel Export")

    workbook_path = "workflow.xlsx"
    sheets_to_read = ["Inprogress", "Completed Jobs"]

    try:
        wb = load_workbook(
            workbook_path,
            read_only=False,
            data_only=True
        )

        all_records = []

        for sheet_name in sheets_to_read:
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]
            header_row = None
            headers = []

            # Find the header row
            for row_num in range(1, min(ws.max_row, 20) + 1):
                row_values = [
                    str(ws.cell(row_num, col).value or "").strip()
                    for col in range(1, ws.max_column + 1)
                ]

                normalized = [h.casefold() for h in row_values]

                if "job name" in normalized and "status" in normalized:
                    header_row = row_num
                    headers = row_values
                    break

            if header_row is None:
                continue

            # Read job records
            for row_num in range(header_row + 1, ws.max_row + 1):
                values = [
                    ws.cell(row_num, col).value
                    for col in range(1, len(headers) + 1)
                ]

                if not any(value is not None for value in values):
                    continue

                record = {
                    headers[i] or f"Column {i + 1}": values[i]
                    for i in range(len(headers))
                }

                record["Source Sheet"] = sheet_name
                all_records.append(record)

        wb.close()

        if not all_records:
            st.warning("No job records found in the workbook.")
            return

        df = pd.DataFrame(all_records)

        column_lookup = {
            str(column).strip().casefold(): column
            for column in df.columns
        }

        status_col = column_lookup.get("status")

        if status_col is None:
            st.error("The Status column was not found.")
            return

        df["_status_normalized"] = (
            df[status_col]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
        )

        # Status counts
        status_counts = {
            "TBI": int((df["_status_normalized"] == "tbi").sum()),
            "Inprogress": int(
                (df["_status_normalized"] == "inprogress").sum()
            ),
            "QC": int((df["_status_normalized"] == "qc").sum()),
            "Delivered": int(
                df["_status_normalized"]
                .isin(["delivered", "completed"]).sum()
            )
        }

        st.markdown("### Job Status Summary")

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("TBI", status_counts["TBI"])
        col2.metric("Inprogress", status_counts["Inprogress"])
        col3.metric("QC", status_counts["QC"])
        col4.metric("Delivered", status_counts["Delivered"])

        st.divider()
        st.markdown("### Filter Report")

        selected_status = st.selectbox(
            "Select Job Status",
            ["All", "TBI", "Inprogress", "QC", "Delivered"],
            key="report_status_filter"
        )

        # Identify available date columns
        preferred_dates = [
            "date",
            "receipt time",
            "initiation time",
            "deadline ist",
            "deadline",
            "delivered time"
        ]

        date_columns = [
            column_lookup[name]
            for name in preferred_dates
            if name in column_lookup
        ]

        selected_date_col = None
        use_date_filter = False
        start_date = None
        end_date = None

        if date_columns:
            preferred_default = [
                "date",
                "receipt time",
                "initiation time",
                "delivered time",
                "deadline ist",
                "deadline"
            ]

            default_col = next(
                (
                    column_lookup[name]
                    for name in preferred_default
                    if name in column_lookup
                ),
                date_columns[0]
            )

            selected_date_col = st.selectbox(
                "Filter Date By",
                date_columns,
                index=date_columns.index(default_col),
                key="report_date_column"
            )

            parsed_dates = pd.to_datetime(
                df[selected_date_col],
                errors="coerce"
            ).dropna()

            if not parsed_dates.empty:
                min_date = parsed_dates.min().date()
                max_date = parsed_dates.max().date()

                date_range = st.date_input(
                    "Select date range",
                    value=(min_date, max_date),
                    min_value=min_date,
                    max_value=max_date,
                    key="report_date_range"
                )

                if (
                    isinstance(date_range, (tuple, list))
                    and len(date_range) == 2
                ):
                    start_date, end_date = date_range
                    use_date_filter = True
                else:
                    st.caption(
                        "Select both a start date and an end date."
                    )
            else:
                st.info(
                    f"No valid dates found in {selected_date_col}."
                )
        else:
            st.info("No recognized date columns were found.")

        # Apply status filter
        filtered_df = df.copy()

        if selected_status == "Delivered":
            filtered_df = filtered_df[
                filtered_df["_status_normalized"].isin(
                    ["delivered", "completed"]
                )
            ]
        elif selected_status != "All":
            filtered_df = filtered_df[
                filtered_df["_status_normalized"]
                == selected_status.casefold()
            ]

        # Apply date filter
        if use_date_filter:
            parsed = pd.to_datetime(
                filtered_df[selected_date_col],
                errors="coerce"
            )

            date_only = parsed.dt.date

            filtered_df = filtered_df[
                date_only.notna()
                & (date_only >= start_date)
                & (date_only <= end_date)
            ]

        report_df = filtered_df.drop(
            columns=["_status_normalized"],
            errors="ignore"
        )

        st.divider()
        st.markdown("### Report Results")
        st.write(f"**Total matching jobs: {len(report_df)}**")

        if report_df.empty:
            st.info("No jobs match the selected filters.")
            return

        st.dataframe(
            report_df,
            use_container_width=True,
            hide_index=True
        )

        # Generate a separate Excel report
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report_df.to_excel(
                writer,
                index=False,
                sheet_name="Job Report"
            )

            summary_df = pd.DataFrame(
                list(status_counts.items()),
                columns=["Status", "Job Count"]
            )

            summary_df.to_excel(
                writer,
                index=False,
                sheet_name="Status Summary"
            )

        output.seek(0)

        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name=(
                "Job_Report_"
                + datetime.now().strftime("%Y%m%d_%H%M%S")
                + ".xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="download_job_report"
        )

    except FileNotFoundError:
        st.error("workflow.xlsx was not found. Check the file path.")

    except Exception as e:
        st.error(f"Unable to generate report: {e}")


def ReportsExport():
    if "show_reports_export" not in st.session_state:
        st.session_state.show_reports_export = False

    if st.button(
        "📊 Reports & Excel Export",
        key="open_reports_export",
        use_container_width=True,
        type="primary"
    ):
        st.session_state.show_reports_export = (
            not st.session_state.show_reports_export
        )

    if st.session_state.show_reports_export:
        ReportsExportContent()

#-------------------------------------Reports & Excel Export button end------------------------------------





# -----------------------------------------------------
# Chatbot 
# -----------------------------------------------------
def Inprogress():

    file_path = "workflow.xlsx"
    if not os.path.exists(file_path):
        st.error("workflow.xlsx was not found.")
        return
    df = cached_excel_dataframe(file_path, os.path.getmtime(file_path))

    inprogress_jobs = df[
        df["Status"].astype(str).str.strip().str.lower() == "inprogress"
    ]

    st.write("Inprogress Job Name:")

    for job in inprogress_jobs["Job Name"].dropna().to_list():
        st.write(job)

    inprogress_count = len(inprogress_jobs)

    st.write("Inprogress Job Count:", inprogress_count)



def TBI():

    file_path = "workflow.xlsx"
    if not os.path.exists(file_path):
        st.error("workflow.xlsx was not found.")
        return
    df = cached_excel_dataframe(file_path, os.path.getmtime(file_path))

    tbi_jobs = df[
        df["Status"].astype(str).str.strip().str.lower() == "tbi"
    ]

    st.write("TBI Job Name:")

    for job in tbi_jobs["Job Name"].dropna().to_list():
        st.write(job)

    tbi_count = len(tbi_jobs)

    st.write("TBI Job Count:", tbi_count)



def QC():

    file_path = "workflow.xlsx"
    if not os.path.exists(file_path):
        st.error("workflow.xlsx was not found.")
        return
    df = cached_excel_dataframe(file_path, os.path.getmtime(file_path))

    qc_jobs = df[
        df["Status"].astype(str).str.strip().str.lower() == "qc"
    ]

    st.write("QC Job Name:")

    for job in qc_jobs["Job Name"].dropna().to_list():
        st.write(job)

    qc_count = len(qc_jobs)

    st.write("QC Job Count:", qc_count)

# -----------------------------------------------------chatbot end---------------------




# -----------------------------------------------------
# AI MODE: cloud LLM + validated local workflow actions
# -----------------------------------------------------
def _ai_get_api_key():
    """Read AICredits API key from Streamlit secrets first, then environment."""
    try:
        key = st.secrets.get("AICREDITS_API_KEY", "")
    except Exception:
        key = ""
    key = key or os.environ.get("AICREDITS_API_KEY", "")
    return str(key).strip()


def _ai_get_model():
    """Return the model name configured for the AICredits OpenAI-compatible API."""
    try:
        model = st.secrets.get("AICREDITS_MODEL", "claude-haiku-4-5")
    except Exception:
        model = "claude-haiku-4-5"
    return os.environ.get("AICREDITS_MODEL", model) or "claude-haiku-4-5"


def _ai_get_client():
    """Create the LangChain client for AICredits' OpenAI-compatible endpoint."""
    api_key = _ai_get_api_key()
    if not api_key:
        raise RuntimeError(
            "AICredits API key is not configured. Add AICREDITS_API_KEY to "
            ".streamlit/secrets.toml. Do not paste your key into the source code."
        )
    return ChatOpenAI(
        model=_ai_get_model(),
        api_key=api_key,
        base_url="https://api.aicredits.in/v1",
        temperature=0,
    )


def _ai_response_text(response):
    """Convert LangChain response content to a plain string safely."""
    content = getattr(response, "content", response)
    if isinstance(content, str):
        return content.strip()
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, str):
                parts.append(item)
            elif isinstance(item, dict) and item.get("text"):
                parts.append(str(item["text"]))
        return "\n".join(parts).strip()
    return str(content).strip()

def _normalize_header(value):
    return " ".join(str(value or "").strip().casefold().replace("_", " ").split())


def _find_sheet_headers(ws, max_scan=20):
    """Find a row containing Job Name and Status; return (row, normalized header map)."""
    for row_num in range(1, min(ws.max_row, max_scan) + 1):
        headers = {}
        for col_num in range(1, ws.max_column + 1):
            value = ws.cell(row_num, col_num).value
            if value is not None and str(value).strip():
                headers[_normalize_header(value)] = col_num
        if "job name" in headers and "status" in headers:
            return row_num, headers
    return None, {}


def _read_workflow_snapshot():
    """Read a compact view of workbook data for reporting; never sends the whole workbook."""
    path = "workflow.xlsx"
    if not os.path.exists(path):
        raise FileNotFoundError("workflow.xlsx was not found in the application folder.")
    wb = load_workbook(path, data_only=True, read_only=False)
    snapshot = {"sheets": {}, "counts": {}, "jobs": []}
    try:
        for sheet_name in ("Inprogress", "Completed Jobs"):
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            header_row, headers = _find_sheet_headers(ws)
            if header_row is None:
                continue
            rows = []
            for row_num in range(header_row + 1, ws.max_row + 1):
                job_name_col = headers.get("job name")
                status_col = headers.get("status")
                job_name = ws.cell(row_num, job_name_col).value if job_name_col else None
                status = ws.cell(row_num, status_col).value if status_col else None
                if not job_name:
                    continue
                item = {"job_name": str(job_name).strip(), "status": str(status or "").strip(), "sheet": sheet_name}
                for field in ("operator", "qc", "deadline ist", "receipt time", "complexity", "total pages", "work order", "project number", "client name", "requestor"):
                    col = headers.get(field)
                    if col:
                        value = ws.cell(row_num, col).value
                        if isinstance(value, (datetime, date)):
                            value = value.isoformat(sep=" ") if isinstance(value, datetime) else value.isoformat()
                        if value is not None:
                            item[field.replace(" ", "_")] = str(value)
                rows.append(item)
            snapshot["sheets"][sheet_name] = rows
            for item in rows:
                snapshot["jobs"].append(item)
        status_counts = {}
        for item in snapshot["jobs"]:
            status = item.get("status", "Unknown") or "Unknown"
            status_counts[status] = status_counts.get(status, 0) + 1
        snapshot["counts"] = status_counts
        return snapshot
    finally:
        wb.close()


def _ai_classify_request(user_text, snapshot):
    """Classify intent with the AICredits-hosted Claude model; actions remain locally validated."""
    client = _ai_get_client()
    schema_description = {
        "intent": "one of query, report, assign_job, move_to_qc, complete_job, unknown",
        "job_name": "string or null",
        "operator": "string or null",
        "qc_name": "string or null",
        "complexity": "string or null",
        "total_pages": "integer or null",
        "question": "string"
    }
    prompt = (
        "You are the intent parser for an office Workflow Automation System. "
        "Classify the user's request as query, report, assign_job, move_to_qc, complete_job, or unknown. "
        "Never claim an action was performed. Actions are only proposals and require explicit UI confirmation. "
        "Extract job_name, operator, qc_name, complexity and total_pages only when explicitly supplied. "
        "Do not invent names or values. For counts, lists, status, deadlines or workload use query/report. "
        "Return ONLY a valid JSON object matching these fields: " + json.dumps(schema_description) + "\n"
        "Workbook status summary (context only): " + json.dumps({"counts": snapshot.get("counts", {})}, ensure_ascii=False) + "\n"
        "User request: " + user_text
    )
    response = client.invoke(prompt)
    response_text = _ai_response_text(response)
    # Some compatible APIs wrap JSON in Markdown fences; strip them before parsing.
    response_text = re.sub(r"^```(?:json)?\s*|\s*```$", "", response_text, flags=re.IGNORECASE).strip()
    try:
        result = json.loads(response_text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", response_text, flags=re.DOTALL)
        if not match:
            raise ValueError("The AI service did not return valid JSON for request classification. Please rephrase the request.")
        result = json.loads(match.group(0))
    allowed = {"query", "report", "assign_job", "move_to_qc", "complete_job", "unknown"}
    if result.get("intent") not in allowed:
        result["intent"] = "unknown"
    for field in ("job_name", "operator", "qc_name", "complexity", "total_pages", "question"):
        result.setdefault(field, None if field != "question" else user_text)
    return result


def _ai_answer_from_snapshot(question, snapshot):
    """Use Python as source of truth for common reports; send only compact derived data to the LLM."""
    q = question.casefold()
    jobs = snapshot.get("jobs", [])
    counts = snapshot.get("counts", {})
    if any(word in q for word in ("count", "how many", "total jobs", "number of jobs")):
        if "inprogress" in q or "in progress" in q:
            selected = [j for j in jobs if j.get("status", "").casefold() == "inprogress"]
            return f"Inprogress Job Count: {len(selected)}\n\nJob names:\n" + ("\n".join(f"- {j['job_name']}" for j in selected) or "No Inprogress jobs found.")
        if "tbi" in q:
            selected = [j for j in jobs if j.get("status", "").casefold() == "tbi"]
            return f"TBI Job Count: {len(selected)}\n\nJob names:\n" + ("\n".join(f"- {j['job_name']}" for j in selected) or "No TBI jobs found.")
        if "qc" in q:
            selected = [j for j in jobs if j.get("status", "").casefold() == "qc"]
            return f"QC Job Count: {len(selected)}\n\nJob names:\n" + ("\n".join(f"- {j['job_name']}" for j in selected) or "No QC jobs found.")
        return "Job counts by status:\n" + ("\n".join(f"- {k}: {v}" for k, v in sorted(counts.items())) or "No jobs found.")
    if "overdue" in q or "deadline" in q or "due soon" in q or "critical" in q:
        now = datetime.now()
        selected = []
        for job in jobs:
            deadline = job.get("deadline_ist")
            if not deadline or job.get("status", "").casefold() not in {"tbi", "inprogress", "qc"}:
                continue
            try:
                dt = datetime.fromisoformat(deadline)
            except ValueError:
                for fmt in ("%d-%m-%Y %I:%M:%S %p", "%d-%m-%Y  %I:%M:%S %p", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %I:%M %p"):
                    try:
                        dt = datetime.strptime(deadline, fmt); break
                    except ValueError:
                        dt = None
                if dt is None:
                    continue
            seconds = (dt - now).total_seconds()
            category = "Overdue" if seconds < 0 else "Critical" if seconds <= 3600 else "Due Soon" if seconds <= 14400 else "Normal"
            if "overdue" in q and category == "Overdue" or "critical" in q and category == "Critical" or "due soon" in q and category == "Due Soon" or "deadline" in q:
                selected.append((job, category, dt))
        if not selected:
            return "No matching active jobs with readable deadlines were found in workflow.xlsx."
        return "Deadline report:\n\n" + "\n".join(f"- {j['job_name']} | {j.get('status')} | {cat} | Deadline: {dt.strftime('%d-%m-%Y %I:%M %p')}" for j, cat, dt in selected)
    if "operator" in q or "workload" in q or "assigned" in q:
        by_operator = {}
        for job in jobs:
            operator = job.get("operator", "Unassigned") or "Unassigned"
            by_operator[operator] = by_operator.get(operator, 0) + 1
        return "Job count by operator:\n" + "\n".join(f"- {name}: {count}" for name, count in sorted(by_operator.items()))
    # Search first, using Python data rather than trusting model-generated records.
    matched = [j for j in jobs if any(token in j.get("job_name", "").casefold() for token in q.split() if len(token) > 2)]
    if matched:
        return "Matching jobs from workflow.xlsx:\n\n" + "\n".join(
            "- " + j["job_name"] + " | Status: " + (j.get("status") or "Unknown") +
            (" | Operator: " + j["operator"] if j.get("operator") else "") +
            (" | QC: " + j["qc"] if j.get("qc") else "") +
            (" | Deadline: " + j["deadline_ist"] if j.get("deadline_ist") else "")
            for j in matched[:30]
        )
    # For open-ended reporting, pass only summarized rows, not raw workbook bytes.
    compact = [{k: v for k, v in j.items() if k in {"job_name", "status", "operator", "qc", "deadline_ist", "complexity", "total_pages"}} for j in jobs[:250]]
    client = _ai_get_client()
    response = client.invoke(
        "Answer questions about the provided workflow snapshot. Treat the data as authoritative; "
        "do not invent jobs, counts, deadlines or names. State when the snapshot does not contain enough information. "
        "Do not execute actions.\nQuestion: " + question + "\nStatus counts: " + json.dumps(counts) +
        "\nWorkflow rows (limited snapshot): " + json.dumps(compact, ensure_ascii=False)
    )
    return _ai_response_text(response) or "The AI service returned an empty response. Please try again."


def _perform_confirmed_workflow_action(action):
    """Validate and apply a confirmed action using OpenPyXL, retaining existing workbook and cell styles."""
    path = "workflow.xlsx"
    if not os.path.exists(path):
        raise FileNotFoundError("workflow.xlsx was not found.")
    wb = load_workbook(path)
    try:
        if "Inprogress" not in wb.sheetnames:
            raise ValueError("The Inprogress worksheet was not found.")
        ws = wb["Inprogress"]
        header_row, src_headers = _find_sheet_headers(ws)
        if header_row is None:
            raise ValueError("Could not find Job Name and Status headers in Inprogress.")
        job_name = str(action.get("job_name") or "").strip()
        if not job_name:
            raise ValueError("A job name is required.")
        name_col, status_col = src_headers["job name"], src_headers["status"]
        found_row = None
        for r in range(header_row + 1, ws.max_row + 1):
            if str(ws.cell(r, name_col).value or "").strip().casefold() == job_name.casefold():
                found_row = r
                break
        if found_row is None:
            raise ValueError(f"Job '{job_name}' was not found in the Inprogress sheet. Refresh and try again.")
        current_status = str(ws.cell(found_row, status_col).value or "").strip().casefold()
        action_type = action.get("intent")
        if action_type == "assign_job":
            operator_col = src_headers.get("operator")
            receipt_col = src_headers.get("receipt time")
            initiation_col = src_headers.get("initiation time")
            deadline_col = src_headers.get("deadline ist")
            if not operator_col:
                raise ValueError("The Operator column was not found in the Inprogress sheet.")
            operator = str(action.get("operator") or "").strip()
            if not operator:
                raise ValueError("Please provide an operator name before confirming assignment.")
            if current_status not in {"tbi", "inprogress"}:
                raise ValueError(f"This job has status '{current_status}'. Only TBI/Inprogress jobs can be assigned.")
            ws.cell(found_row, operator_col).value = operator
            now = datetime.now()
            if initiation_col:
                ws.cell(found_row, initiation_col).value = now
                ws.cell(found_row, initiation_col).number_format = "dd-mm-yyyy hh:mm:ss AM/PM"
            if deadline_col and receipt_col:
                receipt = ws.cell(found_row, receipt_col).value
                if isinstance(receipt, datetime):
                    deadline = receipt + timedelta(hours=24)
                    ws.cell(found_row, deadline_col).value = deadline
                    ws.cell(found_row, deadline_col).number_format = "dd-mm-yyyy hh:mm:ss AM/PM"
            ws.cell(found_row, status_col).value = "Inprogress"
            message = f"Job '{job_name}' assigned to {operator}."
        elif action_type == "move_to_qc":
            qc_col = src_headers.get("qc")
            qc_name = str(action.get("qc_name") or "").strip()
            if not qc_col:
                raise ValueError("The QC column was not found in the Inprogress sheet.")
            if not qc_name:
                raise ValueError("Please provide a QC name before confirming this action.")
            if current_status != "inprogress":
                raise ValueError(f"This job has status '{current_status}'. Only Inprogress jobs can move to QC.")
            # Validate the QC name against the existing Dropdown worksheet when available.
            if "Dropdown" in wb.sheetnames:
                dropdown_ws = wb["Dropdown"]
                qc_header_row = None
                qc_dropdown_col = None
                for rr in range(1, min(dropdown_ws.max_row, 20) + 1):
                    for cc in range(1, dropdown_ws.max_column + 1):
                        if _normalize_header(dropdown_ws.cell(rr, cc).value) == "qc":
                            qc_header_row, qc_dropdown_col = rr, cc
                            break
                    if qc_header_row is not None:
                        break
                if qc_header_row is not None:
                    allowed_qc = {
                        str(dropdown_ws.cell(rr, qc_dropdown_col).value).strip().casefold()
                        for rr in range(qc_header_row + 1, dropdown_ws.max_row + 1)
                        if dropdown_ws.cell(rr, qc_dropdown_col).value is not None
                    }
                    if qc_name.casefold() not in allowed_qc:
                        raise ValueError("The selected QC name is not in the Dropdown worksheet. Use an existing QC name.")
            ws.cell(found_row, qc_col).value = qc_name
            ws.cell(found_row, status_col).value = "QC"
            message = f"Job '{job_name}' moved to QC and assigned to {qc_name}."
        elif action_type == "complete_job":
            if "Completed Jobs" not in wb.sheetnames:
                raise ValueError("The Completed Jobs worksheet was not found.")
            if current_status not in {"inprogress", "qc"}:
                raise ValueError(f"This job has status '{current_status}'. Only Inprogress/QC jobs can be completed.")
            complexity = str(action.get("complexity") or "").strip()
            total_pages = action.get("total_pages")
            operator = str(action.get("operator") or "").strip()
            qc_name = str(action.get("qc_name") or "").strip()
            if not complexity or total_pages is None or not operator or not qc_name:
                raise ValueError("Completion requires Complexity, Total Pages, Operator, and QC. Fill all fields in the confirmation form.")
            dest = wb["Completed Jobs"]
            dest_header_row, dest_headers = _find_sheet_headers(dest)
            if dest_header_row is None:
                raise ValueError("Could not find Job Name and Status headers in Completed Jobs.")
            destination_row = max(dest.max_row + 1, dest_header_row + 1)
            # Copy matching source columns by header and retain cell formatting.
            for dest_header, dest_col in dest_headers.items():
                src_col = src_headers.get(dest_header)
                if src_col:
                    src_cell, dest_cell = ws.cell(found_row, src_col), dest.cell(destination_row, dest_col)
                    dest_cell.value = src_cell.value
                    if src_cell.has_style:
                        dest_cell._style = copy(src_cell._style)
                    if src_cell.number_format:
                        dest_cell.number_format = src_cell.number_format
                    if src_cell.alignment:
                        dest_cell.alignment = copy(src_cell.alignment)
                    if src_cell.protection:
                        dest_cell.protection = copy(src_cell.protection)
            completion_values = {
                "status": "Delivered", "complexity": complexity,
                "total pages": int(total_pages), "operator": operator,
                "qc": qc_name, "delivered time": datetime.now()
            }
            for field, value in completion_values.items():
                col = dest_headers.get(field)
                if col:
                    dest.cell(destination_row, col).value = value
                    if field == "delivered time":
                        dest.cell(destination_row, col).number_format = "dd-mm-yyyy hh:mm:ss AM/PM"
            if ws.row_dimensions[found_row].height:
                dest.row_dimensions[destination_row].height = ws.row_dimensions[found_row].height
            # Remove source row only after the destination row has been populated in memory.
            ws.delete_rows(found_row, 1)
            message = f"Job '{job_name}' moved to Completed Jobs."
        else:
            raise ValueError("Unsupported workflow action.")
        wb.save(path)
    finally:
        wb.close()
    clear_workflow_cache()
    return message


def AIAssistantMode(show_header=True):
    """Chat, live reports and confirmed workflow actions."""
    if show_header:
        # Display the supplied SVG wordmark instead of the text heading.
        ai_assistant_svg = "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxNjAwIiBoZWlnaHQ9IjIyMCIgdmlld0JveD0iMCAwIDE2MDAgMjIwIiByb2xlPSJpbWciIGFyaWEtbGFiZWxsZWRieT0idGl0bGUgZGVzYyI+CiAgPHRpdGxlIGlkPSJ0aXRsZSI+QUkgQXNzaXN0YW50IE1vZGU8L3RpdGxlPgogIDxkZXNjIGlkPSJkZXNjIj5BSSBBc3Npc3RhbnQgTW9kZSB3b3JkbWFyayB3aXRoIGEgYmx1ZS10by1wdXJwbGUtdG8tbWFnZW50YSBncmFkaWVudCwgd2F2ZWZvcm0gaWNvbiwgYW5kIHNwYXJrbGUuPC9kZXNjPgogIDxkZWZzPgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJicmFuZEdyYWRpZW50IiB4MT0iMCIgeTE9IjAiIHgyPSIxIiB5Mj0iMCI+CiAgICAgIDxzdG9wIG9mZnNldD0iMCUiIHN0b3AtY29sb3I9IiM1NkQ1RTUiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIyOCUiIHN0b3AtY29sb3I9IiM3RjlDRUIiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSI1NiUiIHN0b3AtY29sb3I9IiM5QTY3REUiLz4KICAgICAgPHN0b3Agb2Zmc2V0PSI3OCUiIHN0b3AtY29sb3I9IiNDNTRGQzciLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxMDAlIiBzdG9wLWNvbG9yPSIjRTUyRUE5Ii8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogICAgPGxpbmVhckdyYWRpZW50IGlkPSJ3YXZlR3JhZGllbnQiIHgxPSIwIiB5MT0iMCIgeDI9IjEiIHkyPSIwIj4KICAgICAgPHN0b3Agb2Zmc2V0PSIwJSIgc3RvcC1jb2xvcj0iIzZFOUNFQiIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjUyJSIgc3RvcC1jb2xvcj0iI0ExNUJERCIvPgogICAgICA8c3RvcCBvZmZzZXQ9IjEwMCUiIHN0b3AtY29sb3I9IiNENzQzQjkiLz4KICAgIDwvbGluZWFyR3JhZGllbnQ+CiAgPC9kZWZzPgogIDwhLS0gV2F2ZWZvcm0gbWFyayAtLT4KICA8ZyBmaWxsPSJ1cmwoI3dhdmVHcmFkaWVudCkiPgogICAgPHJlY3QgeD0iMjIiIHk9Ijg4IiB3aWR0aD0iMTgiIGhlaWdodD0iNDgiIHJ4PSI5Ii8+CiAgICA8cmVjdCB4PSI0NSIgeT0iNTciIHdpZHRoPSIxOSIgaGVpZ2h0PSIxMTAiIHJ4PSI5LjUiLz4KICAgIDxyZWN0IHg9IjY5IiB5PSIzMSIgd2lkdGg9IjIxIiBoZWlnaHQ9IjE2MCIgcng9IjEwLjUiLz4KICAgIDxyZWN0IHg9Ijk2IiB5PSI2MSIgd2lkdGg9IjE5IiBoZWlnaHQ9IjEwNSIgcng9IjkuNSIvPgogICAgPHJlY3QgeD0iMTIwIiB5PSI0OCIgd2lkdGg9IjIxIiBoZWlnaHQ9IjEyMyIgcng9IjEwLjUiLz4KICAgIDxyZWN0IHg9IjE0NiIgeT0iNzMiIHdpZHRoPSIyMCIgaGVpZ2h0PSI3NiIgcng9IjEwIi8+CiAgPC9nPgogIDwhLS0gV29yZG1hcms7IGVkaXRhYmxlIHRleHQgaW4gdmVjdG9yIGVkaXRvcnMgLS0+CiAgPHRleHQgeD0iMjA1IiB5PSIxNzEiIGZvbnQtZmFtaWx5PSJBcmlhbCwgSGVsdmV0aWNhLCBzYW5zLXNlcmlmIiBmb250LXNpemU9IjE1MSIgZm9udC13ZWlnaHQ9IjcwMCIgbGV0dGVyLXNwYWNpbmc9Ii01IiBmaWxsPSJ1cmwoI2JyYW5kR3JhZGllbnQpIj5BSSBBc3Npc3RhbnQgTW9kZTwvdGV4dD4KICA8IS0tIEZvdXItcG9pbnQgc3BhcmtsZSAtLT4KICA8cGF0aCBkPSJNMTUxMCAxMiBDMTUxNiAyOSAxNTIyIDM1IDE1NDAgNDEgQzE1MjIgNDcgMTUxNiA1MyAxNTEwIDcwIEMxNTA0IDUzIDE0OTggNDcgMTQ4MCA0MSBDMTQ5OCAzNSAxNTA0IDI5IDE1MTAgMTIgWiIgZmlsbD0idXJsKCNicmFuZEdyYWRpZW50KSIvPgo8L3N2Zz4K"
        st.markdown(
            f'<div style="width:100%; max-width:1000px; margin:0 0 0.5rem 0;">'
            f'<img src="data:image/svg+xml;base64,{ai_assistant_svg}" alt="AI Assistant Mode" '
            f'style="display:block; width:100%; height:auto;" />'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.caption("Ask about jobs, request a report, or propose a workflow action. Changes require your confirmation.")
    if "ai_messages" not in st.session_state:
        st.session_state.ai_messages = [{"role": "assistant", "content": "Hello! I can search jobs, report status counts, review deadlines and workload, or prepare an assignment/QC/completion action for confirmation."}]
    if "ai_pending_action" not in st.session_state:
        st.session_state.ai_pending_action = None

    welcome_message = (
        "Hello! I can search jobs, report status counts, review deadlines and workload, "
        "or prepare an assignment/QC/completion action for confirmation."
    )
    for msg in st.session_state.ai_messages[-12:]:
        if msg.get("role") == "assistant" and msg.get("content") == welcome_message:
            continue
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    with st.expander("Example requests", expanded=False):
        st.markdown("- How many Inprogress jobs are there?\n- Show overdue jobs.\n- Summarize workload by operator.\n- Assign EMLI Brochure to Ravi.\n- Move EMLI Brochure to QC and assign Priya.\n- Complete EMLI Brochure.")

    prompt = st.chat_input("Ask about jobs or request an action...", key="ai_chat_input")
    if prompt:
        st.session_state.ai_messages.append({"role": "user", "content": prompt})
        try:
            snapshot = _read_workflow_snapshot()
            intent = _ai_classify_request(prompt, snapshot)
            if intent["intent"] in {"assign_job", "move_to_qc", "complete_job"}:
                st.session_state.ai_pending_action = intent
                explanation = {
                    "assign_job": "I prepared a job assignment. Review the details below; nothing has been changed yet.",
                    "move_to_qc": "I prepared a move-to-QC action. Review the details below; nothing has been changed yet.",
                    "complete_job": "I prepared a completion action. Review the details below; nothing has been changed yet."
                }[intent["intent"]]
                st.session_state.ai_messages.append({"role": "assistant", "content": explanation})
            else:
                answer = _ai_answer_from_snapshot(prompt, snapshot)
                st.session_state.ai_messages.append({"role": "assistant", "content": answer})
        except Exception as exc:
            st.session_state.ai_messages.append({"role": "assistant", "content": f"I couldn't process that request: {exc}"})
        st.rerun()

    pending = st.session_state.ai_pending_action
    if pending:
        st.markdown("#### Review proposed action")
        intent = pending.get("intent")
        st.info({"assign_job": "Assign Job", "move_to_qc": "Move Job to QC", "complete_job": "Move Job to Completed"}.get(intent, "Workflow action"))
        # Allow the user to complete missing details before confirmation.
        job_name = st.text_input("Job Name", value=pending.get("job_name") or "", key="ai_action_job")
        operator = st.text_input("Operator", value=pending.get("operator") or "", key="ai_action_operator") if intent in {"assign_job", "complete_job"} else pending.get("operator")
        qc_name = st.text_input("QC Name", value=pending.get("qc_name") or "", key="ai_action_qc") if intent in {"move_to_qc", "complete_job"} else pending.get("qc_name")
        complexity = st.text_input("Complexity", value=pending.get("complexity") or "", key="ai_action_complexity") if intent == "complete_job" else pending.get("complexity")
        total_pages = st.number_input("Total Pages", min_value=1, step=1, value=int(pending.get("total_pages") or 1), key="ai_action_pages") if intent == "complete_job" else pending.get("total_pages")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Confirm and Execute", key="ai_confirm_action", type="primary", use_container_width=True):
                action = {**pending, "job_name": job_name.strip(), "operator": operator, "qc_name": qc_name, "complexity": complexity, "total_pages": total_pages}
                try:
                    result = _perform_confirmed_workflow_action(action)
                    st.session_state.ai_messages.append({"role": "assistant", "content": "✅ " + result})
                    st.session_state.ai_pending_action = None
                    for k in ("ai_action_job", "ai_action_operator", "ai_action_qc", "ai_action_complexity", "ai_action_pages"):
                        st.session_state.pop(k, None)
                    st.success(result)
                    st.rerun()
                except Exception as exc:
                    st.error(f"Action not completed: {exc}")
        with c2:
            if st.button("Cancel Action", key="ai_cancel_action", use_container_width=True):
                st.session_state.ai_pending_action = None
                st.session_state.ai_messages.append({"role": "assistant", "content": "Action cancelled. No workbook changes were made."})
                for k in ("ai_action_job", "ai_action_operator", "ai_action_qc", "ai_action_complexity", "ai_action_pages"):
                    st.session_state.pop(k, None)
                st.rerun()

    if st.button("Clear AI Chat", key="ai_clear_chat"):
        st.session_state.ai_messages = [{"role": "assistant", "content": "Chat cleared. What would you like to know about the workflow?"}]
        st.session_state.ai_pending_action = None
        st.rerun()



# -----------------------------------------------------
# Fixed modern web-app shell: left navigation, main page, AI panel
# -----------------------------------------------------
st.markdown("""
<style>
:root {
  --brand: #FF7900;
  --brand-dark: #D96500;
  --brand-soft: #FFF2E8;
  --page-bg: #F1F3F6;
  --surface: #FFFFFF;
  --text: #20242B;
  --muted: #737D8D;
  --line: #E2E6EC;
}
.stApp { background: var(--page-bg); color: var(--text); }
[data-testid="stHeader"] { background: rgba(241,243,246,.96); }
[data-testid="stMainBlockContainer"] {
  max-width: 1900px;
  padding: 1rem 1.1rem 2rem 1.1rem;
}
section[data-testid="stSidebar"] {
  background: #FFFFFF;
  border-right: 1px solid #DDE2E9;
}
section[data-testid="stSidebar"] > div { padding-top: .9rem; }
.app-brand {
  display:flex; align-items:center; gap:11px;
  padding:7px 3px 18px; margin-bottom:13px;
  border-bottom:1px solid var(--line);
}
.app-brand-icon {
  width:43px; height:43px; flex:0 0 43px;
  display:flex; align-items:center; justify-content:center;
  border-radius:13px; color:#fff; font-size:22px; font-weight:800;
  background:linear-gradient(135deg,#FF7900,#FFA64D);
  box-shadow:0 5px 14px rgba(255,121,0,.22);
}
.app-brand-name { color:#20242B; font-size:15px; line-height:1.25; font-weight:760; }
.app-brand-caption { margin-top:4px; color:#858B98; font-size:9px; letter-spacing:1px; font-weight:750; }
.app-heading { padding:2px 0 12px; }
.app-heading h1 {
  font-size:clamp(1.55rem,1.9vw,2rem)!important;
  font-weight:760!important; letter-spacing:-.04em;
  color:#20242B!important; margin:0 0 5px!important;
}
.app-heading p { color:#737D8D; font-size:13px; margin:0; }
.ai-panel-card {
  background:#FFFFFF;
  border:1px solid #E0E5EC;
  border-radius:20px 20px 0 0;
  box-shadow:0 14px 36px rgba(22,34,51,.13), 0 2px 7px rgba(22,34,51,.04);
  padding:22px 20px 8px;
  min-height:0;
  margin-bottom:0;
}
/* White card is scoped only to the AI Assistant container */
.st-key-ai_assistant_scope {
  background:#FFFFFF;
  border:1px solid #E0E5EC;
  border-radius:20px;
  box-shadow:0 14px 36px rgba(22,34,51,.13), 0 2px 7px rgba(22,34,51,.04);
  padding:0 0 42px;
  margin-top:18px;
  margin-bottom:8px;
  gap:0.55rem;
}
.st-key-ai_assistant_scope .ai-panel-card {
  border:0;
  border-radius:20px 20px 0 0;
  box-shadow:none;
}
.st-key-ai_assistant_scope div[data-testid="stChatMessage"],
.st-key-ai_assistant_scope div[data-testid="stExpander"],
.st-key-ai_assistant_scope div[data-testid="stChatInput"],
.st-key-ai_assistant_scope div[data-testid="stButton"] {
  margin-left:16px;
  margin-right:16px;
}
.ai-panel-heading {
  display:flex; align-items:center; gap:10px;
  font-size:25px; font-weight:760; letter-spacing:-.7px;
  color:#20242B; line-height:1.2;
}
.ai-panel-spark { color:#FF7900; font-size:30px; line-height:1; }
.ai-panel-subtitle { color:#7A8493; font-size:13px; margin-top:10px; }
.ai-panel-rule { border:0; border-top:1px solid #E5E8EE; margin:20px 0 17px; }
.ai-panel-help {
  background:#F4F6F9; border-radius:12px; padding:12px 13px;
  color:#6F7887; font-size:13px; line-height:1.65; margin-bottom:16px;
}
.ai-panel-label { color:#A855F7; font-size:11px; font-weight:750; margin:0 0 8px; }
div[data-testid="stMetric"] {
  background:#FFFFFF; border:1px solid #E1E6ED;
  border-radius:18px; padding:18px 18px 16px;
  box-shadow:0 6px 16px rgba(24,36,56,.06);
  min-height:126px;
}
div[data-testid="stMetricLabel"] { color:#20242B; font-weight:600; font-size:1rem; }
div[data-testid="stMetricValue"] { color:#111827; font-weight:780; font-size:2.35rem; }

/* Job Status Summary cards: distinct pastel backgrounds and status icons */
.st-key-status_tbi_card div[data-testid="stMetric"] {
  background:linear-gradient(135deg,#FFF9F1 0%,#FFF0E0 100%);
  border-color:#F7DFC4; border-top:4px solid #FF7900;
}

.st-key-status_inprogress_card div[data-testid="stMetric"] {
  background:linear-gradient(135deg,#F5FAFF 0%,#E8F4FF 100%);
  border-color:#D7E9FB; border-top:4px solid #2583D8;
}

.st-key-status_qc_card div[data-testid="stMetric"] {
  background:linear-gradient(135deg,#F6FAFF 0%,#EAF5FC 100%);
  border-color:#D8EAF6; border-top:4px solid #1677D2;
}

.st-key-status_completed_card div[data-testid="stMetric"] {
  background:linear-gradient(135deg,#F3FFF8 0%,#E5F8EE 100%);
  border-color:#D2F0DE; border-top:4px solid #20A464;
}

.stButton > button, .stDownloadButton > button {
  min-height:40px; border-radius:10px; font-weight:650;
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
  background:#FF7900; color:white; border:1px solid #FF7900;
}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {
  background:#E66D00; border-color:#E66D00; color:white;
}
.stButton > button[kind="secondary"]:hover, .stDownloadButton > button[kind="secondary"]:hover {
  border-color:#FF7900; color:#C45A00;
}

/* Colorful dashboard section headings */
.section-heading {
  display:flex; align-items:center; gap:10px;
  font-size:1.35rem; font-weight:750; color:#20242B;
  letter-spacing:-.025em; margin: .3rem 0 1rem 0;
}
.section-heading .section-icon {
  width:36px; height:36px; flex:0 0 36px;
  display:inline-flex; align-items:center; justify-content:center;
  border-radius:11px; font-size:19px;
}
.section-icon.orange { background:#FFF0E3; color:#FF7900; }
.section-icon.blue { background:#E8F2FF; color:#2672D8; }
.section-icon.green { background:#E5F8EE; color:#15945B; }
.section-icon.purple { background:#F3E8FF; color:#8B5CF6; }

/* Distinct form controls: white fill and visible neutral border on the grey page */
div[data-testid="stTextInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-testid="stDateInput"] input,
div[data-testid="stNumberInput"] input {
  background-color: #FFFFFF !important;
  border: 1px solid #CBD2DC !important;
  border-radius: 9px !important;
  color: #20242B !important;
  box-shadow: 0 1px 2px rgba(16,24,40,.04);
}
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
  background-color: #FFFFFF !important;
  border: 1px solid #CBD2DC !important;
  border-radius: 9px !important;
  color: #20242B !important;
  min-height: 42px;
  box-shadow: 0 1px 2px rgba(16,24,40,.04);
}
div[data-testid="stSelectbox"] div[data-baseweb="select"] span,
div[data-testid="stMultiSelect"] div[data-baseweb="select"] span {
  color: #20242B !important;
}
div[data-testid="stSelectbox"] div[data-baseweb="select"]:hover > div,
div[data-testid="stMultiSelect"] div[data-baseweb="select"]:hover > div {
  border-color: #FF7900 !important;
}
div[data-testid="stSelectbox"] div[data-baseweb="select"]:focus-within > div,
div[data-testid="stMultiSelect"] div[data-baseweb="select"]:focus-within > div,
div[data-testid="stTextInput"] input:focus,
div[data-testid="stTextArea"] textarea:focus,
div[data-testid="stDateInput"] input:focus,
div[data-testid="stNumberInput"] input:focus {
  border-color: #FF7900 !important;
  box-shadow: 0 0 0 2px rgba(255,121,0,.16) !important;
}
div[data-testid="stSelectbox"] [role="listbox"],
div[data-testid="stMultiSelect"] [role="listbox"] {
  background-color: #FFFFFF !important;
  border: 1px solid #D8DEE7 !important;
  border-radius: 10px !important;
  box-shadow: 0 12px 28px rgba(16,24,40,.14) !important;
}
/* Keep chatbot controls close below the welcome message */
div[data-testid="stChatInput"] { margin-top: 2px !important; }
div[data-testid="stExpander"] { margin-bottom: 8px !important; }
div[data-testid="stButton"] { margin-top: 4px; }
div[data-testid="stDataFrame"], div[data-testid="stTable"] {
  border:1px solid #E2E6EC; border-radius:11px; overflow:hidden;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label {
  padding:8px 10px; border-radius:9px;
}
section[data-testid="stSidebar"] div[role="radiogroup"] label:hover { background:#FFF2E8; }
/* AI Assistant attention button: repeating orange glow */
@keyframes aiAssistantPulse {
  0%, 100% {
    background-color: #FFF2E8;
    border-color: #FF7900;
    color: #A84D00;
    box-shadow: 0 0 0 0 rgba(255,121,0,.12);
  }
  50% {
    background-color: #FF7900;
    border-color: #FF7900;
    color: #FFFFFF;
    box-shadow: 0 0 0 5px rgba(255,121,0,.16);
  }
}
section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
  animation: aiAssistantPulse 1.6s ease-in-out infinite !important;
  border: 1px solid #FF7900 !important;
  font-weight: 750 !important;
}
section[data-testid="stSidebar"] div[data-testid="stButton"] button {
  transition: box-shadow .2s ease, background-color .2s ease, color .2s ease;
}
@media (prefers-reduced-motion: reduce) {
  section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
    animation: none !important;
    background: #FFF0E3 !important;
    color: #A84D00 !important;
  }
}

hr { border-color:#E2E6EC; }
footer { visibility:hidden; }
@media (max-width: 1100px) {
  [data-testid="stMainBlockContainer"] { padding:.8rem .6rem 1.5rem; }
  .ai-panel-card { padding:16px 13px; }
  .ai-panel-heading { font-size:21px; }
}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("""
    <div class="app-brand">
      <div class="app-brand-icon">W</div>
      <div>
        <div class="app-brand-name">Automation<br>Workflow System</div>
        <div class="app-brand-caption">JOB OPERATIONS PORTAL</div>
      </div>
    </div>
    """, unsafe_allow_html=True)
    # Keep the AI Assistant control above the WORKSPACE navigation section.
    if "show_ai_panel" not in st.session_state:
        st.session_state["show_ai_panel"] = False
    if st.button(
        "✦  AI Assistant" if not st.session_state["show_ai_panel"] else "✦  Hide AI Assistant",
        key="toggle_ai_assistant",
        use_container_width=True,
        type="primary" if st.session_state["show_ai_panel"] else "secondary",
    ):
        st.session_state["show_ai_panel"] = not st.session_state["show_ai_panel"]
        st.rerun()
    st.markdown("---")

    st.caption("WORKSPACE")
    selected_page = st.radio(
        "Main navigation",
        ["Dashboard", "Management Insights", "Job Management", "Deadline Monitoring", "Employee Workload", "Job Search",
         "Edit Job Details", "Job History & Audit Log", "Reports & Export", "Email Automation"],
        key="modern_main_navigation",
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.markdown(
        '<div style="font-size:12px;color:#858B98;line-height:1.7;">'
        '<span style="color:#FF7900;font-weight:700;">●</span> Workflow workspace'
        '<br>Excel-backed job tracking</div>',
        unsafe_allow_html=True,
    )



# -----------------------------------------------------
# Management-level Operations Intelligence
# -----------------------------------------------------
def ManagementInsights():
    """Management KPIs and reports derived only from workflow.xlsx records."""
    st.markdown("## Operations Intelligence")
    st.caption(
        "Management view of delivery performance, workload, workflow queues, and exceptions. "
        "Metrics are calculated from available workbook fields only."
    )

    workbook_path = "workflow.xlsx"
    if not os.path.exists(workbook_path):
        st.error("workflow.xlsx was not found. Management insights cannot be calculated.")
        return

    def norm(value):
        return re.sub(r"\s+", " ", str(value or "").strip().casefold().replace("_", " "))

    def find_header_row(ws):
        for row_num in range(1, min(ws.max_row, 20) + 1):
            vals = [norm(ws.cell(row_num, c).value) for c in range(1, ws.max_column + 1)]
            if "job name" in vals and "status" in vals:
                return row_num, vals
        return None, []

    def pick(record, aliases):
        aliases = {norm(a) for a in aliases}
        for key, value in record.items():
            if norm(key) in aliases:
                return value
        # Allow common expanded header variants.
        for key, value in record.items():
            nk = norm(key)
            if any(a in nk for a in aliases):
                return value
        return None

    records = []
    try:
        wb = load_workbook(workbook_path, read_only=False, data_only=True)
        try:
            for sheet_name in ("Inprogress", "Completed Jobs"):
                if sheet_name not in wb.sheetnames:
                    continue
                ws = wb[sheet_name]
                header_row, normalized_headers = find_header_row(ws)
                if header_row is None:
                    continue
                headers = [
                    str(ws.cell(header_row, c).value or f"Column {c}").strip()
                    for c in range(1, ws.max_column + 1)
                ]
                for row_num in range(header_row + 1, ws.max_row + 1):
                    vals = [ws.cell(row_num, c).value for c in range(1, len(headers) + 1)]
                    if not any(v is not None and str(v).strip() for v in vals):
                        continue
                    record = {headers[i]: vals[i] for i in range(len(headers))}
                    job_name = pick(record, ["job name", "job_name"])
                    status = pick(record, ["status"])
                    if not job_name and not status:
                        continue
                    record["__source_sheet"] = sheet_name
                    record["__job_name"] = str(job_name or "(Unnamed job)").strip()
                    record["__status"] = str(status or "Unknown").strip()
                    records.append(record)
        finally:
            wb.close()
    except PermissionError:
        st.error("Close workflow.xlsx in Excel, then refresh the dashboard.")
        return
    except Exception as exc:
        st.error(f"Unable to read management metrics from workflow.xlsx: {exc}")
        return

    if not records:
        st.info("No job records were found in the Inprogress or Completed Jobs sheets.")
        return

    df = pd.DataFrame(records)
    df["__status_norm"] = df["__status"].astype(str).str.strip().str.casefold()
    active_statuses = {"tbi", "inprogress", "in progress", "qc"}
    delivered_statuses = {"delivered", "completed", "complete"}
    df["__is_active"] = df["__status_norm"].isin(active_statuses)
    df["__is_delivered"] = df["__status_norm"].isin(delivered_statuses)

    def as_datetime(value):
        if value is None or (isinstance(value, str) and not value.strip()):
            return pd.NaT
        try:
            parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
            if pd.isna(parsed):
                return pd.NaT
            parsed = pd.Timestamp(parsed)
            if parsed.tzinfo is not None:
                parsed = parsed.tz_localize(None)
            return parsed
        except Exception:
            return pd.NaT

    def add_datetime_column(aliases, target):
        df[target] = df.apply(lambda row: as_datetime(pick(row.to_dict(), aliases)), axis=1)

    add_datetime_column(["deadline ist", "deadline", "due date", "due time"], "__deadline")
    add_datetime_column(["delivered time", "completion time", "completed time", "delivery time"], "__delivered_at")
    add_datetime_column(["receipt time", "received time", "date received", "date"], "__received_at")
    add_datetime_column(["initiation time", "started time", "start time"], "__started_at")

    operator_values = df.apply(
        lambda row: pick(row.to_dict(), ["operator", "assigned operator", "employee", "employee name"]),
        axis=1,
    )
    df["__operator"] = operator_values.fillna("").astype(str).str.strip()
    df.loc[df["__operator"].str.casefold().isin(["", "nan", "none", "not assigned", "unassigned"]), "__operator"] = "Unassigned"

    now = pd.Timestamp(datetime.now())
    total_jobs = len(df)
    active_df = df[df["__is_active"]].copy()
    delivered_df = df[df["__is_delivered"]].copy()
    overdue_df = active_df[active_df["__deadline"].notna() & (active_df["__deadline"] < now)].copy()
    due_24h_df = active_df[
        active_df["__deadline"].notna()
        & (active_df["__deadline"] >= now)
        & (active_df["__deadline"] <= now + pd.Timedelta(hours=24))
    ].copy()

    # SLA compliance is only computed when both delivery and deadline timestamps exist.
    sla_df = delivered_df[delivered_df["__deadline"].notna() & delivered_df["__delivered_at"].notna()].copy()
    if not sla_df.empty:
        sla_df["__on_time"] = sla_df["__delivered_at"] <= sla_df["__deadline"]
        sla_pct = 100 * float(sla_df["__on_time"].mean())
        sla_text = f"{sla_pct:.1f}%"
        sla_help = f"{int(sla_df['__on_time'].sum())} of {len(sla_df)} delivered jobs have usable deadline and delivery timestamps."
    else:
        sla_text = "N/A"
        sla_help = "Requires both Deadline IST and Delivered Time for delivered jobs."

    # KPI row
    st.markdown("""
    <style>
    .mgmt-kpi {border:1px solid #E1E6ED;border-radius:16px;padding:16px 18px;
      background:#FFFFFF;box-shadow:0 5px 15px rgba(24,36,56,.05);min-height:112px;}
    .mgmt-kpi-label {font-size:13px;color:#687386;margin-bottom:8px;}
    .mgmt-kpi-value {font-size:30px;font-weight:750;line-height:1.1;color:#172033;}
    .mgmt-kpi-foot {font-size:11px;color:#7B8493;margin-top:7px;}
    </style>
    """, unsafe_allow_html=True)

    kpis = [
        ("Total Jobs", f"{total_jobs:,}", "Records in both workflow sheets"),
        ("Active Jobs", f"{len(active_df):,}", "TBI, Inprogress, and QC"),
        ("Delivered", f"{len(delivered_df):,}", "Completed/delivered records"),
        ("Overdue", f"{len(overdue_df):,}", "Active jobs past deadline"),
        ("Due in 24 Hours", f"{len(due_24h_df):,}", "Active jobs with a deadline in next 24h"),
        ("On-time Delivery", sla_text, sla_help),
    ]
    kpi_cols = st.columns(3)
    for idx, (label, value, foot) in enumerate(kpis):
        with kpi_cols[idx % 3]:
            st.markdown(
                f'<div class="mgmt-kpi"><div class="mgmt-kpi-label">{label}</div>'
                f'<div class="mgmt-kpi-value">{value}</div>'
                f'<div class="mgmt-kpi-foot">{foot}</div></div>',
                unsafe_allow_html=True,
            )
    st.markdown("")

    # Status distribution and trend analysis
    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("### Workflow Status Mix")
        status_counts = (
            df["__status"].fillna("Unknown").astype(str).str.strip()
            .replace("", "Unknown").value_counts().rename_axis("Status").reset_index(name="Jobs")
        )
        st.plotly_chart(
            px.bar(status_counts, x="Status", y="Jobs", color="Status",
                   title="Jobs by current status", template="plotly_white"),
            use_container_width=True,
            key="mgmt_status_mix_chart",
        )
    with right:
        st.markdown("### Delivery Trend")
        trend_df = delivered_df[delivered_df["__delivered_at"].notna()].copy()
        if not trend_df.empty:
            trend_df["Week"] = trend_df["__delivered_at"].dt.to_period("W").dt.start_time
            weekly = trend_df.groupby("Week").size().reset_index(name="Delivered Jobs").sort_values("Week").tail(12)
            st.plotly_chart(
                px.line(weekly, x="Week", y="Delivered Jobs", markers=True,
                        title="Delivered jobs per week (last 12 available weeks)", template="plotly_white"),
                use_container_width=True,
                key="mgmt_delivery_trend_chart",
            )
        else:
            st.info("Delivery trend is unavailable because Delivered Time is missing or has no valid dates.")

    # Team workload and productivity
    st.markdown("### Team Productivity & Workload")
    team_df = df[df["__operator"] != "Unassigned"].groupby("__operator").agg(
        Total_Jobs=("__job_name", "count"),
        Active_Jobs=("__is_active", "sum"),
        Delivered_Jobs=("__is_delivered", "sum"),
    ).reset_index().rename(columns={"__operator": "Operator", "Total_Jobs": "Total Jobs",
                                    "Active_Jobs": "Active Jobs", "Delivered_Jobs": "Delivered Jobs"})
    unassigned_count = int((df["__operator"] == "Unassigned").sum())
    team_col, unassigned_col = st.columns([3, 1], gap="large")
    with team_col:
        if not team_df.empty:
            team_long = team_df.melt(id_vars="Operator", value_vars=["Active Jobs", "Delivered Jobs"],
                                     var_name="Work Type", value_name="Jobs")
            st.plotly_chart(
                px.bar(team_long, x="Operator", y="Jobs", color="Work Type", barmode="group",
                       title="Active vs delivered jobs by operator", template="plotly_white"),
                use_container_width=True,
                key="mgmt_team_workload_chart",
            )
            st.dataframe(team_df.sort_values(["Active Jobs", "Total Jobs"], ascending=False),
                         use_container_width=True, hide_index=True)
        else:
            st.info("Operator workload chart is unavailable because no assigned operator values were found.")
    with unassigned_col:
        st.metric("Unassigned records", unassigned_count)
        st.caption("Includes records with a blank, unassigned, or not-assigned operator field.")

    # Bottleneck analysis: current queue sizes and available age signals.
    st.markdown("### Workflow Bottleneck Analysis")
    active_status_counts = active_df["__status"].astype(str).str.strip().value_counts()
    stage_cols = st.columns(3)
    for i, stage in enumerate(["TBI", "Inprogress", "QC"]):
        with stage_cols[i]:
            count = int(active_df["__status_norm"].isin({stage.casefold(), "in progress"} if stage == "Inprogress" else {stage.casefold()}).sum())
            st.metric(f"{stage} Queue", count)
    age_source = active_df["__started_at"].fillna(active_df["__received_at"])
    aged_mask = age_source.notna()
    if aged_mask.any():
        active_df["__age_hours"] = (now - age_source).dt.total_seconds() / 3600
        age_summary = active_df[active_df["__age_hours"].notna()].groupby("__status")["__age_hours"].agg(["count", "mean", "max"]).reset_index()
        age_summary = age_summary.rename(columns={"__status": "Status", "count": "Jobs with timestamps",
                                                  "mean": "Average age (hours)", "max": "Oldest age (hours)"})
        age_summary[["Average age (hours)", "Oldest age (hours)"]] = age_summary[
            ["Average age (hours)", "Oldest age (hours)"]
        ].round(1)
        st.dataframe(age_summary, use_container_width=True, hide_index=True)
    else:
        st.info("Queue ageing is not calculated because Receipt Time/Initiation Time values are missing or invalid.")

    # Exception list for management action.
    st.markdown("### Management Attention Required")
    exception_frames = []
    if not overdue_df.empty:
        overdue_view = overdue_df.copy()
        overdue_view["Attention Reason"] = "Overdue"
        exception_frames.append(overdue_view)
    if not due_24h_df.empty:
        due_view = due_24h_df.copy()
        due_view["Attention Reason"] = "Due within 24 hours"
        exception_frames.append(due_view)
    if exception_frames:
        exceptions = pd.concat(exception_frames, ignore_index=True).drop_duplicates(subset=["__source_sheet", "__job_name", "__status"])
        exception_view = pd.DataFrame({
            "Job Name": exceptions["__job_name"],
            "Status": exceptions["__status"],
            "Operator": exceptions["__operator"],
            "Deadline": exceptions["__deadline"].dt.strftime("%d-%b-%Y %I:%M %p"),
            "Attention Reason": exceptions["Attention Reason"],
            "Source Sheet": exceptions["__source_sheet"],
        })
        st.dataframe(exception_view, use_container_width=True, hide_index=True)
    else:
        st.success("No active jobs are overdue or due within the next 24 hours based on valid deadline data.")

    # Export a management summary from the same calculated records.
    st.markdown("### Management Report Export")
    export_df = df.drop(columns=[c for c in df.columns if c.startswith("__")], errors="ignore").copy()
    export_df.insert(0, "Source Sheet", df["__source_sheet"].values)
    export_csv = export_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Download Management Report (CSV)",
        data=export_csv,
        file_name=f"workflow_management_report_{datetime.now():%Y%m%d_%H%M}.csv",
        mime="text/csv",
        key="mgmt_download_csv",
    )
    st.caption("This report is generated from the current workbook snapshot. No workbook data is changed by this dashboard.")



# -----------------------------------------------------
# Job Details Editor + Workbook Audit Log
# -----------------------------------------------------
def _audit_normalize(value):
    return re.sub(r"\s+", " ", str(value or "").strip().casefold().replace("_", " "))


def _audit_find_headers(ws, max_scan=20):
    """Return (header_row, headers) by locating a row with Job Name and Status."""
    for row_num in range(1, min(ws.max_row, max_scan) + 1):
        headers = [ws.cell(row_num, col).value for col in range(1, ws.max_column + 1)]
        normalized = [_audit_normalize(value) for value in headers]
        if "job name" in normalized and "status" in normalized:
            return row_num, headers
    return None, []


def _audit_job_label(ws, header_row, headers, row_num):
    vals = [ws.cell(row_num, col).value for col in range(1, len(headers) + 1)]
    record = {_audit_normalize(headers[i]): vals[i] for i in range(len(headers)) if headers[i] is not None}
    job_name = record.get("job name", "")
    wo_number = record.get("wo number", record.get("work order number", ""))
    project_number = record.get("project number", "")
    parts = [str(x).strip() for x in (job_name, wo_number, project_number) if x is not None and str(x).strip()]
    return " | ".join(parts) if parts else f"Row {row_num}"


def _audit_get_records(wb):
    records = []
    for sheet_name in ("Inprogress", "Completed Jobs"):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row, headers = _audit_find_headers(ws)
        if header_row is None:
            continue
        for row_num in range(header_row + 1, ws.max_row + 1):
            values = [ws.cell(row_num, col).value for col in range(1, len(headers) + 1)]
            if not any(value is not None and str(value).strip() for value in values):
                continue
            label = _audit_job_label(ws, header_row, headers, row_num)
            job_name_col = next((i for i, h in enumerate(headers, 1) if _audit_normalize(h) == "job name"), None)
            job_name = ws.cell(row_num, job_name_col).value if job_name_col else None
            status_col = next((i for i, h in enumerate(headers, 1) if _audit_normalize(h) == "status"), None)
            status = ws.cell(row_num, status_col).value if status_col else None
            records.append({
                "sheet": sheet_name, "row": row_num, "header_row": header_row,
                "headers": headers, "label": label,
                "job_name": str(job_name or "").strip(), "status": str(status or "").strip(),
            })
    return records


def _audit_append(wb, entries):
    """Append field-level changes to a dedicated worksheet without altering source sheets."""
    sheet_name = "Audit Log"
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)
        ws.append(["Timestamp", "Action", "Sheet", "Excel Row", "Job Identifier", "Field", "Old Value", "New Value", "Edited By"])
        from openpyxl.styles import Font, PatternFill, Alignment
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="FF7900")
            cell.alignment = Alignment(vertical="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col, width in {"A": 21, "B": 18, "C": 18, "D": 12, "E": 36, "F": 24, "G": 32, "H": 32, "I": 22}.items():
            ws.column_dimensions[col].width = width
    if ws.max_row == 0:
        ws.append(["Timestamp", "Action", "Sheet", "Excel Row", "Job Identifier", "Field", "Old Value", "New Value", "Edited By"])
    for item in entries:
        ws.append([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Edit Job Details",
            item["sheet"], item["row"], item["label"], item["field"],
            item["old"], item["new"], item["editor"],
        ])
        for cell in ws[ws.max_row]:
            cell.alignment = __import__("openpyxl").styles.Alignment(vertical="top", wrap_text=True)
    ws.auto_filter.ref = ws.dimensions


def EditJobDetails():
    st.markdown("## Edit Job Details")
    st.caption("Update an existing job record. Each changed field is recorded in the workbook's Audit Log sheet.")
    workbook_path = "workflow.xlsx"
    if not os.path.exists(workbook_path):
        st.error("workflow.xlsx was not found in the application folder.")
        return
    try:
        wb = load_workbook(workbook_path)
    except PermissionError:
        st.error("Please close workflow.xlsx in Excel and try again.")
        return
    except Exception as exc:
        st.error(f"Unable to open workflow.xlsx: {exc}")
        return
    try:
        records = _audit_get_records(wb)
        if not records:
            st.warning("No job records were found in the Inprogress or Completed Jobs sheets.")
            return
        choices = [f"{r['sheet']} — {r['label']} — {r['status'] or 'No status'}" for r in records]
        job_search = st.text_input(
            "Select a job to edit",
            placeholder="Type job name, job ID, or status to find a job...",
            key="edit_job_record_search",
        )
        if not job_search.strip():
            st.caption("Start typing a job name or ID to find the record you want to edit.")
            wb.close()
            return
        query = job_search.strip().casefold()
        matching_indices = [i for i, choice in enumerate(choices) if query in choice.casefold()]
        if not matching_indices:
            st.warning("No matching jobs found. Try another job name or ID.")
            wb.close()
            return
        if len(matching_indices) == 1:
            selected_index = matching_indices[0]
            st.caption(f"Selected: {choices[selected_index]}")
        else:
            st.caption(f"Found {len(matching_indices)} matching jobs. Choose the correct record below.")
            selected_index = st.selectbox(
                "Matching jobs", matching_indices,
                format_func=lambda idx: choices[idx], key="edit_job_record_match_select",
            )
        record = records[selected_index]
        ws = wb[record["sheet"]]
        headers = record["headers"]
        header_row = record["header_row"]
        row_num = record["row"]
        st.info(f"Editing record in **{record['sheet']}**, Excel row **{row_num}**. Existing workbook sheets and cell formatting are retained; only submitted cell values are changed.")
        editor = st.text_input("Edited by (name or username)", key="edit_job_editor")
        changes = []
        with st.form("edit_job_details_form", clear_on_submit=False):
            st.markdown("### Job fields")
            columns = st.columns(2)
            for index, header in enumerate(headers, start=1):
                if header is None or not str(header).strip():
                    continue
                cell = ws.cell(row_num, index)
                # Formula cells are never offered for editing.
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue
                label = str(header).strip()
                current = cell.value
                if isinstance(current, datetime):
                    default_text = current.strftime("%d-%m-%Y %I:%M:%S %p")
                elif isinstance(current, date):
                    default_text = current.strftime("%d-%m-%Y")
                else:
                    default_text = "" if current is None else str(current)
                with columns[(index - 1) % 2]:
                    new_value = st.text_input(label, value=default_text, key=f"edit_job_field_{row_num}_{index}", help="For date/time fields, use the existing format where possible.")
                changes.append((index, label, current, new_value))
            submitted = st.form_submit_button("Save Job Changes", type="primary", use_container_width=True)
        if submitted:
            audit_entries = []
            for col_idx, label, old_value, new_text in changes:
                old_text = "" if old_value is None else (old_value.strftime("%d-%m-%Y %I:%M:%S %p") if isinstance(old_value, datetime) else old_value.strftime("%d-%m-%Y") if isinstance(old_value, date) else str(old_value))
                if new_text == old_text:
                    continue
                # Preserve date/datetime cell types when the existing cell uses them.
                new_value = new_text
                if isinstance(old_value, datetime):
                    parsed = pd.to_datetime(new_text, errors="coerce", dayfirst=True)
                    if pd.isna(parsed):
                        st.error(f"'{label}' expects a valid date/time. Use DD-MM-YYYY HH:MM:SS AM/PM.")
                        wb.close()
                        return
                    new_value = parsed.to_pydatetime()
                elif isinstance(old_value, date):
                    parsed = pd.to_datetime(new_text, errors="coerce", dayfirst=True)
                    if pd.isna(parsed):
                        st.error(f"'{label}' expects a valid date. Use DD-MM-YYYY.")
                        wb.close()
                        return
                    new_value = parsed.date()
                ws.cell(row_num, col_idx).value = new_value
                audit_entries.append({
                    "sheet": record["sheet"], "row": row_num, "label": record["label"],
                    "field": label, "old": old_text, "new": new_text, "editor": editor.strip() or "Not provided",
                })
            if not audit_entries:
                st.info("No changes detected. The workbook was not modified.")
            else:
                _audit_append(wb, audit_entries)
                try:
                    wb.save(workbook_path)
                except PermissionError:
                    st.error("Could not save because workflow.xlsx is open or locked. Close it and try again.")
                    return
                clear_workflow_cache()
                st.success(f"Saved {len(audit_entries)} field change(s). The changes were recorded in the Audit Log sheet.")
                st.rerun()
    finally:
        try:
            wb.close()
        except Exception:
            pass



@st.cache_data(ttl=30, show_spinner=False)
def _load_job_history_snapshot(workbook_path, modified_time):
    """Read history data sequentially and cache it until workflow.xlsx changes."""
    history_rows = []
    audit_rows = []
    audit_headers = []

    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        if "Audit Log" in wb.sheetnames:
            audit_ws = wb["Audit Log"]
            row_iter = audit_ws.iter_rows(values_only=True)
            audit_headers = [str(v or "") for v in next(row_iter, ())]
            for values in row_iter:
                if any(v is not None for v in values):
                    audit_rows.append(tuple(values))

        for sheet_name in ("Inprogress", "Completed Jobs"):
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            header_row = None
            headers = []
            # Sequential iteration avoids repeated random-access cell reads.
            for row_num, values in enumerate(ws.iter_rows(values_only=True), start=1):
                if row_num > 20:
                    break
                normalized = [_audit_normalize(v) for v in values]
                if "job name" in normalized and "status" in normalized:
                    header_row = row_num
                    headers = list(values)
                    break
            if header_row is None:
                continue

            for row_num, values in enumerate(
                ws.iter_rows(min_row=header_row + 1, max_col=len(headers), values_only=True),
                start=header_row + 1,
            ):
                values = tuple(values)
                if not any(v is not None and str(v).strip() for v in values):
                    continue
                normalized_record = {
                    _audit_normalize(headers[i]): values[i]
                    for i in range(min(len(headers), len(values)))
                    if headers[i] is not None
                }
                job_name = normalized_record.get("job name", "")
                wo_number = normalized_record.get("wo number", normalized_record.get("work order number", ""))
                project_number = normalized_record.get("project number", "")
                parts = [str(x).strip() for x in (job_name, wo_number, project_number)
                         if x is not None and str(x).strip()]
                label = " | ".join(parts) if parts else f"Row {row_num}"
                row_dict = {
                    str(headers[i] or f"Column {i+1}"): values[i]
                    for i in range(min(len(headers), len(values)))
                    if i < len(headers)
                }
                row_dict["Source Sheet"] = sheet_name
                history_rows.append((sheet_name, row_num, label, row_dict))
    finally:
        wb.close()

    return audit_headers, audit_rows, history_rows

def JobHistoryAuditLog():
    st.markdown("## Job History & Audit Log")
    st.caption("Review job changes recorded by this application. Existing workbook data is displayed read-only here.")
    workbook_path = "workflow.xlsx"
    if not os.path.exists(workbook_path):
        st.error("workflow.xlsx was not found in the application folder.")
        return

    try:
        modified_time = os.path.getmtime(workbook_path)
        with st.spinner("Loading job history…"):
            audit_headers, audit_rows, history_rows = _load_job_history_snapshot(
                workbook_path, modified_time
            )

        st.markdown("### Change Audit")
        if audit_rows:
            audit_df = pd.DataFrame(audit_rows, columns=audit_headers[:len(audit_rows[0])])
            filter_col1, filter_col2 = st.columns(2)
            with filter_col1:
                job_filter = st.text_input("Filter by job identifier", key="audit_job_filter")
            with filter_col2:
                action_filter = st.text_input("Filter by field or editor", key="audit_field_filter")

            # Filter only the relevant columns instead of scanning every column repeatedly.
            if job_filter.strip() and "Job Identifier" in audit_df.columns:
                audit_df = audit_df[
                    audit_df["Job Identifier"].astype(str).str.contains(
                        job_filter.strip(), case=False, na=False, regex=False
                    )
                ]
            if action_filter.strip():
                candidate_cols = [
                    col for col in ("Field", "Edited By", "Action", "Sheet")
                    if col in audit_df.columns
                ]
                if candidate_cols:
                    mask = pd.Series(False, index=audit_df.index)
                    for col in candidate_cols:
                        mask |= audit_df[col].astype(str).str.contains(
                            action_filter.strip(), case=False, na=False, regex=False
                        )
                    audit_df = audit_df[mask]

            st.dataframe(audit_df.iloc[::-1], use_container_width=True, hide_index=True)
            st.download_button(
                "Download Audit Log CSV",
                audit_df.to_csv(index=False).encode("utf-8-sig"),
                "workflow_audit_log.csv", "text/csv", key="download_audit_log",
            )
        else:
            st.info("No audit entries yet. Changes saved through Edit Job Details will be recorded here.")

        st.markdown("### Current Job Records")
        if history_rows:
            history_df = pd.DataFrame([item[3] for item in history_rows])
            search = st.text_input("Search current job history", key="current_job_history_search")
            if search.strip():
                # A single joined search string per row is faster than applying string conversion to every column.
                search_blob = history_df.fillna("").astype(str).agg(" | ".join, axis=1)
                history_df = history_df[
                    search_blob.str.contains(search.strip(), case=False, na=False, regex=False)
                ]
            st.dataframe(history_df, use_container_width=True, hide_index=True)
        else:
            st.info("No current job records were found.")
    except PermissionError:
        st.error("Unable to read workflow.xlsx. Close it in Excel and refresh.")
    except Exception as exc:
        st.error(f"Unable to load job history: {exc}")

def _background_gmail_notification_watcher():
    """Collect background Gmail events for the in-app Notification panel; never show toasts."""
    try:
        events = get_pending_job_notifications()
        if events:
            clear_workflow_cache()
            stored = st.session_state.setdefault("job_notifications", [])
            known_ids = {str(item.get("event_id", "")) for item in stored}
            for event in events:
                event_id = str(event.get("event_id") or f"{event.get('job_name','New job')}|{event.get('created_at','')}")
                if event_id not in known_ids:
                    item = dict(event)
                    item["event_id"] = event_id
                    item.setdefault("job_name", "New job")
                    item.setdefault("created_at", datetime.now().isoformat(timespec="seconds"))
                    stored.append(item)
                    known_ids.add(event_id)
            st.session_state["job_notifications"] = stored[-200:]
    except Exception:
        # Keep transient file/poller errors from breaking the workflow UI.
        pass


def _render_notification_center():
    """Show a Notification button and details directly underneath only when clicked."""
    st.session_state.setdefault("job_notifications", [])
    st.session_state.setdefault("read_job_notification_ids", [])
    st.session_state.setdefault("notification_panel_open", False)
    notifications = st.session_state["job_notifications"]
    read_ids = set(st.session_state["read_job_notification_ids"])
    unread_count = sum(1 for item in notifications if str(item.get("event_id", "")) not in read_ids)

    # Give the Notification control its own explicit styling so global button/theme
    # rules cannot make the label faint or leave a red underline/border.
    st.markdown("""
    <style>
    .st-key-job_notification_button button {
        background: #FFFFFF !important;
        color: #253047 !important;
        border: 1px solid #D5DCE6 !important;
        border-radius: 12px !important;
        min-height: 46px !important;
        font-weight: 600 !important;
        opacity: 1 !important;
        box-shadow: 0 3px 10px rgba(20, 30, 50, 0.10) !important;
        text-decoration: none !important;
    }
    .st-key-job_notification_button button p,
    .st-key-job_notification_button button span {
        color: #253047 !important;
        opacity: 1 !important;
    }
    .st-key-job_notification_button button:hover {
        background: #F8FAFC !important;
        color: #253047 !important;
        border-color: #B8C2D0 !important;
    }
    .st-key-job_notification_button button:focus,
    .st-key-job_notification_button button:focus-visible {
        outline: 2px solid #8B5CF6 !important;
        outline-offset: 2px !important;
        box-shadow: 0 0 0 2px rgba(139, 92, 246, 0.15) !important;
    }
    </style>
    """, unsafe_allow_html=True)

    left, right = st.columns([8, 2])
    with right:
        if st.button(f"🔔 Notification  {unread_count}", key="job_notification_button", use_container_width=True, type="secondary"):
            st.session_state["notification_panel_open"] = not st.session_state["notification_panel_open"]
            if st.session_state["notification_panel_open"]:
                st.session_state["read_job_notification_ids"] = list({
                    *st.session_state["read_job_notification_ids"],
                    *(str(item.get("event_id", "")) for item in notifications),
                })
            st.rerun()

    if st.session_state["notification_panel_open"]:
        with right:
            st.markdown("""
            <style>
            .job-notification-card { background:linear-gradient(135deg,#A98BFF,#B79AFF); padding:12px; border-radius:18px; margin:8px 0 14px; }
            .job-notification-inner { background:#fff; border-radius:14px; padding:16px 18px; color:#253047; box-shadow:0 4px 14px rgba(30,35,55,.08); }
            .job-notification-title { font-size:20px; font-weight:750; margin-bottom:8px; }
            .job-notification-detail { font-size:15px; margin-bottom:6px; overflow-wrap:anywhere; }
            .job-notification-time { color:#747B8B; font-size:12px; }
            </style>
            """, unsafe_allow_html=True)
            if notifications:
                for item in reversed(notifications):
                    job_name = str(item.get("job_name", "New job"))
                    created_at = str(item.get("created_at", ""))
                    st.markdown(
                        '<div class="job-notification-card"><div class="job-notification-inner">'
                        '<div class="job-notification-title">✉️ New Job Received</div>'
                        f'<div class="job-notification-detail">New job added to Excel: {job_name}</div>'
                        f'<div class="job-notification-time">{created_at}</div>'
                        '</div></div>',
                        unsafe_allow_html=True,
                    )
            else:
                st.info("No new job notifications yet.")


def EmailAutomation():
    """Gmail polling status and safe manual test for email-to-Excel automation."""
    st.markdown("## Gmail Email-to-Excel Automation")
    st.write("Read incoming Gmail messages and add new jobs to the existing `Inprogress` sheet. Only the Job Name is required in the email.")
    status = gmail_setup_status()

    c1, c2, c3 = st.columns(3)
    c1.metric("Gmail OAuth credentials", "Found" if status["credentials_file"] else "Missing")
    c2.metric("Gmail authorization token", "Connected" if status["token_file"] else "Not connected")
    c3.metric("Excel workbook", "Found" if status["workbook_file"] else "Missing")

    with st.expander("Setup and monitoring settings", expanded=not status["token_file"]):
        st.write(f"**Polling interval:** every {status['poll_seconds']} seconds")
        st.write(f"**Gmail search query:** `{status['query']}`")
        st.write("**Data source:** email body only. Attachments are not read by this feature.")
        st.write("**Destination:** existing Inprogress worksheet. Existing workbook sheets and styles are retained; a backup should be made before enabling automation.")
        st.markdown("""
        1. Enable the Gmail API in Google Cloud and create an OAuth client of type **Desktop app**.
        2. Save the downloaded file as `credentials.json` in the same folder as `WAS.py`.
        3. Install the packages listed in `requirements-gmail.txt`.
        4. Run `python authorize_gmail.py` once in that folder and approve Gmail read-only access.
        5. Keep `workflow.xlsx` in that same folder and restart Streamlit.
        """)
        st.caption("Do not share credentials.json or token.json. They can grant access to your mailbox.")

    if status["token_file"] and status["workbook_file"]:
        started = start_background_poller()
        st.success("Automatic Gmail polling is enabled for this app process. The background service checks for new matching messages every 60 seconds." if started else "Automatic Gmail polling is already running in this app process.")
        # The live notification watcher is mounted globally below so alerts can appear
        # even when the user navigates away from Email Automation.
        if not hasattr(st, "fragment"):
            st.warning("Automatic popup alerts require a recent Streamlit version. Upgrade with: python -m pip install --upgrade streamlit")
    else:
        st.warning("Automatic processing has not started. Complete the Gmail authorization steps above first.")

    st.divider()
    st.markdown("### Test email processing now")
    st.write("Use this button to run an immediate scan and see what was imported. Emails are not marked as read.")
    if st.button("📩 Read New Emails & Update Excel", key="gmail_process_now", type="primary", use_container_width=True):
        try:
            with st.spinner("Checking Gmail and updating workflow.xlsx..."):
                results = process_new_emails()
            if results:
                st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)
                added_jobs = [
                    row.get("detail", "")
                    for row in results
                    if str(row.get("status", "")).lower() == "added"
                ]
                if added_jobs:
                    clear_workflow_cache()
                    added_names = [
                        detail.split("Added '", 1)[1].split("'", 1)[0]
                        for detail in added_jobs
                        if "Added '" in detail
                    ]
                    if added_names:
                        stored = st.session_state.setdefault("job_notifications", [])
                        for name in added_names:
                            event_id = f"manual|{name}|{datetime.now().isoformat(timespec='seconds')}"
                            stored.append({"event_id": event_id, "job_name": name,
                                           "created_at": datetime.now().isoformat(timespec="seconds")})
                        st.session_state["job_notifications"] = stored[-200:]
            else:
                st.info("No unprocessed matching emails were found.")
        except Exception as exc:
            st.error(f"Email processing failed: {exc}")

    st.markdown("### Expected email format")
    st.code("Job Name: Brochure Project", language="text")
    st.caption("Only a clearly labeled Job Name is required. Job ID, Client, Total Pages, Operator, Complexity, and Deadline do not need to be included. The job is added with status TBI and receipt time when those columns exist in Excel. Duplicate job names or IDs are skipped.")


page_details = {
    "Dashboard": ("Dashboard", "Monitor job status, workload, and active operations."),
    "Management Insights": ("Management Insights", "Executive KPIs, delivery trends, team productivity, and operational risks."),
    "Job Management": ("Job Management", "Create, assign, route to QC, and complete jobs."),
    "Edit Job Details": ("Edit Job Details", "Update existing job fields with a recorded change history."),
    "Job History & Audit Log": ("Job History & Audit Log", "Review current job records and logged changes."),
    "Deadline Monitoring": ("Deadline Monitoring", "Review due-soon, critical, and overdue jobs."),
    "Job Search": ("Job Search", "Search job records and review workflow details."),
    "Employee Workload": ("Employee Workload", "Review operator assignments and workload."),
    "Reports & Export": ("Reports & Excel Export", "Generate operational reports and export job data."),
    "Email Automation": ("Email Automation", "Read new Gmail job requests and update workflow.xlsx automatically."),
}
page_title, page_description = page_details[selected_page]

# Notification center stays in the app page flow; details appear only after clicking the button.
try:
    _gmail_status_for_notifications = gmail_setup_status()
    if _gmail_status_for_notifications["token_file"] and _gmail_status_for_notifications["workbook_file"]:
        start_background_poller()
        if hasattr(st, "fragment"):
            @st.fragment(run_every="5s")
            def _notification_center_fragment():
                _background_gmail_notification_watcher()
                _render_notification_center()
            _notification_center_fragment()
        else:
            _background_gmail_notification_watcher()
            _render_notification_center()
    else:
        _render_notification_center()
except Exception:
    _render_notification_center()

# Use a consistent main/assistant ratio; the assistant has its own white card and shadow.
if st.session_state.get("show_ai_panel", False):
    main_col, assistant_col = st.columns([2.05, 1.0], gap="large")
else:
    main_col = st.container()
    assistant_col = None

with main_col:
    st.markdown(
        f'<div class="app-heading"><h1>{page_title}</h1><p>{page_description}</p></div>',
        unsafe_allow_html=True,
    )
    if selected_page == "Dashboard":
        # Keep the Dashboard focused on the core workflow status view.
        # Management reports remain available from the separate Management Insights page.
        DashboardStatus()
    elif selected_page == "Management Insights":
        ManagementInsights()
    elif selected_page == "Edit Job Details":
        EditJobDetails()
    elif selected_page == "Job History & Audit Log":
        JobHistoryAuditLog()
    elif selected_page == "Job Management":
        st.subheader("Create New Job")
        NewJob()
        st.divider()
        st.subheader("Assign Job")
        AssignJob()
        st.divider()
        st.subheader("Move Job to QC")
        MoveJobToQC()
        st.divider()
        st.subheader("Complete Job")
        MoveJobToCompleted()
    elif selected_page == "Deadline Monitoring":
        DeadlineMonitoring()
    elif selected_page == "Job Search":
        JobSearch()
    elif selected_page == "Employee Workload":
        EmployeeWorkloadDashboard()
    elif selected_page == "Reports & Export":
        ReportsExport()
    elif selected_page == "Email Automation":
        EmailAutomation()

if assistant_col is not None:
  with assistant_col:
    with st.container(key="ai_assistant_scope"):
          st.markdown(
              '<div class="ai-panel-card">'
              '<div class="ai-panel-heading"><span class="ai-panel-spark">✦</span>AI Assistant</div>'
              '<div class="ai-panel-subtitle">Ask about jobs, workload, and deadlines</div>'
              '<div style="display:flex;align-items:flex-start;gap:10px;margin:18px 0 16px;">'
              '<div style="flex:0 0 38px;width:38px;height:38px;border-radius:12px;'
              'background:#FFB13B;color:#FFFFFF;display:flex;align-items:center;'
              'justify-content:center;font-size:20px;">♙</div>'
              '<div style="background:#F4F6F9;border:1px solid #E6EAF0;border-radius:12px;'
              'padding:11px 13px;color:#384152;font-size:14px;line-height:1.65;">'
              'Hello! I can search jobs, report status counts, review deadlines and workload, '
              'or prepare an assignment/QC/completion action for confirmation.'
              '</div></div>'
                    '</div>',
              unsafe_allow_html=True,
          )
          # Render the existing chat function without its duplicate SVG title and introductory caption.
          AIAssistantMode(show_header=False)



