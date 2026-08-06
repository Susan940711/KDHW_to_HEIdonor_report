import tempfile
from pathlib import Path

import streamlit as st

from transform_indicator_to_semester_report import build_semester_report


st.set_page_config(page_title="Indicator to Semester Report", page_icon="📊", layout="centered")
st.title("Indicator to Semester Report")
st.write("Upload the quarterly compile workbook and generate a semester report sheet in the browser.")

uploaded_file = st.file_uploader("Choose an Excel workbook", type=["xlsx"], help="Select the workbook that contains the Indicator sheet")
source_sheet = st.text_input("Source sheet", value="Indicator")
target_sheet = st.text_input("Target sheet", value="Semester_Report")

if uploaded_file is not None:
    if st.button("Generate report"):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir_path = Path(tmpdir)
                input_path = tmpdir_path / uploaded_file.name
                output_path = tmpdir_path / "semester_report.xlsx"
                input_path.write_bytes(uploaded_file.getvalue())

                build_semester_report(
                    input_path=input_path,
                    output_path=output_path,
                    source_sheet=source_sheet,
                    target_sheet=target_sheet,
                )

                output_bytes = output_path.read_bytes()
                st.success("Report generated successfully.")
                st.download_button(
                    label="Download generated workbook",
                    data=output_bytes,
                    file_name=f"{uploaded_file.name.replace('.xlsx', '')}_semester_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as exc:
            st.error(f"Generation failed: {exc}")
else:
    st.info("Please upload an Excel workbook to begin.")
