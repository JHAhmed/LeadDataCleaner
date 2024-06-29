import openpyxl
from openpyxl.utils import column_index_from_string
import pandas as pd

from pathlib import Path
import streamlit as st
import os

st.set_page_config(page_title="Lead Data Cleaner")
st.title("Lead Data Cleaner")

if "file_path" not in st.session_state:
    st.session_state.file_path = ""

if "file_name" not in st.session_state:
    st.session_state.file_name = ""

if "file_processed" not in st.session_state:
    st.session_state.file_processed = False

if "output_file_path" not in st.session_state:
    st.session_state.output_file_path = False

if "output_file_name" not in st.session_state:
    st.session_state.output_file_name = False

def delete_file ():
    output_path = Path(st.session_state.output_file_path)
    if os.path.exists(output_path):
        os.remove(output_path)

    upload_path = Path(st.session_state.file_path)
    if os.path.exists(upload_path):
        os.remove(upload_path)

def remove_all(input_file, columns_to_keep):
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active

    columns_to_keep_indices = [column_index_from_string(col.upper()) for col in columns_to_keep]
    all_columns = list(range(1, sheet.max_column + 1))
    columns_to_remove = [col for col in all_columns if col not in columns_to_keep_indices]

    for col_index in sorted(columns_to_remove, reverse=True):
        sheet.delete_cols(col_index)

    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width

    # Save the modified workbook
    output_name = "Output - " + str(st.session_state.file_name).replace(".csv", ".xlsx")
    st.session_state.output_file_name = output_name

    output_path = Path("Data", "Output", output_name)
    st.session_state.output_file_path = output_path
    wb.save(output_path)

def check_file():
    file_name = st.session_state.file_name
    file_path = st.session_state.file_path
    if (file_name[-3::] == "csv"):
        df = pd.read_csv(st.session_state.file_path, encoding='utf-8')
        
        output_path = str(file_path).replace(".csv", ".xlsx")
        output_file = pd.ExcelWriter(file_name.replace(".csv", ".xlsx"), engine='xlsxwriter')
        df.to_excel(output_path, index=False)
        
        output_file.close()

def upload_docs():
    upload_dir = Path("Data", "Upload")
    output_dir = Path("Data", "Output")
    
    if not upload_dir.exists():
        upload_dir.mkdir(parents=True)
    
    if not output_dir.exists():
        output_dir.mkdir(parents=True)

    with st.spinner("Uploading..."):
        doc = st.session_state.uploaded_file
        file_path = upload_dir / doc.name
        with open(file_path, "wb") as f:
            f.write(doc.getbuffer())
        
        st.session_state.file_name = doc.name
        st.session_state.file_path = file_path
    
    st.success('Done!')

def process():
    check_file()
    # remove_all(st.session_state.file_path, )

    if lead_tool == "Outscraper":
        columns_to_remove = [1, 3, 5, 6, ]
        columns_to_keep = ["B", "D", "Z", "AR", "AS", "CF"]
        
    elif lead_tool == "SocLeads":
        columns_to_remove = [2, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 18, 19, 20, 21, 23, 24, 25, 26, 27]
        columns_to_keep = ["A", "C", "D", "P", "Q", "V"]
        # process_excel(Path("Data", st.session_state.file_name), columns_to_remove)
    # elif lead_tool == "Octoparse":    
        # columns_to_remove = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
    # elif lead_tool == "Telescope":
        # columns_to_remove = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23]
    else:
        columns_to_keep = []
    
    remove_all(st.session_state.file_path, columns_to_keep)
    st.session_state.file_processed = True

def create_upload () :  
    create_upload_button = st.file_uploader("Upload Files", type=["xlsx", "csv"], accept_multiple_files=False, key="uploaded_file",
                        help="Upload Excel file to clean up!", on_change=upload_docs, label_visibility="visible")

    return create_upload_button

file_uploaded = not len(st.session_state.file_name) > 1

upload_button = create_upload()
st.text(f"Current Document: {st.session_state.file_name}")

lead_tool = st.selectbox(
   "Choose Lead Generation Tool",
#    ("Outscraper", "SocLeads", "Octoparse", "Telescope"),
   ("Outscraper", ),
    key="lead_tool"
)

st.button("Process", type="secondary", on_click=process, disabled=file_uploaded)

if (st.session_state.file_processed):
    st.write("Processed!")
    with open(st.session_state.output_file_path, "rb") as file:
        btn = st.download_button(
                label="Download Output",
                data=file,
                on_click=delete_file,
                file_name=str(st.session_state.output_file_name),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )