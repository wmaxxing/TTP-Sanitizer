
import pandas as pd
import extraFunction
import streamlit as st
from datetime import datetime

filePath = "./testSheets/test sheet2.xlsx"
dataFile = pd.read_excel(filePath, header=None)

dataList = []
columns = ["Session Date", "Site", "Session Type", "Num Students", "S1", "S2", "S3", "S4"]

dataNums = extraFunction.findBlocks(dataFile)

for i in range(len(dataNums)):
    totalDataFrame = [] 
    y1 = dataNums[i][0]
    y2 = dataNums[i][1]
    extraFunction.excelSetUp(dataFile, totalDataFrame, y1, y2)
    # (DATA AGGREGATION FOR STUDENTS AND LESSON TIMES)
    y1+=2 #REORIENT OUR POINTER
    for k in range (y1, y2):
        tempRowList = ["", "", "", "", "", "", "", ""]
        temp = []
        extraFunction.handleRows(dataFile, temp, tempRowList, k, totalDataFrame)
            
    extraFunction.dataAccum(dataList, totalDataFrame, columns)
    extraFunction.excelEmptyRow(dataList, columns)
    
outputFile = pd.concat(dataList, ignore_index=True)

# === STREAMLIT INTERFACE FOR EDITING ===

# SETUP
st.set_page_config(layout="wide")
st.title("Sanitized Data")

# INITIALIZE SESSION STATE
if 'outputFile' not in st.session_state:
    # INDEX + COLUMNS
    outputFile = outputFile.reset_index(drop=True)
    outputFile.columns = ["Session Date", "Time", "Session Type", "# Students", "S1", "S2", "S3", "S4"]

    # FORCE TYPES FOR EDITING
    outputFile["# Students"] = outputFile["# Students"].astype(str)

    # STORE IN SESSION STATE
    st.session_state.outputFile = outputFile.copy()

# === DYNAMIC ROW INSERTION CONTROLS ===
st.subheader("Insert Rows Dynamically")

# USER CHOOSES WHERE TO INSERT
insert_at = st.number_input(
    "Insert new row at index:",
    min_value=1,
    max_value=len(st.session_state.outputFile) + 1,
    value=1
)
insert_at = insert_at - 1

# USER CHOOSES HOW MANY ROWS TO INSERT
num_rows_to_insert = st.number_input(
    "Number of rows to insert:",
    min_value=1,
    max_value=100,
    value=1
)

# INSERT ROW BUTTON
if st.button("Insert Empty Rows"):
    new_rows = pd.DataFrame(
        [[""] * len(st.session_state.outputFile.columns)] * num_rows_to_insert,
        columns=st.session_state.outputFile.columns
    )
    st.session_state.outputFile = pd.concat([
        st.session_state.outputFile.iloc[:insert_at],
        new_rows,
        st.session_state.outputFile.iloc[insert_at:]
    ], ignore_index=True)

    st.success(f"Inserted {num_rows_to_insert} row(s) at index {insert_at}.")

# === ADD ROW NUMBER COLUMN FOR DISPLAY ===

# MAKE A COPY OF DATAFRAME FOR DISPLAY
outputFile_display = st.session_state.outputFile.copy()

# ADD ROW COLUMN (1-BASED)
outputFile_display.insert(0, 'Row', outputFile_display.index + 1)

# === DATA EDITOR ===
edited_df = st.data_editor(
    outputFile_display,
    use_container_width=True,
    num_rows="dynamic",
    height=800
)

# === SAVE ON BUTTON PRESS ===
if st.button("Save Final Version"):
    # REMOVE ROW COLUMN BEFORE SAVING
    edited_df_no_row = edited_df.drop(columns=['Row'])

    # CONVERT Num Students BACK TO NUMBER
    edited_df_no_row["# Students"] = edited_df_no_row["# Students"].apply(
        lambda x: int(x) if str(x).strip().isdigit() else x
    )

    # SAVE TO EXCEL
    extraFunction.saveCleanExcel(edited_df_no_row, "block_1.xlsx")
    st.success("Final version saved as block_1.xlsx")


