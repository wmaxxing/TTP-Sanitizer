import pandas as pd
import extractionFunctions
import streamlit as st
from datetime import datetime

filePath = ""
outputFile = ""
rotation = ""
startDate = ""
endDate = ""

# === STREAMLIT INTERFACE FOR EDITING ===

# SETUP
st.set_page_config(layout="wide")

# === SET SCREEN STATE ===
if "screen" not in st.session_state:
    st.session_state.screen = "collectData"

import datetime

# === SCREEN: DATA COLLECT ===
if st.session_state.screen == "collectData":
    st.title("Rotation Information")

    # LOAD PREVIOUS VALUES IF AVAILABLE
    filePath = st.text_input("File Path", value=st.session_state.get("filePath", ""), placeholder="C:/path/to/file.xlsx")
    rotation = st.text_input("Rotation Name", value=st.session_state.get("rotation", ""), placeholder="A1")
    startDate = st.date_input("Start Date", value=st.session_state.get("startDate", datetime.date.today()))
    endDate = st.date_input("End Date", value=st.session_state.get("endDate", datetime.date.today()))

    # === CONTINUE ON BUTTON PRESS ===
    if st.button("Continue"):
        if not filePath or not rotation:
            st.warning("Please fill out all required fields.")
        else:
            try:
                outputFile = extractionFunctions.fileExtractor(filePath.strip('"'))
                st.session_state.outputFile = outputFile

                # Save form data for future reuse
                st.session_state.filePath = filePath
                st.session_state.rotation = rotation
                st.session_state.startDate = startDate
                st.session_state.endDate = endDate

                st.session_state.screen = "edit"
                st.rerun()

            except Exception as e:
                st.error("Could not load or process the selected file.")
                st.exception(e)


    

# === SCREEN: EDIT DATA ===
if st.session_state.screen == "edit":
    st.title("Sanitized Data Editor")
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
    st.subheader("Insert Rows")

    # USER CHOOSES WHERE TO INSERT
    insertAt = st.number_input(
        "Insert new row at index:",
        min_value=1,
        max_value=len(st.session_state.outputFile) + 1,
        value=1
    )
    insertAt = insertAt - 1

    # USER CHOOSES HOW MANY ROWS TO INSERT
    numRowsToInsert = st.number_input(
        "Number of rows to insert:",
        min_value=1,
        max_value=100,
        value=1
    )

    # INSERT ROW BUTTON
    if st.button("Insert Empty Rows"):
        newRows = pd.DataFrame(
            [[""] * len(st.session_state.outputFile.columns)] * numRowsToInsert,
            columns=st.session_state.outputFile.columns
        )
        st.session_state.outputFile = pd.concat([
            st.session_state.outputFile.iloc[:insertAt],
            newRows,
            st.session_state.outputFile.iloc[insertAt:]
        ], ignore_index=True)

        st.success(f"Inserted {numRowsToInsert} row(s) at index {insertAt}.")

    # === ADD ROW NUMBER COLUMN FOR DISPLAY ===

    # MAKE A COPY OF DATAFRAME FOR DISPLAY
    outputFile_display = st.session_state.outputFile.copy()

    # ADD ROW COLUMN (1-BASED)
    outputFile_display.insert(0, 'Row', outputFile_display.index + 1)

    # === DATA EDITOR ===
    editedDataFile = st.data_editor(
        outputFile_display,
        use_container_width=True,
        num_rows="dynamic",
        height=800
    )

    # === BUTTONS SIDE BY SIDE ===
    col1, col2, _ = st.columns([0.15, 0.15, 1.75])

    # === UNIVERSAL CLEANING FUNCTION ===
    def cleanEditedData(df):
        # DROP ROW COLUMN IF EXISTS
        if "Row" in df.columns:
            df = df.drop(columns=['Row'])

        # SAFE CONVERT "# Students"
        if "# Students" in df.columns:
            df["# Students"] = df["# Students"].apply(
                lambda x: int(x) if str(x).strip().isdigit() else x
            )

        return df

    with col1:
        # === CONTINUE ON BUTTON PRESS ===
        if st.button("Continue", key="edit_continue"):
            editedDataFile = cleanEditedData(editedDataFile)
            st.session_state.cleanedData = editedDataFile.copy()
            st.session_state.screen = "processed"
            st.rerun()

    with col2:
        # === RETURN TO DATA ENTRY === 
        if st.button("Return", key="edit_return"):
            editedDataFile = cleanEditedData(editedDataFile)
            st.session_state.cleanedData = editedDataFile.copy()
            st.session_state.screen = "collectData"
            st.rerun()


        
# === SCREEN: CLEANED OUTPUT ===
# elif st.session_state.screen == "processed":
#     st.title("Processed Data")

#     st.dataframe(st.session_state.cleanedData)

#     # EXPORT BUTTON
#     if st.button("Export to Excel"):
#         extraFunction.saveCleanExcel(st.session_state.cleanedData, "block_1.xlsx")
#         st.success("Final version exported as block_1.xlsx")
#         st.rerun()

#     # BACK BUTTON
#     if st.button("Back to Editor"):
#         st.session_state.screen = "edit"
#         st.rerun()
        



