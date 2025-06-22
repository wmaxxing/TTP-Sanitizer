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
    if st.button("Continue", key="continue"):
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

    # FORCE "# Students" to string for editability
    if "# Students" in outputFile_display.columns:
        outputFile_display["# Students"] = outputFile_display["# Students"].astype(str)
        
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
    col1, col2, col3, col4, _ = st.columns([0.13, 0.18, 0.15, 0.15, 1.75])

    with col1:
        # === TTPS ON BUTTON PRESS ===
        if st.button("TTPS", key="editTTP"):
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            processedData = extractionFunctions.dataTTPS(editedDataFile)
            st.session_state.ttpsData = processedData.copy()
            st.session_state.screen = "ttps"
            st.rerun()
    
    with col2:
    # === TRACKER ON BUTTON PRESS ===
        if st.button("TRACKER", key="editTracker"):
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            processedData = extractionFunctions.dataTracker(editedDataFile)
            st.session_state.trackerData = processedData.copy()
            st.session_state.screen = "tracker"
            st.rerun()
            
    with col3:
        # === CONTINUE ON BUTTON PRESS ===
        if st.button("ONE45", key="editOne45"):
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            processedData = extractionFunctions.dataOne45(editedDataFile)
            st.session_state.one45Data = processedData.copy()
            st.session_state.screen = "one45"
            st.rerun()
    
    with col4:
        # === RETURN TO DATA ENTRY === 
        if st.button("Return", key="editReturn"):
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            st.session_state.cleanedData = editedDataFile.copy()
            st.session_state.screen = "collectData"
            st.rerun()


        
    # === SCREEN: TTPS OUTPUT ===
if st.session_state.screen == "ttps":
    st.title("TTPS Data")
    st.dataframe(st.session_state.ttpsData)
        
    # === RETURN TO DATA EDITING === 
    if st.button("Return", key="editReturnTTPS"):
        st.session_state.screen = "edit"
        st.rerun()
        
# === SCREEN: INTERNAL TRACKER OUTPUT ===
if st.session_state.screen == "tracker":
    st.title("Internal Tracker Data")
    st.dataframe(st.session_state.trackerData)
    
    # === RETURN TO DATA EDITING === 
    if st.button("Return", key="editReturnTracker"):
        st.session_state.screen = "edit"
        st.rerun()


# === SCREEN: ONE45 OUTPUT ===
if st.session_state.screen == "one45":
    st.title("One45 Data")
    st.dataframe(st.session_state.one45Data)
        
    # === RETURN TO DATA EDITING === 
    if st.button("Return", key="editReturnOne45"):
        st.session_state.screen = "edit"
        st.rerun()




