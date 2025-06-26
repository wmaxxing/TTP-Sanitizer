import pandas as pd
import extractionFunctions
import streamlit as st
from datetime import datetime

filePath = ""
outputFile = ""
rotation = ""
startDate = ""
endDate = ""

# === STREAMLIT INTERFACE ===

# Setup
st.set_page_config(layout="wide")

# Set screen state
if "screen" not in st.session_state:
    st.session_state.screen = "collectData"

import datetime

# Screen: Data Collection
if st.session_state.screen == "collectData":
    st.title("Rotation Information")

    # Load previous values if available
    filePath = st.text_input("File Path", value=st.session_state.get("filePath", ""), placeholder="C:/path/to/file.xlsx")
    rotation = st.text_input("Rotation Name", value=st.session_state.get("rotation", ""), placeholder="A1")
    startDate = st.date_input("Start Date", value=st.session_state.get("startDate", datetime.date.today()))
    endDate = st.date_input("End Date", value=st.session_state.get("endDate", datetime.date.today()))

    # Continue to Editing on button press
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


    

# Screen: Edit Data
if st.session_state.screen == "edit":
    st.title("Sanitized Data Editor")

    if 'outputFile' not in st.session_state:
        st.error("Output file not found. Please return to the previous screen and re-load the file.")
        st.session_state.screen = "collectData"
        st.rerun()

    outputFile = st.session_state.outputFile.copy()

    st.subheader("Insert Rows")

    insertAt = st.number_input(
        "Insert new row at index:",
        min_value=0,
        max_value=len(outputFile) + 1,
        value=0
    )

    numRowsToInsert = st.number_input(
        "Number of rows to insert:",
        min_value=1,
        max_value=100,
        value=1
    )

    # Insert empty rows on button press for ease of editing
    if st.button("Insert Empty Rows"):
        newRows = pd.DataFrame(
            [[""] * len(outputFile.columns)] * numRowsToInsert,
            columns=outputFile.columns
        )
        outputFile = pd.concat([
            outputFile.iloc[:insertAt],
            newRows,
            outputFile.iloc[insertAt:]
        ], ignore_index=True)

        st.session_state.outputFile = outputFile.copy()
        st.success(f"Inserted {numRowsToInsert} row(s) at index {insertAt}.")

    # Display with clean rown numbers using index
    outputFile_display = outputFile.reset_index(drop=True)

    if "# Students" in outputFile_display.columns:
        outputFile_display["# Students"] = outputFile_display["# Students"].astype(str)

    editedDataFile = st.data_editor(
        outputFile_display,
        use_container_width=True,
        num_rows="dynamic",
        height=800,
        hide_index=False
    )

    col1, col2, col3, col4, _ = st.columns([0.13, 0.18, 0.15, 0.15, 1.75])

    with col1:
        # Move to TTPS screen on button press
        if st.button("TTPS", key="editTTP"):
            
            # Save latest edits to session state
            st.session_state.outputFile = editedDataFile.copy()
    
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            processedData = extractionFunctions.dataTTPS(editedDataFile)
            st.session_state.ttpsData = processedData.copy()
            st.session_state.screen = "ttps"
            st.rerun()

    with col2:
        # Move to Tracker screen on button press
        if st.button("TRACKER", key="editTracker"):
            
            # Save latest edits to session state
            st.session_state.outputFile = editedDataFile.copy()
            
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            processedData = extractionFunctions.dataTracker(editedDataFile)
            st.session_state.trackerData = processedData.copy()
            st.session_state.screen = "tracker"
            st.rerun()

    with col3:
        # Move to One45 screen on button press
        if st.button("ONE45", key="editOne45"):
            
            # Save latest edits to session state
            st.session_state.outputFile = editedDataFile.copy()
            
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            processedData = extractionFunctions.dataOne45(editedDataFile)
            st.session_state.one45Data = processedData.copy()
            st.session_state.screen = "one45"
            st.rerun()

    with col4:
        # Return to Data Collection screen on button press
        if st.button("Return", key="editReturn"):
            editedDataFile = extractionFunctions.cleanEditedData(editedDataFile)
            st.session_state.cleanedData = editedDataFile.copy()
            st.session_state.screen = "collectData"
            st.rerun()

        
# Screen: TTPS output
if st.session_state.screen == "ttps":
    st.title("TTPS Data")
    st.dataframe(st.session_state.ttpsData)
        
    # Return to Editing on button press
    if st.button("Return", key="editReturnTTPS"):
        st.session_state.screen = "edit"
        st.rerun()
        
# Screen: Internal Tracker output
if st.session_state.screen == "tracker":
    st.title("Internal Tracker Data")
    st.dataframe(st.session_state.trackerData)
    
    # Return to Editing on button press
    if st.button("Return", key="editReturnTracker"):
        st.session_state.screen = "edit"
        st.rerun()


# Screen: One45 output
if st.session_state.screen == "one45":
    st.title("One45 Data")
    st.dataframe(st.session_state.one45Data)
        
    # Return to Editing on button press
    if st.button("Return", key="editReturnOne45"):
        st.session_state.screen = "edit"
        st.rerun()




