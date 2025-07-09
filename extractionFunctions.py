import pandas as pd
import string
from openpyxl import load_workbook
from openpyxl.styles import numbers
from datetime import datetime

buffer = 1
capture = 4
start = 0 
s1Col = 4
s4Col = 8
offset = 2

paymentLUT = ["FFS", "GFT", "PS", "CSC", "SESSIONAL", "CHES FELLOW", "AIP", "CF", "PS", "SHA", "START TIME", "TSF"]

# Used to clean the output into the excel file and make sure that all formatting is correct
def saveCleanExcel(df: pd.DataFrame, filename: str):
    if df.empty:
        raise ValueError("DataFrame is empty. Nothing to save.")

    df.to_excel(filename, index=False, header=False)

    wb = load_workbook(filename)
    ws = wb.active

    for colCells in ws.columns:
        maxLen = 0
        colLetter = colCells[0].column_letter

        for cell in colCells:
            if isinstance(cell.value, datetime):
                cell.number_format = numbers.FORMAT_DATE_YYYYMMDD2
            if cell.value is not None:
                maxLen = max(maxLen, len(str(cell.value)))

        ws.column_dimensions[colLetter].width = maxLen + 3

    wb.save(filename)
    print(f"Saved to: {filename}")
    
# ==== DATA EXTRACTION FUNCTIONS ====

# Used to pull the data block from the original file  
def findBlocks(dataFile: pd.DataFrame):
    lastRow = dataFile.dropna(how='all').index.max() + 1
    dataNums = []
    # Loops through whole doc and finds check for boxes of content
    for i in range(lastRow):
        testForDate = str(dataFile.iloc[i, start])
        # If the keyword date is found begins the box searching process
        if (testForDate.strip().upper() == "DATE"):
            testForStudent = dataFile.iloc[i+1, 2]
            testForStat = str(dataFile.iloc[i+1, 1])
            # If the keyword date is found but the box is content empty skips the data box
            if ((not testForStat == "STAT") and (not isinstance(testForStudent, str))):
                continue
            # Once the box has been indentified as good finds the end of the data box
            for k in range(i + 1, lastRow):
                testForEnd1 = dataFile.iloc[k, start]
                testForEnd2 = str(dataFile.iloc[k, start + 1])
                # Records the start and end of the data box into a list where it can be used upon return
                if (isinstance(testForEnd1, str)):
                    tempData = [i-1, k] 
                    dataNums.append(tempData)
                    break
                if (pd.isna(testForEnd1) and (testForEnd2 in paymentLUT)):
                    tempData = [i-1, k] 
                    dataNums.append(tempData)
                    break
    return dataNums

# Used to handle all the pulling of data from the time columns of the original document
def handleRows (dataFile: pd.DataFrame, temp: list, tempRowList: list, k: int, totalDataFrame: list):
            # Sets up the temp row of data from the spread sheet
        for l in range(3):
            temp.append(dataFile.iloc[k,l])
            
        if (isinstance(temp[0], datetime)):
            tempRowList[0] = pd.to_datetime(dataFile.iloc[k, 0]).strftime('%Y-%m-%d')
            # Handles the time      
            timeHander(dataFile, temp, tempRowList, k)  
            # Handles the students
            studentHander(dataFile, k, tempRowList)
            # Duplicates rows depending on morning and afternoon coverage
            rowDupe(totalDataFrame, tempRowList)
        
        elif ((pd.isna(temp[0])) and (str(temp[1])[0] in tuple("0123456789"))):
            # Pulls in the previous date from the nearest above line
            curr = k - 1
            currCell = dataFile.iloc[curr, 0]
            while (pd.isna(currCell)):
                curr-=1
                currCell = dataFile.iloc[curr, 0]
            tempRowList[0] = pd.to_datetime(currCell).strftime('%Y-%m-%d')
            
            timeHander(dataFile, temp, tempRowList, k)
            studentHander(dataFile, k, tempRowList)
            rowDupe(totalDataFrame, tempRowList)

# Used to pull all of the data relating to students out of the spread sheets, returns a name list of students
def studentHander(dataFile: pd.DataFrame, k: int, tempRowList: list):
    studentList = []
    # Case where first tile is full and names follow below
    if (not pd.isna(dataFile.iloc[k, 2])):
        studentList.append(str(dataFile.iloc[k, 2]))
        curr = k + 1
        currCell = dataFile.iloc[curr, 0]
        while (pd.isna(currCell)):
            if (not pd.isna(dataFile.iloc[curr, 1]) and (str(dataFile.iloc[curr, 1])[0] in tuple("0123456789"))):
                break
            if (not pd.isna(dataFile.iloc[curr, 2])):
                studentList.append(str(dataFile.iloc[curr, 2]))
            curr+=1
            currCell = dataFile.iloc[curr, 0]
    # Case where first tile is empty and must backtrack up for entries
    elif (pd.isna(dataFile.iloc[k, 2])): 
        curr = k - 1
        currCell = dataFile.iloc[curr, 0]
        while (pd.isna(currCell)):
            curr-=1
            currCell = dataFile.iloc[curr, 0]
            if (not pd.isna(dataFile.iloc[curr, 2])):
                studentList.append(str(dataFile.iloc[curr, 2]))
        curr = k - 1
        studentList.append(str(dataFile.iloc[curr, 2]))
    

    for i in range(0, len(studentList)):
       tempRowList[4 + i] = studentList[i]
    tempRowList[3] = len(studentList)
                    
# Used to duplicate rows when a "BOTH" type session is encountered
def rowDupe(totalDataFrame: list, tempRowList: list):

    # Duplicates rows depending on morning and afternoon coverage
    if (tempRowList[1] == "Both"):
        tempRowList[1] = "Morning"
        totalDataFrame.append(tempRowList)
        dupRowList = [tempRowList[0], "Afternoon", tempRowList[2], tempRowList[3], tempRowList[4], tempRowList[5] , tempRowList[6], tempRowList[7]]
        totalDataFrame.append(dupRowList)
    else:
        totalDataFrame.append(tempRowList)
    
# Used to set up the document and orient basic requried info
def excelSetUp(dataFile: pd.DataFrame, totalDataFrame: list, y1: int, y2: int):
    # (DOCTOR NAME, SPECIALTY, TTP TYPE, SPECIAL INFO)
    if (str(dataFile.iloc[y1, 0]) == "PRECEPTOR TO BE DECIDED"):
        topRowList = [dataFile.iloc[y1,0], dataFile.iloc[y1,3], "*TBD*", dataFile.iloc[y2,2], "", "", "", ""]
        totalDataFrame.append(topRowList)
    else:
        topRowList = [dataFile.iloc[y1,0], dataFile.iloc[y1,3], dataFile.iloc[y2,1], dataFile.iloc[y2,2], "", "", "", ""]
        totalDataFrame.append(topRowList)
    # (DATE, TIME, STUDENT NAMES, NUMBER OF STUDENTS)
    secondRowList = ["DATE", "TIME", "EXTRA INFO", "# OF STUDENTS", "S1", "S2", "S3", "S4"]
    totalDataFrame.append(secondRowList)
    
# Creats an empty row in excel   
def emptyRow(dataList: pd.DataFrame, columns: list):
    # Empty row for ease of reading
    testList = []
    for i in range(len(columns)):
        testList.append("")
    dataList.append(pd.DataFrame([testList], columns=columns))
    
# Used to accumulate the data from a cycle of a block
def dataAccum(dataList: pd.DataFrame, totalDataFrame: list, columns:list):
    collectedData = pd.DataFrame(totalDataFrame, columns=columns)
    dataList.append(collectedData)
    
# Used to clevely determine if a session is morning afternoon or both depending on its time
def timeOfSession(timeOfSession: str):
    timeCutOff = 420
    lon = timeExtractor(timeOfSession)
    if ((lon[0] < 1200) and (lon[1] - lon[0] <= timeCutOff)):
        return "Morning"
    elif (lon[1] - lon[0] > timeCutOff):
        return "Both" 
    else: 
        return "Afternoon"
    
    
# Used to remove the interger values of a time from a string
def timeExtractor(timeOfSession: str):
    #Extracting the number from the incoming string
    noOne = ""
    noTwo = ""
    currIndex = 0
    for i in range(len(timeOfSession)):
        if (timeOfSession[i] in tuple("0123456789")):
            noOne += timeOfSession[i]
        elif (timeOfSession[i] == "-"):
            currIndex = i + 1
            break
        else:
            continue
    for k in range(currIndex, len(timeOfSession)):
        if (timeOfSession[k] in tuple("0123456789")):
            noTwo += timeOfSession[k]
        else:
            continue
    return [int(noOne), int(noTwo)]   

# Used to extract times and pick day time for sessions
def timeHander(dataFile: pd.DataFrame, temp: list, tempRowList: list, k: int):
    # Pulls the first time out of the sheet
    if (str(temp[1]).startswith(tuple("0123456789"))):
        tempRowList[1] = timeOfSession(temp[1])
        # Handles below row comments
        currNum = k + 1
        curr = dataFile.iloc[currNum,1]
        while ((not str(curr) in paymentLUT) and (not str(curr).startswith(tuple("0123456789")))):
            if ((not str(curr).startswith(tuple("0123456789"))) and (not pd.isna(curr))):
                tempRowList[2] += str(curr)
            currNum += 1
            curr = dataFile.iloc[currNum, 1]
    # Deals with empty rows pulling the time back in
    elif (pd.isna(temp[1])):
        curr = k - 1
        currCell = dataFile.iloc[curr, 1]
        while (pd.isna(currCell)):
            curr-=1
            currCell = dataFile.iloc[curr, 1]
        tempRowList[1] = timeOfSession(str(currCell))
    # Deals with special case inputs in the time column
    else:
        if (str(temp[1] == "STAT")):
            return
        tempRowList[2] += str(temp[1]) 

# Main function that runs the file extraction program
def fileExtractor(filePath: str):
    try:
        dataFile = pd.read_excel(filePath, header=None)

        dataList = []
        columns = ["Session Date", "Site", "Session Type", "# Students", "S1", "S2", "S3", "S4"]

        dataNums = findBlocks(dataFile)

        for i in range(len(dataNums)):
            totalDataFrame = [] 
            y1 = dataNums[i][0]
            y2 = dataNums[i][1]
            excelSetUp(dataFile, totalDataFrame, y1, y2)
            # (DATA AGGREGATION FOR STUDENTS AND LESSON TIMES)
            y1+=2 #REORIENT OUR POINTER
            for k in range (y1, y2):
                tempRowList = ["", "", "", "", "", "", "", ""]
                temp = []
                handleRows(dataFile, temp, tempRowList, k, totalDataFrame)
                    
            dataAccum(dataList, totalDataFrame, columns)
            emptyRow(dataList, columns)
            
        outputFile = pd.concat(dataList, ignore_index=True)
        return outputFile
    except Exception as e:
        print(f"[fileExtractor] Error processing file: {e}")
        raise 
    
# ==== END DATA FORMATTING FUNCTIONS ==== 

# UNIVERSAL CLEANING FUNCTION
def cleanEditedData(df):
    # Drop row column if exists
    if "Row" in df.columns:
        df = df.drop(columns=['Row'])

    # SAFE CONVERT "# Students"
    if "# Students" in df.columns:
        df["# Students"] = df["# Students"].apply(
            lambda x: int(x) if str(x).strip().isdigit() else x
        )

    return df

# Organizes the data in the format needed for TTPS
def dataTTPS(dataFile: pd.DataFrame, startDate: str, endDate: str, location: str, rotation: str):
    # List of all TTPS enterable data
    loTTPS = []
    result = pd.DataFrame([])
    lastRow = dataFile.dropna(how='all').index.max() + 1
    
    # Pulls from the processed data and orgainzes the loTTPS list
    for i in range(lastRow):
        testCell = dataFile.iloc[i, 0]
        if (str(dataFile.iloc[i, 3]) == "* Cannot Input into TTP *"):
            continue
        if ((not testCell == "DATE") and (str(testCell).startswith(tuple(string.ascii_letters)))):
            preceptorName = str(testCell)
            recievingFunction = str(dataFile.iloc[i, 1])
            oneLearner = [startDate, location, "1", rotation + " | " + "Student: ", recievingFunction, endDate, 0, preceptorName, []]
            twoPlusLearners = [startDate, location, "2+", rotation + " | " + " Student: ", recievingFunction, endDate, 0, preceptorName, []]
            index = i + offset
            while index < lastRow:
                if (str(dataFile.iloc[index, 0]) == ""):
                    break
                if (str(dataFile.iloc[index, 2]) in ["Teaching Session", "End of Rotation"]):
                    index += 1
                    continue
                if (dataFile.iloc[index, 3] == 1):
                    oneLearner[6] += 1
                    for k in range (s1Col, s4Col):
                        tempName = str(dataFile.iloc[index, k])
                        if (not tempName in oneLearner[3]):
                            oneLearner[8].append(tempName)
                            oneLearner[3] += " " + str(dataFile.iloc[index, k])
                if (dataFile.iloc[index, 3] > 1):
                    twoPlusLearners[6] += 1
                    for k in range (s1Col, s4Col):
                        tempName = str(dataFile.iloc[index, k])
                        if (not tempName in twoPlusLearners[3]):
                            twoPlusLearners[8].append(tempName)
                            twoPlusLearners[3] += " " + str(dataFile.iloc[index, k])
                index += 1
            if (oneLearner[6] > 0):
                loTTPS.append(oneLearner)
            if (twoPlusLearners[6] > 0):
                loTTPS.append(twoPlusLearners)
            # Displays the relevant info onto the screen when the TTPS Button is pressed
            columns = ["Start Date", "Location", "# Learners", "Comments", "Receiving Function", "End Date", "# Sessions", "Preceptor"]
            # Processing loTTPS
            for row in loTTPS:
                if (len(row) == 9):
                    row.pop() 

            result = pd.DataFrame(loTTPS, columns=columns)
    return result
    
# Organizes the data in the format needed for the Internal Tracker
def dataTracker(dataFile: pd.DataFrame, academicYear: str, rotation: str, location: str):
    # List of tracks where each entry is a row in the internal tracker
    loTracks = []
    result = pd.DataFrame([])
    lastRow = dataFile.dropna(how='all').index.max() + 1
    
    # Pulls from the processed data and orgainzes the loTracks list
    for i in range(lastRow):
        testCell = dataFile.iloc[i, 0]
        if ((not testCell == "DATE") and (str(testCell).startswith(tuple(string.ascii_letters)))):
            preceptorName = str(testCell)
            paymentType = str(dataFile.iloc[i, 2])
            index = i + offset
            while index < lastRow:
                singleTrack = [academicYear, rotation, location, "", "", preceptorName, "", paymentType, ""]
                if (str(dataFile.iloc[index, 0]) == ""):
                    break
                singleTrack[3] = dataFile.iloc[index, 0]
                singleTrack[4] = dataFile.iloc[index, 1]
                singleTrack[6] = dataFile.iloc[index, 3]
                singleTrack[8] = dataFile.iloc[index, 2]
                loTracks.append(singleTrack)
                index += 1
    
    # Displays the relevant info onto the screen when the Tracker Button is pressed
    columns = ["Academic Year", "Rotation", "Location", "Date", "Time Block", "Preceptor", "# Learners", "Payment Type", "Notes"]
    # Processing loTracks
    result = pd.DataFrame(loTracks, columns=columns)
    return result
    
# Organizes the data in the format needed for One45
def dataOne45(dataFile: pd.DataFrame):
    # List of preceptors and the there respective students
    preceptorsToStudents = []
    lastRow = dataFile.dropna(how='all').index.max() + 1
    
    # Pulls from the processed data and orgainzes the preceptorsToStudents list
    for i in range(lastRow):
        testCell = dataFile.iloc[i, 0]
        if ((not testCell == "DATE") and (str(testCell).startswith(tuple(string.ascii_letters)))):
            tempData = [str(testCell), []]
            index = i + offset
            while index < lastRow:
                if (str(dataFile.iloc[index, 0]) == ""):
                    break
                for k in range (s1Col, s4Col):
                    tempName = str(dataFile.iloc[index, k])
                    if (not tempName in tempData[1] and not tempName == ""):
                        (tempData[1]).append(tempName)
                index += 1
            preceptorsToStudents.append(tempData)
    
    # Displays the relevant info onto the screen when the One45 button is pressed
    columns = ["Preceptors", "---->", "Student 1", "Student 2", "Student 3", "Student 4"]
    dataCollection = []
    # Processing Preceptors -> Students
    for i in range(len(preceptorsToStudents)):
        tempList = []
        tempList.append(preceptorsToStudents[i][0])
        tempList.append("---->")
        preceptorsToStudents[i][1] = sorted(preceptorsToStudents[i][1])
        for k in range(len(preceptorsToStudents[i][1])):
            tempList.append(preceptorsToStudents[i][1][k])
        if (len(tempList) < 6):
            for t in range(6 - len(tempList)):
                tempList.append("")
        dataCollection.append(pd.DataFrame([tempList], columns=columns))
    
    return pd.concat(dataCollection)