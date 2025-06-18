import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import numbers
from datetime import datetime

buffer = 1
capture = 4
start = 0 

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
    
# Used to pull the data block from the original file  
def findBlocks(dataFile: pd.DataFrame):
    lastRow = dataFile.dropna(how='all').index.max() + 1
    dataNums = []
    #Loops through whole doc and finds check for boxes of content
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
            # SETS UP THE TEMP ROW OF DATA FROM THE SPREAD SHEET
        for l in range(3):
            temp.append(dataFile.iloc[k,l])
            
        if (isinstance(temp[0], datetime)):
            tempRowList[0] = pd.to_datetime(dataFile.iloc[k, 0]).strftime('%Y-%m-%d')
            # HANDLES THE TIME        
            timeHander(dataFile, temp, tempRowList, k)  
            # HANDLES THE STUDENTS
            studentHander(dataFile, k, tempRowList)
            # DUPLICATES ROWS DEPENDING ON MORNING AND AFTERNOON COVERAGE
            rowDupe(totalDataFrame, tempRowList)
        
        elif ((pd.isna(temp[0])) and (str(temp[1])[0] in tuple("0123456789"))):
            # PULLS IN THE PREVIOUS DATE FROM THE NEAREST ABOVE LINE
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
def studentHander(dataFile: pd.DataFrame, k: int, tempRowList: list):            # HANDLES THE STUDENTS
    studentList = []
    # CASE WHERE FIRST TILE IS FULL AND NAMES FOLLOW BELOW
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
    # CASE WHERE FIRST TILE IS EMPTY AND MUST BACKTRACK UP FOR ENTRIES
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

    # DUPLICATES ROWS DEPENDING ON MORNING AND AFTERNOON COVERAGE
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
    topRowList = [dataFile.iloc[y1,0], dataFile.iloc[y1,3], dataFile.iloc[y2,1], dataFile.iloc[y2,2], "", "", "", ""]
    totalDataFrame.append(topRowList)
    # (DATE, TIME, STUDENT NAMES, NUMBER OF STUDENTS)
    secondRowList = ["DATE", "TIME", "EXTRA INFO", "# OF STUDENTS", "S1", "S2", "S3", "S4"]
    totalDataFrame.append(secondRowList)
    
# Creats an empty row in excel   
def excelEmptyRow(dataList: pd.DataFrame, columns: list):
    # EMPTY ROW FOR EASE OF READ
    dataList.append(pd.DataFrame([["", "", "", "", "", "", "", ""]], columns=columns))
    
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
    # FUNCTION THAT PULLS IN THE TIME OF THE CLASS
    # PULLS THE FIRST TIME OUT OF THE SHEET
    if (str(temp[1]).startswith(tuple("0123456789"))):
        tempRowList[1] = timeOfSession(temp[1])
        # HANDLES BELOW ROW COMMENTS
        currNum = k + 1
        curr = dataFile.iloc[currNum,1]
        while ((not str(curr) in paymentLUT) and (not str(curr).startswith(tuple("0123456789")))):
            if ((not str(curr).startswith(tuple("0123456789"))) and (not pd.isna(curr))):
                tempRowList[2] += " " + str(curr)
            currNum += 1
            curr = dataFile.iloc[currNum, 1]
    # DEALS WITH EMPTY ROWS PULLING THE TIME BACK IN
    elif (pd.isna(temp[1])):
        curr = k - 1
        currCell = dataFile.iloc[curr, 1]
        while (pd.isna(currCell)):
            curr-=1
            currCell = dataFile.iloc[curr, 1]
        tempRowList[1] = timeOfSession(str(currCell))
    # DEALS WITH SPECIAL CASE INPUTS IN TIME COLUMN 
    else:
        if (str(temp[1] == "STAT")):
            return
        tempRowList[2] += str(temp[1]) 
    