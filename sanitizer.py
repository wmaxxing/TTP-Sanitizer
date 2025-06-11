
import pandas as pd
import extraFunction
from datetime import datetime

filePath = "./testSheets/test sheet2.xlsx"
dataFile = pd.read_excel(filePath, header=None)

dataList = []
columns = ["0", "1", "2", "3", "4"]

dataNums = extraFunction.findBlocks(dataFile)

for i in range(len(dataNums)):
    totalDataFrame = [] 
    y1 = dataNums[i][0]
    y2 = dataNums[i][1]
    extraFunction.excelSetUp(dataFile, totalDataFrame, y1, y2)
    # (DATA AGGREGATION FOR STUDENTS AND LESSON TIMES)
    y1+=2 #REORIENT OUR POINTER
    for k in range (y1, y2):
        tempRowList = [dataFile.iloc[k,0], "", "", "", ""]
        temp = []
        extraFunction.handleRows(dataFile, temp, tempRowList, k, totalDataFrame)
            
    extraFunction.dataAccum(dataList, totalDataFrame, columns)
    extraFunction.excelEmptyRow(dataList, columns)
    
outputFile = pd.concat(dataList, ignore_index=True)

extraFunction.saveCleanExcel(outputFile, "block_1.xlsx")
