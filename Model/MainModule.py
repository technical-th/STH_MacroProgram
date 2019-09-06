import sys
from Model import BuildDataDict,ManageExcel
 
def Runtime(InputPath):
    ReturnStatus = "Unknown Error"
    try:
        if(BuildDataDict.CheckExcelFormat(InputPath) == False): return "Error: Wrong File Format"
        print('=== Begin ==============================================================')
        print("InputPath:", InputPath)
        DataDict = BuildDataDict.BuildDict(InputPath)
        print("DataDict:", DataDict)
        ManageExcel.MainExcel(DataDict)
        ReturnStatus = "Conversion Successful."
        print('=== Done ===============================================================')
    except Exception as e:
        ReturnStatus =  str(e)
    return ReturnStatus
 
InputPath = 'C:/Users/sjg/Desktop/STH; Macro program XN-1000/INPUT (05032019).csv'
#InputPath = 'C:/Users/sjg/Desktop/STH; XN-1000 Filter Excel/Source 4/INPUT (24042019_Corrupt).csv'
# Status = Runtime(InputPath)
# print(Status)


