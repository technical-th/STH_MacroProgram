import sys,csv
from time import gmtime, strftime

def BuildDict(InputPathFile):
    print('=== BuildDict ===========================================================')
    DataDict = {}
    DataDict['InputPathFile'] = InputPathFile
    DataDict['OutputPathFile'] = GetOutPutPath(InputPathFile)
    return DataDict

def GetOutPutPath(InputPath):
    Timenow = strftime("%Y%m%d_%H.%M.%S")
    OutputPath = InputPath[:InputPath.rindex('/')+1] + "XN-1000_Output_"+Timenow+".xlsx"
    return OutputPath

def CheckExcelFormat(InputPathFile):
    with open(InputPathFile) as f:
        reader = csv.reader(f, delimiter = ',')
        for Row in reader:
            Header = Row[0]
            #print("Header:",Row[0]);sys.exit()
            if Header == 'Do not use items enclosed in [ ] for reports.': return True
            else: return False
            break
        
        
        
        