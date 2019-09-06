import openpyxl, math
from openpyxl import load_workbook
import csv,sys

def CreateOffScore(ListExcel):    
    #Add Column OffScore
    HeaderRow = ListExcel[0]
     
    #Calculate OffScore
    Index_HB_RET = GetOffScoreParamPosition(HeaderRow)
    if Index_HB_RET[0] == 0 or Index_HB_RET[1] == 0: return False #Not calculate if missing position parameter
    else:
        ListExcel[0].append("OffScore")
        MaxRow = len(ListExcel)
        for idx in range(1,MaxRow): #Each Loop all row in sheet
            ThisRow = ListExcel[idx]
            Hb = ThisRow[Index_HB_RET[0]]
            Ret = ThisRow[Index_HB_RET[1]]
            OffScore = CalculateOffScore(Hb,Ret)
            #print(Hb,Ret,OffScore)
            ThisRow.append(OffScore)
    return True

def GetRowDataList(ListWorkSheet,idx):
    RawRow = ListWorkSheet[idx-1]
    HeaderRow = []
    for RawCell in RawRow:
        HeaderRow.append(RawCell.value)
    return HeaderRow

def GetOffScoreParamPosition(HeaderRow):
    Index_HB = 0
    Index_RET = 0
    Index_HB_RET = (Index_HB,Index_RET)
    try:
        Param_HB = "HGB(g/dL)"
        Param_RET = "RET%(%)"
        Index_HB = HeaderRow.index(Param_HB)
        Index_RET = HeaderRow.index(Param_RET)
        Index_HB_RET = (Index_HB,Index_RET)
    except Exception as e:
        print(e)
    #print("Index_HB_RET:", Index_HB_RET)
    return Index_HB_RET
    
def CalculateOffScore(Hb,Ret):     
    OffScore = '-'    
    try:
        Hb = float(Hb)
        Ret = float(Ret)
        OffScore = (10*Hb)-(60*math.sqrt(Ret)) #OFF-score = 10*Hb(g/dL)-60*square root(ret%)
        OffScore = ("%.2f" % OffScore)
    except Exception as e:
        print(e)
    return OffScore

def TextValue(Key):
    ValueReturn = ""
    TextValueDict = {}
    TextValueDict['XN-1000-1-A'] = 'XN-1000'
    if Key in TextValueDict.keys(): ValueReturn = TextValueDict[Key]
    else: ValueReturn = Key
    return ValueReturn
    
    
    
    
    
    
    
    
    
