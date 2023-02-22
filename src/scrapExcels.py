import json
import pandas as pd
import os
from utils import get_current_path

#read json file
configxlsxPath=os.path.join(get_current_path(),"config.xlsx")
indexColumnsPathJson=os.path.join(get_current_path(),"src","target","indexColumnsConfig.json")
kwordsRowLimitsPathJson=os.path.join(get_current_path(),"src","target","kwordsRowLimitsConfig.json")

dfC=pd.read_excel(configxlsxPath,sheet_name="columnas")
#print(df)
#conver the df into collection of dictionaries
dataColumns=dfC.values.tolist()
columnsDict = {}
for d in dataColumns:
    if d[0] not in columnsDict:
        columnsDict[d[0]] = {}
    if d[1] not in columnsDict[d[0]]:
        columnsDict[d[0]][d[1]] = {}
    if d[2] not in columnsDict[d[0]][d[1]]:
        columnsDict[d[0]][d[1]][d[2]] = {}
    if d[3] not in columnsDict[d[0]][d[1]][d[2]]:
        columnsDict[d[0]][d[1]][d[2]][d[3]] = {}
    if d[4] not in columnsDict[d[0]][d[1]][d[2]][d[3]]:
        columnsDict[d[0]][d[1]][d[2]][d[3]][d[4]] = {}
    columnsDict[d[0]][d[1]][d[2]][d[3]][d[4]][d[5]] = d[6]
with open(indexColumnsPathJson, 'w') as outfile:
    json.dump(columnsDict, outfile,indent=4)


dfKwords=pd.read_excel(configxlsxPath,sheet_name="kwords")
dataKeywords=dfKwords.values.tolist()
kwordsDict = {}
for d in dataKeywords:
    if d[0] not in kwordsDict:
        kwordsDict[d[0]] = {}
    if d[1] not in kwordsDict[d[0]]:
        kwordsDict[d[0]][d[1]] = {}
    if d[2] not in kwordsDict[d[0]][d[1]]:
        kwordsDict[d[0]][d[1]][d[2]] = {}
    kwordsDict[d[0]][d[1]][d[2]][d[3]] = d[4]

with open(kwordsRowLimitsPathJson, 'w') as outfile:
    json.dump(kwordsDict, outfile,indent=4)
