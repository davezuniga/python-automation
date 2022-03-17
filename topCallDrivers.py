"""
Created on Wed April 16 2020
@author: David Zuniga Rojas
"""
from nltk.tokenize import sent_tokenize, word_tokenize
from difflib import SequenceMatcher
from openpyxl import load_workbook
from nltk import pos_tag
import matplotlib.pyplot as plt
import pandas as pd
import nltk

##Excel format must be: Ticket column A (0) Short description column E (4)
df = pd.read_excel("")

excelPath = ""

apps = [""]

knowIssues = [""]

skipWords = [""]

issuesLen = len(apps)
dfClassified = pd.DataFrame()
dfWords = pd.DataFrame()  
dfLen = len(df.index)
counter = dfLen
i = 0
x = 0
keywords = []
finalData = { } #Dictionary
itemInList = False


def keywordsToString():
    output = ""
    
    for item in keywords:
        output += item.capitalize()
        output += " "
        
    return output


for i in range(0, dfLen):
    value = df.iloc[i][4]
    shortDescription =  str(value)
    
    print(counter)
    counter = counter - 1
    
    words = word_tokenize(shortDescription.lower())
    
    for itemList in apps:
        for item in words:
            similarity = SequenceMatcher(None, itemList.lower(), item.lower()).ratio()
            
            if similarity >= 0.8:
                keywords.append(itemList)   
                itemInList = True
    
    
    for itemIssues in knowIssues:
        for item1 in words:
            similarity = SequenceMatcher(None, itemIssues.lower(), item1.lower()).ratio()
            
            if similarity >= 0.76:
                keywords.append(itemIssues)   
                itemInList = True     
    
    #If the word is not on know issues        
    if itemInList == False:
        tokensTag = pos_tag(words)

        for x in range(0, len(tokensTag)):
            if "NN" in tokensTag[x] or "NNP" in tokensTag[x]:
                strs = " ".join(tokensTag[x])
                final = word_tokenize(strs) 
                
                for item in final:
                    if item != "NN" and item != "NNP" and item.lower() not in skipWords: 
                        dfWords = dfWords.append({"Keywords": item.capitalize(), "Ticket": df.iloc[i][0],
                         "Short description": shortDescription.capitalize()}, ignore_index= True)
                        keywords.append(item)

        dfClassified = dfClassified.append({"Keywords": keywordsToString(), "Ticket": df.iloc[i][0],
                         "Short description": shortDescription}, ignore_index= True)
     
    else:
        dfClassified = dfClassified.append({"Keywords": keywordsToString(), "Ticket": df.iloc[i][0],
                         "Short description": shortDescription.capitalize()}, ignore_index= True)
        dfWords = dfWords.append({"Keywords": keywordsToString(), "Ticket": df.iloc[i][0],
                         "Short description": shortDescription.capitalize()}, ignore_index= True)
    
    keywords = []
    itemInList = False

final = dfClassified[["Ticket", "Keywords"]]
    
#Set index
dfWords = dfWords.set_index("Ticket")
dfClassified = dfClassified.set_index("Ticket")

#Load excel document
writer = pd.ExcelWriter(excelPath, engine='openpyxl')

try:
    writer.book = load_workbook(excelPath)
    writer.sheets = dict(
        (ws.title, ws) for ws in writer.book.worksheets)
except IOError:
    print("Error while wrting information to file")
    pass

#Overwrite only the especify sheets
dfWords.to_excel(writer, sheet_name = "RawDataByWord")
dfClassified.to_excel(writer, sheet_name = "RawDataByTicket")

#Saves document
writer.save()


showFinal = final.set_index("Ticket")

showFinal.head(10)
showFinal.hist(column="Keywords", bins=len(showFinal))
plt.show()
