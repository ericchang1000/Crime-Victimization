from typing import NamedTuple
from operator import attrgetter
import xlsxwriter
import pandas


class CurrentObject(NamedTuple):
        incidentNum: str
        narrative: str
        score: int


#opens file and creates tuple object for each row with incident number and narrative
def createList(readFile):
        print("Opening File...")
        objectList = []
        data = pandas.read_excel(readFile)
        data.drop([data.columns[0], data.columns[1] ,data.columns[2]], axis=1, inplace=True)
        
        incidentNumber = data[data.columns[0]].tolist()
        narrative = data[data.columns[1]].tolist()
        
        newNumList = []
        newNarList = []
        totalCount = 0
       
        for num in incidentNumber:
                if (totalCount < 1):
                        newNumList.append(num)
                        newNarList.append(str(narrative[totalCount]))
                        # print(newNumList.index(num))
                else:
                        if (num == incidentNumber[totalCount-1]):
                                newNarList[newNumList.index(num)] += str(narrative[totalCount])
                        else:
                                newNumList.append(num)
                                newNarList.append(str(narrative[totalCount]))
                
                totalCount += 1


        for count in range(len(newNumList)):
                Object = CurrentObject(newNumList[count], str(newNarList[count]), 0)
                objectList.append(Object)
                print(Object)
#         for count in range(len(incidentNumber)):
#                 Object = CurrentObject(incidentNumber[count], str(narrative[count]), 0)
#                 objectList.append(Object)
#                 print(Object)

#         print(len(newNumList))
#         print(len(newNarList))
        return objectList
        
        
#creates keyword list
def createKeyList(keywordsFile):
        print("Creating Keywords List...")
        keywordsList = []
        with open(keywordsFile, 'r', encoding = "utf-8") as file:        
                lines = file.read()
                splitter = lines.split("\n")
                for count, text in enumerate(splitter, 1):
                        keywordsList.append(text)
                return keywordsList
                
                
#compares list and keywords
def compare(list, keywords):
        print("Comparing words...")
        for obj in list:
                totalScore = 0
                tempObj = obj
                for words in keywords:
                        totalScore += (str(obj.narrative.lower()).count(words.lower()))                
                obj = obj._replace(score = totalScore)
                list[list.index(tempObj)] = obj


#removes scores below a value and sorts sources by keyword score from greatest to least 
def sortByScore(list, sortingValueMinimum):
        print("Removing extras...")
        newList = sorted(list, reverse = True, key=attrgetter('score'))
        for obj in newList:
                if (obj.score < sortingValueMinimum):
                        indexValue = newList.index(obj)
                        break
        del newList[indexValue: ]
 
        #for obj in newList:
                #print(obj)
        return newList
        
        
#writes sources to file if over a certain score by order of prevalence
def writeToFile(list, resultsFile):
        print("Writing to File...")
        workbook = xlsxwriter.Workbook(resultsFile)
        worksheet = workbook.add_worksheet()
        
        tempCol = 0
        categories = ["Number", "Narrative"]
        for item in categories:
                worksheet.write(0, tempCol, item)
                tempCol += 1
                
        row = 1
        for obj in list:
                # print(obj)
                col = 0
                callValue = str(obj.incidentNum)
                narrativeValue = str(obj.narrative)
                content = [callValue, narrativeValue]
                for currentText in content:
                        worksheet.write(row, col, currentText)
                        col += 1
                row += 1
        workbook.close()


#Main
# def main(filename):
#         readFile = filename               #location of sources file
def main():
        readFile = "Sources/2016narrative.xlsx"
        keywordsFile = "keywords.txt"                      #location of keywords file
        sortingValueMinimum = 2                          #minimum number of keywords found in file to become a result
        
        list = createList(readFile)
        keyList = createKeyList(keywordsFile)
       
        compare(list, keyList)
         
        finalList = sortByScore(list, sortingValueMinimum)
        writeToFile(finalList, readFile+".result.xlsx")
        print("Script has finished")
        
main()
# def init():
#         filenameList = ["August2017.xlsx", "December2017.xlsx", "July2017.xlsx", "June2017.xlsx", "May2017.xlsx", "November2017.xlsx", "October2017.xlsx", "September2017.xlsx"]
#         for item in filenameList:
#                 file = "Sources/" + item
#                 main(file)
# init()