from typing import NamedTuple
from operator import attrgetter

class CurrentObject(NamedTuple):
        source: str
        body: str
        score: int

#creates list of all sources with source text and abstract text        
def createArray(filename): 
        list = [] 
        with open(filename, 'r', encoding = "utf-8") as file:        
                lines = file.read()
                splitter = lines.split("\n\n")
                for count, text in enumerate(splitter, 1):
                        splitter2= text.split("\t")
                        for count2, text in enumerate(splitter2, 1):
                                if type(text) == str:
                                        text = text.encode(encoding = 'UTF-8',errors = 'ignore')
                                else:
                                        text = str(text)
                                if count2 == 1:
                                        source = text
                                else:
                                        body = text
                        tempObj = CurrentObject(source, body, 0)
                        list.append(tempObj)
        return list
       
#creates list of keywords in keywords file       
def createKeyList(keywordsFile):
        keywordsList = []
        with open(keywordsFile, 'r', encoding = "utf-8") as file:        
                lines = file.read()
                splitter = lines.split("\n")
                for count, text in enumerate(splitter, 1):
                        keywordsList.append(text)
                return keywordsList

#creates list of whitelisted words in whitelist file
def createWhitelist(whitelistFile):
        whitelistList = []
        with open(whitelistFile, 'r', encoding = "utf-8") as file:        
                lines = file.read()
                splitter = lines.split("\n")
                for count, text in enumerate(splitter, 1):
                        whitelistList.append(text)
                return whitelistList


#creates a score for each source based on prevalence of keywords and whitelist words
def compare(list, keywords, whitelist):
        for obj in list:
                totalScore = 0
                tempObj = obj
                whitelistCount = 0
                for words in keywords:
                        totalScore += (str(obj.source.lower()).count(words.lower()) + (str(obj.body.lower()).count(words.lower())))
                for words in whitelist:
                        whitelistCount += (str(obj.source.lower()).count(words.lower()) + (str(obj.body.lower()).count(words.lower())))
                                
                totalScore -= whitelistCount

                obj = obj._replace(score = totalScore)
                list[list.index(tempObj)] = obj

#removes scores below a value and sorts sources by keyword score from greatest to least 
def sortByScore(list, sortingValueMinimum):
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
        with open(resultsFile, 'w+', encoding = "utf-8") as file:
                for obj in list:
                        print([obj])
                        scorePrintStatement = "Score: " + str(obj.score) + "\n"
                        sourcePrintStatement = str(obj.source)
                        bodyPrintStatement = str(obj.body) + "\n\n\n"
                        file.write(scorePrintStatement)
                        file.write(sourcePrintStatement)
                        file.write("\n")
                        file.write(bodyPrintStatement)
                        
def main():
        filename = "Sources.txt"               #location of sources file
        keywordsFile = "Keywords.txt"     #location of keywords file
        whitelistFile = "Whitelist.txt"          #location of whitelist words file
        sortingValueMinimum = 2             #minimum number of keywords found in file to become a result
        
        
        editedList = createArray(filename)
        keywordList = createKeyList(keywordsFile)
        whitelistedList = createWhitelist(whitelistFile)
        
        compare(editedList, keywordList, whitelistedList)
        
        sortedList = sortByScore(editedList, sortingValueMinimum)
        writeToFile(sortedList, filename+".result.txt")
main()