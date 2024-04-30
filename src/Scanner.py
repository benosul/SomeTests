'''
A class that uses the data contained in a Dataloader Object to apply the rules to the content of the files
found by the Dataloader 

Output Formats:
rules:                  { have : [Rules], avoid : [Rules]}
ruleSeverity:           { Rule : Severity }
ruleNames:              { Rule : Name }
Mode == User:
    violations:         { File : { Severity : { Rule : [Linenumber] } } }
Mode == Reviewer:
    violations:         { Severity : { Rule : { File : [Linenumber] } } }

'''
import re
import os

class DataScanner():

    def __init__(self,data) -> None:
        self.rules              = {"have":[],"avoid":[]}
        self.ruleSeverity       = {}
        self.ruleNames          = {}
        self.violationsUser     = {}
        self.violationsReview   = {}
        self.data               = data
        self.parseRules()
        self.findViolations()

    def getRules(self): return self.rules
    def getViolations(self): return self.violationsUser, self.violationsReview
    def getRuleSeverity(self): return self.ruleSeverity
    def getRuleName(self): return self.ruleNames

    def parseRules(self,rulesFile=""):
        if rulesFile == "":
            rulesFile = self.data.getRulesFiles()
        self.rulesReader(rulesFile)   
    
    def rulesReader(self, fileList):
        if type(fileList)==list:
            for file in fileList:
                self.rulesReader(file)
        elif type(fileList)==str:
            filename = os.path.basename(fileList)
            if filename == "Rules_Have.txt":
                with open(fileList, "r") as ruleFiles:
                    for line in ruleFiles.readlines():
                        lineList = line.split(" ")
                        rule     = lineList[0].strip()
                        severity = lineList[1].strip()
                        name     = lineList[2].strip().replace("\\s"," ")
                        self.addToRulesDict(self.rules,"have",rule)
                        self.addToRulesDict(self.ruleSeverity, rule, severity)
                        self.addToRulesDict(self.ruleNames,rule,name)
            elif filename == "Rules_Avoid.txt":
                with open(fileList,"r") as ruleFiles:
                    for line in ruleFiles.readlines():
                        lineList = line.split(" ")
                        rule     = lineList[0].strip()
                        severity = lineList[1].strip()
                        name     = lineList[2].strip().replace("\\s"," ")
                        self.addToRulesDict(self.rules,"avoid",rule)
                        self.addToRulesDict(self.ruleSeverity, rule, severity)
                        self.addToRulesDict(self.ruleNames,rule,name)            

    def addToRulesDict(self,targetdict:dict,key,value):
        if targetdict == self.rules:
            if key not in targetdict:
                targetdict.update({key : []})
            targetdict[key] += [value]
        else:
            if key not in targetdict:
                targetdict[key] = value
            else: raise Exception("A one to one pairing was almost overwritten: " + key)

    def findViolations(self,files:list=[""]):
        if files == [""]:
            files = self.data.getSourceCodeFiles()
        for ruleType in self.rules:
            if self.rules[ruleType] != []: # if there are no rules in the rule type, we don't proceed
                for sourcefile in files:
                    with open(sourcefile,"r") as file:
                        content = file.readlines()
                    if ruleType=="have":
                        for rule in self.rules[ruleType]:
                            severity = self.ruleSeverity[rule]
                            if re.search(rule,str(content)) is None:
                                if not sourcefile in self.violationsUser:
                                    self.violationsUser.update({ sourcefile : {} })
                                if not severity in self.violationsUser[sourcefile]:
                                    self.violationsUser[sourcefile].update({severity : {} })
                                self.violationsUser[sourcefile][severity][rule] = True
                                if not severity in self.violationsReview:
                                    self.violationsReview.update({ severity : {} })
                                if not rule in self.violationsReview[severity]:
                                    self.violationsReview[severity].update({ rule : {}})
                                self.violationsReview[severity][rule][sourcefile] = True
                    elif ruleType=="avoid":
                        for rule in self.rules[ruleType]:
                            if not re.search(rule,str(content)) is None:
                                tempList = []
                                severity = self.ruleSeverity[rule]
                                for linenumber, line in enumerate(content):
                                    if re.search(rule,line):
                                        tempList.append(str(linenumber+1))
                                        if not sourcefile in self.violationsUser:
                                            self.violationsUser.update({sourcefile:{}})
                                        if not severity in self.violationsUser[sourcefile]:
                                            self.violationsUser[sourcefile].update({severity : {}})
                                        if not rule in self.violationsUser[sourcefile]:
                                            self.violationsUser[sourcefile][severity].update({rule:[]})
                                        if not severity in self.violationsReview:
                                            self.violationsReview.update({ severity : {} })
                                        if not rule in self.violationsReview[severity]:
                                            self.violationsReview[severity].update({ rule : {}})
                                        if not file in self.violationsReview[severity][rule]:
                                            self.violationsReview[severity][rule].update({ sourcefile : [] })
                                if tempList != []:
                                    self.violationsUser[sourcefile][severity][rule] = tempList
                                    self.violationsReview[severity][rule][sourcefile]= tempList

class LoadedScanner(DataScanner):

    def __init__(self,data,rules,ruleSeverity,ruleNames):
        self.rules              = rules
        self.ruleSeverity       = ruleSeverity
        self.ruleNames          = ruleNames
        self.violationsUser     = {}
        self.violationsReview   = {}
        self.data               = data
        self.findViolations()
