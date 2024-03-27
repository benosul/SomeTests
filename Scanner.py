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

    def __init__(self,data,reportType:str="User") -> None:
        self.rules        = {"have":[],"avoid":[]}
        self.ruleSeverity = {}
        self.ruleNames    = {}
        self.violations   = {}
        self.data         = data
        self.ReportType   = reportType # "User" or "Reviewer"
        self.parseRules()
        self.findViolations()

    def getRules(self) -> dict:
        return self.rules
    def getViolations(self):
        return self.violations
    def getRuleSeverity(self):
        return self.ruleSeverity
    def getRuleName(self):
        return self.ruleNames

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
                                if self.ReportType == "User":
                                    if not sourcefile in self.violations:
                                        self.violations.update({ sourcefile : {} })
                                    if not severity in self.violations[sourcefile]:
                                        self.violations[sourcefile].update({severity : {} })
                                    self.violations[sourcefile][severity][rule] = True
                                elif self.ReportType == "Reviewer":
                                    if not severity in self.violations:
                                        self.violations.update({ severity : {} })
                                    if not rule in self.violations[severity]:
                                        self.violations[severity].update({ rule : {}})
                                    self.violations[severity][rule][sourcefile] = True
                                else: raise Exception("Invalid Report type provided.")
                    elif ruleType=="avoid":
                        for rule in self.rules[ruleType]:
                            if not re.search(rule,str(content)) is None:
                                tempList = []
                                severity = self.ruleSeverity[rule]
                                for linenumber, line in enumerate(content):
                                    if re.search(rule,line):
                                        tempList.append(str(linenumber+1))
                                        if self.ReportType == "User":
                                            if not sourcefile in self.violations:
                                                self.violations.update({sourcefile:{}})
                                            if not severity in self.violations[sourcefile]:
                                                self.violations[sourcefile].update({severity : {}})
                                            if not rule in self.violations[sourcefile]:
                                                self.violations[sourcefile][severity].update({rule:[]})
                                        elif self.ReportType == "Reviewer":
                                            if not severity in self.violations:
                                                self.violations.update({ severity : {} })
                                            if not rule in self.violations[severity]:
                                                self.violations[severity].update({ rule : {}})
                                            if not file in self.violations[severity][rule]:
                                                self.violations[severity][rule].update({ sourcefile : [] })
                                        else: raise Exception("Invalid Report type provided.")
                                if tempList != []:
                                    if self.ReportType == "User":
                                        self.violations[sourcefile][severity][rule] = tempList
                                    elif self.ReportType == "Reviewer":
                                        self.violations[severity][rule][sourcefile]= tempList
