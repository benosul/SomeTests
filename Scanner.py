
# A class that uses the data contained in a Dataloader Object to apply the rules to the content of the files
# found by the Dataloader 
import Finder

import re
import os

class DataScanner():

    def __init__(self,data:Finder.DataFinder) -> None:
        self.rules      = {"have":[],"avoid":[],"other":[]}
        self.violations = {}
        self.data       = data
        #self.rules = self.parseRules(data.getRulesFiles())
        #self.violations = self.findViolations(data.getSourceCodeFiles())

    def getRules(self) -> dict:
        return self.rules
    
    def getViolations(self):
        return self.violations

    def parseRules(self,rulesFile=""):
        if rulesFile == "":
            rulesFile = self.data.getRulesFiles()
        self.rulesReader(rulesFile)
        
    
    def rulesReader(self, fileList):
        rules = []
        if type(fileList)==list:
            for file in fileList:
                self.rulesReader(file)
        elif type(fileList)==str:
            filename = os.path.basename(fileList)
            if filename == "Rules_Have.txt":
                with open(fileList, "r") as ruleFiles:
                    rules = rules + [line.strip() for line in ruleFiles.readlines()]
                self.rules["have"] = self.rules["have"] + rules
            elif filename == "Rules_Avoid.txt":
                with open(fileList,"r") as ruleFiles:
                    self.rules["avoid"] = self.rules["avoid"] + [line.strip() for line in ruleFiles.readlines()]
            else:
                with open(fileList,"r") as ruleFiles:
                    self.rules["other"] = self.rules["other"] + [line.strip() for line in ruleFiles.readlines()]

    def findViolations(self,files:list[str]=[""]) -> dict:
        if files == [""]:
            files = self.data.getSourceCodeFiles()
        for ruleType in self.rules:
            if self.rules[ruleType] != []:
                for sourcefile in files:
                    with open(sourcefile,"r") as file:
                        content = file.readlines()
                    if ruleType=="have":
                        for rule in self.rules[ruleType]:
                            if re.search(rule,str(content)) is None:
                                print("++++++++++++++ VIOLATION FOUND ++++++++++++")
                                if not sourcefile in self.violations:
                                    self.violations.update({sourcefile:{}})
                                if not rule in self.violations[sourcefile]:
                                    self.violations[sourcefile].update({rule:[]})
                                self.violations[sourcefile][rule].append(-1)
                    else:
                        for rule in self.rules[ruleType]:
                            for linenumber, line in enumerate(content):
                                if re.search(rule,line):
                                    print("++++++++++++++ VIOLATION FOUND ++++++++++++")
                                    if not sourcefile in self.violations:
                                        self.violations.update({sourcefile:{}})
                                    if not rule in self.violations[sourcefile]:
                                        self.violations[sourcefile].update({rule:[]})
                                    self.violations[sourcefile][rule].append(linenumber)
        return self.violations