'''
 A class that loads and contains the data required for the reporting process
 from loading and processing the rules from the rule files to finding the files that should be scanned.
'''

import os

class DataFinder():

    def __init__(self, gitDirectory:str) -> None:
        self.gitDirectory    = gitDirectory
        self.VBADirectories  = self.findVbaDirectories()
        self.SourceCodeFiles = self.findFiles()
        self.RulesFiles      = self.findRulesFiles()

    def getSourceCodeFiles(self) -> list[str]:
        return self.SourceCodeFiles
    def getVBADirectories(self) -> list[str]:
        return self.VBADirectories
    def getRulesFiles(self):
        return self.RulesFiles

    def findFiles(self) -> list[str]:
        listOfFiles = [directory+"\\"+file for directory in self.VBADirectories 
                            for file in os.listdir(directory) 
                            if file.endswith(".bas") or 
                            file.endswith(".frm") or
                            file.endswith(".cls")]
        return listOfFiles

    def findVbaDirectories(self) -> list[str]:
        vbaDirectories = [vbaDirectory for vbaDirectory,_,_ in os.walk(self.gitDirectory) if vbaDirectory.endswith("_vba")]
        return vbaDirectories

    def findRulesFiles(self) -> list[str]:
        rulesFiles = [os.path.join(root,file) for root,dirs,files in os.walk(self.gitDirectory) for file in files if file.startswith("Rules_")]
        return rulesFiles
