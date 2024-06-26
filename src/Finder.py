'''
 A class that loads and contains the data required for the reporting process
 from loading and processing the rules from the rule files to finding the files that should be scanned.
'''

import os

class DataFinder():

    def __init__(self, gitDirectory:str='.') -> None:
        self.gitDirectory    = gitDirectory
        self.VBADirectories  = self.findVbaDirectories()
        self.SourceCodeFiles = self.findFiles()
        self.RulesFiles      = self.findRulesFiles()

    def getSourceCodeFiles(self) -> list:
        return self.SourceCodeFiles
    def getVBADirectories(self) -> list:
        return self.VBADirectories
    def getRulesFiles(self) -> list:
        return self.RulesFiles

    def findFiles(self) -> list:
        listOfFiles = [directory+"/"+file for directory in self.VBADirectories 
                            for file in os.listdir(directory) 
                            if file.endswith(".bas") or 
                            file.endswith(".frm") or
                            file.endswith(".cls")]
        return listOfFiles

    def findVbaDirectories(self) -> list:
        vbaDirectories = [vbaDirectory for vbaDirectory,_,_ in os.walk(self.gitDirectory) if vbaDirectory.endswith("_vba")]
        if len(vbaDirectories) == 0:
            raise Exception("No vba source code files found.")
        return vbaDirectories

    def findRulesFiles(self) -> list:
        rulesFiles = [os.path.join(root,file) for root,dirs,files in os.walk('.') for file in files if file.startswith("Rules_")]
        if len(rulesFiles) == 0:
            raise Exception("No rules files found.")

        return rulesFiles
