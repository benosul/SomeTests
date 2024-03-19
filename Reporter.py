'''
A class that generates a report based off of the finding of the Scanner object
'''
# generiere 2 Reports UserReport / ReviewerReport oder Gesamtsicht / Usersicht 
import os
from datetime import datetime
import logging

class DataReporter():

    def __init__(self,violations:dict,severity:dict,names:dict,mode:str="User",dirPath:str=".\\CodeScanning\\") -> None:
        self.violations = violations
        self.severity   = severity
        self.mode       = mode
        self.names       = names
        self.dirPath    = dirPath
        timestamp       = str(datetime.now())[:20].replace(".","").replace(":","-").replace(" ","_")
        self.reportPath = "\\report.txt"
        # Setting up logging:
        logging.basicConfig(filename="log_Report.log",format='%(levelname)-8s: %(message)s')

    def getViolations(self):
        return self.violations
    def getSeverity(self):
        return self.severity
    def getNames(self):
        return self.names
    def getMode(self):
        return self.mode
    def getDirPath(self):
        return self.dirPath
    def getReportPath(self):
        return self.reportPath
    
    def generateReportPath(self):
        self.reportPath = self.dirPath + self.reportPath
        
    def generateReport(self):
        self.generateReportPath()
        if not os.path.exists(self.reportPath):
            with open(self.reportPath,"w"):
                pass
        if self.mode == "User":
            with open(self.reportPath,"a") as file:
                for key1 in self.violations:
                    file.write("File:\t" + key1 + "\n")
                    for key2 in self.violations[key1]:
                        file.write("\tSeverity:\t" + key2 + "\n")
                        for key3 in self.violations[key1][key2]:
                            file.write("\t\tRule:\t" + self.names[key3] + "\n")
                            lines= self.violations[key1][key2][key3]
                            if lines != True:
                                file.write("\t\t\tLines: " + str(lines) + "\n")

        elif self.mode == "Reviewer":
            with open(self.reportPath,"a") as file:
                for key1 in self.violations:
                    file.write("Severity:\t" + key1 + "\n")
                    for key2 in self.violations[key1]:
                        file.write("\tRule:\t" + self.names[key2] + "\n")
                        for key3 in self.violations[key1][key2]:
                            file.write("\t\tFile:\t" + key3 + "\n")
                            lines= self.violations[key1][key2][key3]
                            if lines != True:
                                file.write("\t\t\tLines: " + str(lines) + "\n")

    def generateReportLogging(self):
        if self.mode == "User":          
            for location in self.violations:
                for severity in self.violations[location]:
                    for rule in self.violations[location][severity]:
                        lines   = self.violations[location][severity][rule]
                        if lines == True:
                            output = " File: " + location + "\n\t\t Rule: " + self.names[rule]
                        else:
                            output = " File: " + location + " \n\t\t Rule: " + self.names[rule] +"\n\t\t\tLines: " + str(lines)
                        if severity == "Low":
                            logging.info(output)
                        elif severity == "Medium":
                            logging.warning(output)
                        elif severity == "High":
                            logging.error(output)
                        else:
                            raise Exception("Invalid Severity Level found:" + severity)
        if self.mode == "Reviewer":
            for severity in self.violations:
                for rule in self.violations[severity]:
                    for location in self.violations[severity][rule]:
                        lines   = self.violations[severity][rule][location]
                        if lines == True:
                            output = "Rule: " + self.names[rule] + " \n\t\t Rule: " + location
                        else:
                            output = "Rule: " + self.names[rule] + " \n\t\t Rule: " + location + "\n\t\t\tLines: " + str(lines)
                        if severity == "Low":
                            logging.info(output)
                        elif severity == "Medium":
                            logging.warning(output)
                        elif severity == "High":
                            logging.error(output)
                        else:
                            raise Exception("Invalid Severity Level found:" + severity)
                            