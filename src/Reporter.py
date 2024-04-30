'''
A class that generates a report based off of the finding of the Scanner object
'''
# generiere 2 Reports UserReport / ReviewerReport oder Gesamtsicht / Usersicht 
import logging

class DataReporter():

    def __init__(self,violations,severity:dict,names:dict,dirPath) -> None:
        self.violationsUser     = violations[0]
        self.violationsReview   = violations[1]
        self.severity           = severity
        self.names              = names
        self.dirPath            = dirPath
        self.reportPath         = "Report.txt"
        

    def getViolations(self):
        return self.violationsUser,self.violationsReview
    def getSeverity(self):
        return self.severity
    def getNames(self):
        return self.names
    def getDirPath(self):
        return self.dirPath
    def getReportPath(self):
        return self.reportPath
    
    def generateReportPath(self):
        self.reportPath = self.dirPath + self.reportPath
        
    def generateUserReport(self):
        with open(self.dirPath + "/" + "User" + self.reportPath,"a") as file:
            for key1 in self.violationsUser:
                file.write("File:\t" + key1 + "\n")
                for key2 in self.violationsUser[key1]:
                    file.write("\tSeverity:\t" + key2 + "\n")
                    for key3 in self.violationsUser[key1][key2]:
                        file.write("\t\tRule:\t" + self.names[key3] + "\n")
                        lines= self.violationsUser[key1][key2][key3]
                        if lines != True:
                            file.write("\t\t\tLines: " + str(lines) + "\n")

    def generateReviewReport(self):
        with open(self.dirPath + "/" +"Review" + self.reportPath,"a") as file:
            for key1 in self.violationsReview:
                file.write("Severity:\t" + key1 + "\n")
                for key2 in self.violationsReview[key1]:
                    file.write("\tRule:\t" + self.names[key2] + "\n")
                    for key3 in self.violationsReview[key1][key2]:
                        file.write("\t\tFile:\t" + key3 + "\n")
                        lines= self.violationsReview[key1][key2][key3]
                        if lines != True:
                            file.write("\t\t\tLines: " + str(lines) + "\n")

    def generateUserReportLog(self):
        # Setting up logging:
        logging.basicConfig(filename= self.dirPath + "/UserReport.log",format='%(levelname)-8s: %(message)s')        
        for location in self.violationsUser:
            for severity in self.violationsUser[location]:
                for rule in self.violationsUser[location][severity]:
                    lines   = self.violationsUser[location][severity][rule]
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
    def generateReviewReportLog(self):
        # Setting up logging:
        logging.basicConfig(filename= self.dirPath + "/ReviewReport.log",format='%(levelname)-8s: %(message)s')
        for severity in self.violationsReview:
            for rule in self.violationsReview[severity]:
                for location in self.violationsReview[severity][rule]:
                    lines   = self.violationsReview[severity][rule][location]
                    if lines == True:
                        output = "Rule: " + self.names[rule] + " \n\t\t File: " + location
                    else:
                        output = "Rule: " + self.names[rule] + " \n\t\t File: " + location + "\n\t\t\tLines: " + str(lines)
                    if severity == "Low":
                        logging.info(output)
                    elif severity == "Medium":
                        logging.warning(output)
                    elif severity == "High":
                        logging.error(output)
                    else:
                        raise Exception("Invalid Severity Level found:" + severity)
                            
