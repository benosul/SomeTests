# A file to start with the code scanning

# We require a function to load in the .bas,.cls, and .frm files

# Then we have maybe a class that does the checking provided a file?
# We could have a class the logs the errors and creates two logs of must have violations and should have violations?

import src.Finder as Finder
import src.Scanner as Scanner
import src.Reporter as Reporter
import sys
import os

for dir,subdir,file in os.walk("../"):
    for correct in  ["../app","../Reports"]:
        if correct in dir: 
            print(dir)
            print(subdir)
            print(file)
    
print("Hello World")

with open("./output/testFile1.txt",'w'):
    pass
  
dirPath     = "Reports/output"
  
loader      = Finder.DataFinder(dirPath)
scanner     = Scanner.DataScanner(loader,ruleDicts[0],ruleDicts[1],ruleDicts[2])
reporter    = Reporter.DataReporter(scanner.getViolations(),scanner.getRuleSeverity(),scanner.getRuleName())
  
reporter.generateUserReport()
reporter.generateReviewReport()
# reporter.generateUserReportLog()
reporter.generateReviewReportLog()
