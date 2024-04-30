# A file to start with the code scanning

# We require a function to load in the .bas,.cls, and .frm files

# Then we have maybe a class that does the checking provided a file?
# We could have a class the logs the errors and creates two logs of must have violations and should have violations?

import src.Finder
import src.Scanner
import src.Reporter
import sys
import os

for dir,subdir,file in os.walk("../"):
    print(dir)
    print(subdir)
    print(file)
    
print("Hello World")

with open("./output/testFile1.txt",'w'):
    pass
  
dirPath     = "./output"
  
loader      = Finder.DataFinder(dirPath)
scanner     = Scanner.DataScanner(loader,ruleDicts[0],ruleDicts[1],ruleDicts[2])
reporter    = Reporter.DataReporter(scanner.getViolations(),scanner.getRuleSeverity(),scanner.getRuleName())
  
reporter.generateUserReport()
reporter.generateReviewReport()
# reporter.generateUserReportLog()
reporter.generateReviewReportLog()
